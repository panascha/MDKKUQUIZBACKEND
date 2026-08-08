/*
   =========================================
   ส่วนที่ 4: AI Expert & API Quota
   AI_Config = Gemini key pool โครงแบบ IntelSphere_Keys (คอลัมน์ <model>_Remaining ต่อโมเดล)
   AI_Models = ทะเบียนโมเดล free tier — เพิ่ม/ปิดโมเดลได้จากชีตโดยตรง ไม่ต้อง deploy ใหม่
   (AI Studio แยกโควต้า RPD ต่อโมเดล — นับรวมทั้ง key เดียวแบบเดิมไม่ได้แล้ว)
   =========================================
*/

var AI_CONFIG_SHEET_NAME = "AI_Config";
var AI_MODELS_SHEET_NAME = "AI_Models";
var AI_CONFIG_FIXED_HEADERS = ["API_Key", "Donor_Name", "Status", "Last_Used", "Last_Reset_Date"];

// ค่าตั้งต้นทะเบียนโมเดล ตามหน้า Rate Limit free tier ของ AI Studio (RPD ต่อโมเดล ต่อ key)
// ทุกตัว = ID ที่ยืนยันจาก models.list จริง (discoverGeminiModels) — gemini-3-flash/gemini-3.1-pro ถูกตัดออก
// เพราะไม่มีใน list (มีแต่ -preview) เป็น seed ปลอม; priority คงเลขเดิมของตัวที่รอด (ช่องว่างไม่เป็นไร)
var AI_MODELS_DEFAULTS = [
  // [Model, RPD_Limit, Priority, Status, Notes]
  ["gemini-3.5-flash",      20,  1, "Active",   "text-out หลัก"],
  ["gemini-2.5-flash",      20,  3, "Active",   ""],
  ["gemini-3.1-flash-lite", 500, 4, "Active",   "RPD สูงสุดใน free tier"],
  ["gemini-2.5-flash-lite", 20,  5, "Active",   ""],
  ["gemini-2.5-pro",        0,  91, "Disabled", "free tier RPD = 0"]
];

// Retry helper — เอกสารใหญ่ เจอ "บริการ สเปรดชีต หมดเวลา" เป็นพักๆ ระหว่าง write ติดกัน
function aiSheetRetry_(fn) {
  var lastErr;
  for (var a = 0; a < 3; a++) {
    try { return fn(); } catch (e) {
      lastErr = e;
      // งบ execution ใกล้หมด → เลิก retry ทันที (นอนรอ+ลองใหม่ = ชนเพดาน 6 นาทีเปล่าๆ)
      if (execRemainingMs_() < 20000) break;
      Utilities.sleep(1500 * (a + 1));
    }
  }
  throw lastErr;
}

// อ่านทะเบียนโมเดลจากชีต AI_Models (สร้าง+seed อัตโนมัติถ้ายังไม่มี "หรือว่างเปล่า") — cache ต่อ 1 execution
var _aiModelRegistryCache = null;
function getAIModelRegistry_(ss) {
  if (_aiModelRegistryCache) return _aiModelRegistryCache;
  ss = ss || SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(AI_MODELS_SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(AI_MODELS_SHEET_NAME);
  // seed เมื่อชีตว่าง — รวมกรณีชีตถูกสร้างค้างไว้จากรอบที่ write fail (ไม่ใช่แค่กรณีชีตหาย)
  if (!String(sheet.getRange(1, 1).getValue() || "").trim()) {
    aiSheetRetry_(function() {
      sheet.getRange(1, 1, 1, 5).setValues([["Model", "RPD_Limit", "Priority", "Status", "Notes"]])
        .setFontWeight("bold").setBackground("#e6f7ff");
      sheet.setFrozenRows(1);
      sheet.getRange(2, 1, AI_MODELS_DEFAULTS.length, 5).setValues(AI_MODELS_DEFAULTS);
      SpreadsheetApp.flush();
    });
  }
  var data = sheet.getDataRange().getValues();
  var models = [];
  for (var i = 1; i < data.length; i++) {
    var name = String(data[i][0] || "").trim();
    if (!name) continue;
    var limit = parseInt(data[i][1], 10) || 0;
    models.push({
      model: name,
      limit: limit,
      priority: parseInt(data[i][2], 10) || 999,
      active: String(data[i][3]).trim() === "Active" && limit > 0
    });
  }
  models.sort(function(a, b) { return a.priority - b.priority; });
  _aiModelRegistryCache = models;
  return models;
}

// เปิดชีต AI_Config โครงใหม่ (สร้าง/migrate จากโครงเดิม/เติมคอลัมน์โมเดลที่ขาด อัตโนมัติ)
function getAIConfigSheet_(ss) {
  ss = ss || SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(AI_CONFIG_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(AI_CONFIG_SHEET_NAME);
    sheet.getRange(1, 1, 1, AI_CONFIG_FIXED_HEADERS.length).setValues([AI_CONFIG_FIXED_HEADERS])
      .setFontWeight("bold").setBackground("#e6f7ff");
    sheet.setFrozenRows(1);
  }
  sheet = migrateAIConfigLegacy_(ss, sheet);
  ensureAIConfigModelColumns_(sheet, getAIModelRegistry_(ss));
  return sheet;
}

// โครงเดิม: A=API_Key B=Provider C=Model D=Daily_Limit E=Usage_Count F=Last_Used G=Status
// Rebuild = ลบชีตเดิมแล้วสร้างใหม่ทั้งชีต — sheet.clear() ไม่พอ เพราะ Sheets Table structure
// ของชีตเดิมรอด clear แล้วเขียน header ใหม่ไม่ติด (กลายเป็น "Column 6"/"Column 7")
function migrateAIConfigLegacy_(ss, sheet) {
  var headers = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 1)).getValues()[0];
  var isLegacy = headers.indexOf("Last_Reset_Date") < 0;
  var hasJunk = headers.some(function(h) { return /^Column \d+$/.test(String(h)); });
  if (!isLegacy && !hasJunk) return sheet; // โครงใหม่สมบูรณ์แล้ว

  var data = sheet.getDataRange().getValues();
  var todayStr = Utilities.formatDate(new Date(), "Asia/Bangkok", "yyyy-MM-dd");
  var rows = [];
  for (var i = 1; i < data.length; i++) {
    var key = String(data[i][0] || "").trim();
    if (!key) continue;
    if (isLegacy) {
      // Exhausted เดิมนับรวมราย key — โครงใหม่นับต่อโมเดล จึงปลุกกลับเป็น Active; Disabled คงไว้
      var status = (String(data[i][6] || "").trim() === "Disabled") ? "Disabled" : "Active";
      rows.push([key, "", status, data[i][5] || "", todayStr]);
    } else {
      // โครงใหม่แต่มีคอลัมน์ junk (migration รอบก่อน fail กลางทาง) — คงค่า 5 คอลัมน์แรกไว้
      var st = (String(data[i][2] || "").trim() === "Disabled") ? "Disabled" : "Active";
      rows.push([key, data[i][1] || "", st, data[i][3] || "", todayStr]);
    }
  }

  // แยก step + retry — เคยเจอ timeout กลางทางทำให้ key หาย (ลบแล้วเขียนกลับไม่ทัน)
  aiSheetRetry_(function() {
    var old = ss.getSheetByName(AI_CONFIG_SHEET_NAME);
    if (old) ss.deleteSheet(old);
    SpreadsheetApp.flush();
  });
  aiSheetRetry_(function() {
    if (!ss.getSheetByName(AI_CONFIG_SHEET_NAME)) ss.insertSheet(AI_CONFIG_SHEET_NAME);
  });
  sheet = ss.getSheetByName(AI_CONFIG_SHEET_NAME);
  aiSheetRetry_(function() {
    sheet.getRange(1, 1, 1, AI_CONFIG_FIXED_HEADERS.length).setValues([AI_CONFIG_FIXED_HEADERS])
      .setFontWeight("bold").setBackground("#e6f7ff");
    sheet.setFrozenRows(1);
    if (rows.length) sheet.getRange(2, 1, rows.length, AI_CONFIG_FIXED_HEADERS.length).setValues(rows);
    SpreadsheetApp.flush();
  });
  return sheet;
}

// กู้ key จาก revision history ของ spreadsheet (กรณี migration ทำ key หาย)
// GET ?action=recoverAIConfigKeys&before=<ISO> — เลือก revision ล่าสุดที่เก่ากว่า cutoff
function recoverAIConfigKeys(beforeIso) {
  function out(obj) {
    return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
  }
  var token = ScriptApp.getOAuthToken();
  var listResp = UrlFetchApp.fetch(
    "https://www.googleapis.com/drive/v2/files/" + SHEET_ID + "/revisions?maxResults=1000",
    { headers: { Authorization: "Bearer " + token }, muteHttpExceptions: true });
  if (listResp.getResponseCode() !== 200) {
    return out({ result: 'error', step: 'listRevisions', code: listResp.getResponseCode(),
                 body: String(listResp.getContentText()).slice(0, 300) });
  }
  var items = JSON.parse(listResp.getContentText()).items || [];
  var cutoff = beforeIso ? new Date(beforeIso).getTime() : Date.now();
  var best = null;
  for (var i = 0; i < items.length; i++) {
    var t = new Date(items[i].modifiedDate).getTime();
    if (t < cutoff && (!best || t > new Date(best.modifiedDate).getTime())) best = items[i];
  }
  if (!best) return out({ result: 'error', message: 'no revision before cutoff', revisions: items.length });
  var xlsxUrl = best.exportLinks &&
    best.exportLinks["application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"];
  if (!xlsxUrl) return out({ result: 'error', message: 'no xlsx exportLink', revisionId: best.id });

  var blob = UrlFetchApp.fetch(xlsxUrl, { headers: { Authorization: "Bearer " + token } }).getBlob();
  var tempMeta = Drive.Files.create(
    { name: "TMP_AI_Config_recovery", mimeType: "application/vnd.google-apps.spreadsheet" }, blob);
  var keys = [];
  try {
    var tmp = SpreadsheetApp.openById(tempMeta.id).getSheetByName(AI_CONFIG_SHEET_NAME);
    if (!tmp) return out({ result: 'error', message: 'AI_Config tab not in revision', revisionId: best.id });
    var vals = tmp.getDataRange().getValues();
    var skipped = [];
    for (var r = 1; r < vals.length; r++) {
      var k = String(vals[r][0] || "").trim();
      if (k.indexOf("AIza") === 0) keys.push(k); // Gemini key ขึ้นต้น AIza เสมอ (กัน header/ขยะ)
      else if (k) skipped.push(k.slice(0, 6) + "..." + k.slice(-3) + " (len " + k.length + ", B=" + String(vals[r][1] || "") + ")");
    }
    keys.skippedInfo = skipped;
  } finally {
    try { DriveApp.getFileById(tempMeta.id).setTrashed(true); } catch (e) {}
  }

  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = getAIConfigSheet_(ss);
  var models = getAIModelRegistry_(ss);
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var existing = {};
  for (var x = 1; x < data.length; x++) existing[String(data[x][0]).trim()] = true;
  var todayStr = Utilities.formatDate(new Date(), "Asia/Bangkok", "yyyy-MM-dd");
  var added = 0;
  keys.forEach(function(k) {
    if (existing[k]) return;
    var rowValues = new Array(headers.length).fill("");
    function setColV(name, val) { var c = headers.indexOf(name); if (c >= 0) rowValues[c] = val; }
    setColV("API_Key", k);
    setColV("Donor_Name", "recovered");
    setColV("Status", "Active");
    setColV("Last_Used", new Date());
    setColV("Last_Reset_Date", todayStr);
    models.forEach(function(m) { setColV(m.model + "_Remaining", m.limit); });
    aiSheetRetry_(function() { sheet.appendRow(rowValues); });
    added++;
  });
  return out({ result: 'success', revisionId: best.id, revisionDate: best.modifiedDate,
               keysFound: keys.length, keysAdded: added, keyRows: sheet.getLastRow() - 1,
               skippedRows: keys.skippedInfo || [] });
}

// เติมคอลัมน์ <model>_Remaining ที่ยังไม่มี — แถว key เดิม prefill โควต้าเต็มของโมเดลนั้น
function ensureAIConfigModelColumns_(sheet, models) {
  var headers = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 1)).getValues()[0];
  var lastRow = sheet.getLastRow();
  for (var i = 0; i < models.length; i++) {
    var h = models[i].model + "_Remaining";
    if (headers.indexOf(h) >= 0) continue;
    var col = sheet.getLastColumn() + 1;
    if (col > sheet.getMaxColumns()) sheet.insertColumnsAfter(sheet.getMaxColumns(), 1);
    sheet.getRange(1, col).setValue(h).setFontWeight("bold").setBackground("#e6f7ff");
    if (lastRow > 1) {
      var fill = [];
      for (var r = 2; r <= lastRow; r++) fill.push([models[i].limit]);
      sheet.getRange(2, col, fill.length, 1).setValues(fill);
    }
    headers.push(h);
  }
  return headers;
}

/**
 * เลือก (key, model) ที่ยังมีโควต้า — per-model RPD ตามทะเบียน AI_Models
 * คืน shape เดิม {key, model, index, usage, limit} + {remaining, fallbackModels}
 * preferredModel (optional): ใช้โมเดลนี้ก่อนถ้ายังมีโควต้า ไม่งั้นไล่ตาม Priority
 * reserveCount (optional, default 0): จอง key ที่มีโควต้าคงเหลือรวมมากสุด N อันดับแรกไว้ให้ผู้ใช้สาธารณะ
 *   — ผู้เรียกที่เป็น owner (agentQuery) เท่านั้นที่ส่งค่า > 0; converter/นิสิตใช้ default 0 = ใช้ได้ทุก key
 *   ถ้า active key มีไม่เกิน reserveCount → คืน null (ตั้งใจ: shared pool มาก่อน owner)
 */
function getAvailableAIKey(provider, preferredModel, reserveCount, avoidModels) {
  if (provider && provider !== "Gemini") return null; // pool นี้มีแต่ Gemini (IntelSphere แยกชีตของตัวเอง)
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = getAIConfigSheet_(ss);
  var models = getAIModelRegistry_(ss);
  var activeModels = models.filter(function(m) { return m.active; });
  if (activeModels.length === 0) return null;

  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var colStatus = headers.indexOf("Status");
  var colLastReset = headers.indexOf("Last_Reset_Date");
  var tz = "Asia/Bangkok"; // อย่าใช้ timezone ของ script (อาจเป็น UTC — reset ช้า 7 ชม.)
  var todayStr = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd");

  var candidates = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!String(row[0] || "").trim()) continue;
    if (row[colStatus] !== "Active") continue;

    // Daily reset ต่อแถว: refill ทุกคอลัมน์โมเดลกลับเป็นโควต้าเต็มเมื่อขึ้นวันใหม่
    var lastResetStr = row[colLastReset]
      ? Utilities.formatDate(new Date(row[colLastReset]), tz, "yyyy-MM-dd") : "";
    if (lastResetStr !== todayStr) {
      for (var m = 0; m < models.length; m++) {
        var rc = headers.indexOf(models[m].model + "_Remaining");
        if (rc >= 0) sheet.getRange(i + 1, rc + 1).setValue(models[m].limit);
      }
      sheet.getRange(i + 1, colLastReset + 1).setValue(todayStr);
      SpreadsheetApp.flush();
      data = sheet.getDataRange().getValues();
      row = data[i];
    }

    for (var j = 0; j < activeModels.length; j++) {
      var remCol = headers.indexOf(activeModels[j].model + "_Remaining");
      if (remCol < 0) continue;
      var remaining = Number(row[remCol]);
      if (isNaN(remaining) || remaining <= 0) continue;
      if (isModelCoolingDown_(row[0], activeModels[j].model)) continue; // ข้าม (key,model) ที่กำลัง RPM cooldown
      if (avoidModels && avoidModels[activeModels[j].model]) continue;   // ข้ามโมเดลที่เพิ่ง 429/5xx ใน call นี้ → cascade ไป priority ถัดไป
      candidates.push({
        key: row[0], model: activeModels[j].model, index: i + 1,
        usage: activeModels[j].limit - remaining, limit: activeModels[j].limit,
        remaining: remaining, priority: activeModels[j].priority
      });
    }
  }
  if (candidates.length === 0) return null;

  // จอง key ไว้ให้ผู้ใช้สาธารณะ (แบบเดียวกับ AGENT_QUERY_KEY_RESERVE_COUNT ของ IntelSphere_Keys):
  // รวม remaining ทุกโมเดลต่อ key → เรียงมาก→น้อย → ตัด top-N ออกจาก candidates
  if (reserveCount > 0) {
    var totalsByKey = {};
    for (var c = 0; c < candidates.length; c++) {
      var kid = String(candidates[c].key);
      totalsByKey[kid] = (totalsByKey[kid] || 0) + candidates[c].remaining;
    }
    var ranked = Object.keys(totalsByKey).sort(function(a, b) {
      return totalsByKey[b] - totalsByKey[a];
    });
    var reserved = {};
    for (var rr = 0; rr < reserveCount && rr < ranked.length; rr++) reserved[ranked[rr]] = true;
    candidates = candidates.filter(function(x) { return !reserved[String(x.key)]; });
    if (candidates.length === 0) return null; // เหลือแต่ key ที่จองไว้ — owner ไม่แตะ
  }

  // เคารพ preferredModel ถ้ายังมีโควต้า ไม่งั้นเอาโมเดล priority ดีสุดที่เหลือโควต้า
  var pool = null;
  if (preferredModel) {
    pool = candidates.filter(function(c) { return c.model === preferredModel; });
    if (pool.length === 0) pool = null;
  }
  if (!pool) {
    var bestPriority = Math.min.apply(null, candidates.map(function(c) { return c.priority; }));
    pool = candidates.filter(function(c) { return c.priority === bestPriority; });
  }

  // กระจายโหลดระหว่าง key: weighted-random ตาม remaining (แบบ IntelSphere_Keys)
  var total = pool.reduce(function(s, c) { return s + c.remaining; }, 0);
  var dart = Math.random() * total;
  var picked = pool[pool.length - 1];
  for (var k = 0; k < pool.length; k++) {
    dart -= pool[k].remaining;
    if (dart <= 0) { picked = pool[k]; break; }
  }
  // fallback chain สำหรับ converter: โมเดล Active เรียงตาม Priority จากทะเบียน
  picked.fallbackModels = activeModels.map(function(m) { return m.model; });
  return picked;
}

/**
 * หักโควต้าหลังเรียก AI สำเร็จ — ลด <usedModel>_Remaining ของแถว key นั้นลง 1
 */
function updateAIUsage(apiKeyInfo, usedModel) {
  var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(AI_CONFIG_SHEET_NAME);
  if (!sheet) return;
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var remCol = headers.indexOf((usedModel || apiKeyInfo.model) + "_Remaining");
  if (remCol >= 0) {
    var cell = sheet.getRange(apiKeyInfo.index, remCol + 1);
    var remaining = Number(cell.getValue());
    if (isNaN(remaining)) remaining = 0;
    cell.setValue(Math.max(0, remaining - 1));
  }
  var colLastUsed = headers.indexOf("Last_Used");
  if (colLastUsed >= 0) sheet.getRange(apiKeyInfo.index, colLastUsed + 1).setValue(new Date());
}

// โดน 429 จริงจาก Google — ตัดโควต้าโมเดลนั้นของ key นี้เป็น 0 กันเรียกซ้ำทั้งวัน
function markModelExhausted_(apiKeyInfo, model) {
  try {
    var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(AI_CONFIG_SHEET_NAME);
    if (!sheet) return;
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var remCol = headers.indexOf(model + "_Remaining");
    if (remCol >= 0) { sheet.getRange(apiKeyInfo.index, remCol + 1).setValue(0); SpreadsheetApp.flush(); } // flush: ให้ getAvailableAIKey รอบถัดไปเห็นค่า 0
  } catch (e) { console.warn("markModelExhausted_ failed: " + e.message); }
}

// Diagnostic อ่านอย่างเดียว — GET ?action=aiConfigStatus (key ถูก mask, endpoint สาธารณะ)
function getAIConfigStatus() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var out = { result: 'success' };
  var cfg = ss.getSheetByName(AI_CONFIG_SHEET_NAME);
  if (!cfg) { out.aiConfig = null; }
  else {
    var data = cfg.getDataRange().getValues();
    out.aiConfig = {
      headers: data[0],
      rows: data.slice(1).filter(function(r) { return String(r[0] || "").trim(); }).map(function(r) {
        var k = String(r[0]);
        return [k.slice(0, 8) + "..." + k.slice(-4)].concat(r.slice(1));
      })
    };
  }
  var mdl = ss.getSheetByName(AI_MODELS_SHEET_NAME);
  out.aiModels = mdl ? mdl.getDataRange().getValues() : null;
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

// Public read-only — GET ?action=getAIModels: ทะเบียนโมเดลสำหรับ admin panel (P2-Q6)
// คืน [{model, rpd, priority, status, notes}] ตรงจากชีต AI_Models (ไม่ sensitive → public เหมือน aiConfigStatus)
function getAIModels() {
  function out(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(AI_MODELS_SHEET_NAME);
  if (!sheet) return out({ result: 'error', message: 'AI_Models not found' });
  var data = sheet.getDataRange().getValues();
  var rows = [];
  for (var i = 1; i < data.length; i++) {
    var name = String(data[i][0] || "").trim();
    if (!name) continue;
    rows.push({ model: name, rpd: data[i][1], priority: data[i][2],
                status: String(data[i][3] || "").trim(), notes: String(data[i][4] || "") });
  }
  return out({ result: 'success', models: rows });
}

// Admin write — POST action=setModelRpd {model, rpd?, priority?} (P2-Q1/Q5: RPD = human go-live gate; P2-Q7: priority override)
// เขียนเฉพาะ field ที่ส่งมา (อย่างน้อย 1). rpd = non-negative int; priority = int (ติดลบได้ = flagship min-1)
// reset registry cache + ensureAIConfigModelColumns_ + backfill _Remaining=rpd ทุก key Active → serve วันนี้เลย
// auth ทำที่ doPost (mirror getFeedback: sessionToken admin หรือ username+adminPass) — ฟังก์ชันนี้ถือว่า authed แล้ว
function setModelRpd(model, rpd, priority) {
  function out(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
  var name = String(model || "").trim();
  if (!name) return out({ result: 'error', message: 'missing model' });
  var hasRpd = (rpd !== undefined && rpd !== null && String(rpd).trim() !== "");
  var hasPrio = (priority !== undefined && priority !== null && String(priority).trim() !== "");
  if (!hasRpd && !hasPrio) return out({ result: 'error', message: 'nothing to set (need rpd or priority)' });
  var rpdN, prioN;
  if (hasRpd) {
    rpdN = Number(rpd);
    if (!isFinite(rpdN) || rpdN < 0 || Math.floor(rpdN) !== rpdN) return out({ result: 'error', message: 'rpd must be a non-negative integer' });
  }
  if (hasPrio) {
    prioN = Number(priority);
    if (!isFinite(prioN) || Math.floor(prioN) !== prioN) return out({ result: 'error', message: 'priority must be an integer' });
  }
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(AI_MODELS_SHEET_NAME);
  if (!sheet) return out({ result: 'error', message: 'AI_Models not found' });
  var data = sheet.getDataRange().getValues();
  var row = -1;
  for (var i = 1; i < data.length; i++) { if (String(data[i][0]).trim() === name) { row = i + 1; break; } }
  if (row < 0) return out({ result: 'error', message: 'model not in AI_Models: ' + name });

  aiSheetRetry_(function() {
    if (hasRpd) sheet.getRange(row, 2).setValue(rpdN);   // col 2 = RPD_Limit
    if (hasPrio) sheet.getRange(row, 3).setValue(prioN); // col 3 = Priority
  });
  SpreadsheetApp.flush();
  _aiModelRegistryCache = null; // registry re-read เห็น limit/priority ใหม่
  var backfilled = 0;
  try {
    var cfg = getAIConfigSheet_(ss); // ensureAIConfigModelColumns_ ในตัว → คอลัมน์ _Remaining มีแน่
    if (hasRpd) {
      // backfill _Remaining=rpd ให้ทุก key Active → serve วันนี้เลย (house-style activateDisabledGeminiModels)
      var cfgData = cfg.getDataRange().getValues(), cfgHeaders = cfgData[0];
      var cfgStatusCol = cfgHeaders.indexOf("Status"), remCol = cfgHeaders.indexOf(name + "_Remaining");
      if (remCol >= 0) {
        for (var k = 1; k < cfgData.length; k++) {
          if (String(cfgData[k][0] || "").trim() && String(cfgData[k][cfgStatusCol]).trim() === "Active") {
            cfg.getRange(k + 1, remCol + 1).setValue(rpdN); backfilled++;
          }
        }
        SpreadsheetApp.flush();
      }
    }
  } catch (e) { /* คอลัมน์/backfill รอบหน้าได้ (daily-reset ก็เติมให้) */ }
  return out({ result: 'success', model: name,
               rpd: hasRpd ? rpdN : undefined, priority: hasPrio ? prioN : undefined,
               backfilledKeys: backfilled,
               note: hasRpd ? 'serve ได้ทันที (backfill ' + backfilled + ' keys)' : 'priority updated' });
}

// ดึง live models จาก models.list — แหล่งความจริงเดียวว่ามีโมเดลไหนจริง (แชร์ discover + reconcile)
// อ่าน Active key ตรงจากชีต ไม่ผ่าน getAvailableAIKey (ตัวนั้น write ตอน daily-reset → จะละเมิด read-only)
// คืน [{id, displayName, methods}] ตามหน้า (paginate). โยน error ถ้าดึงไม่ได้
function fetchLiveGeminiModels_() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(AI_CONFIG_SHEET_NAME);
  if (!sheet) throw new Error('AI_Config sheet not found');
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var colKey = headers.indexOf("API_Key");
  var colStatus = headers.indexOf("Status");
  var apiKey = "";
  for (var i = 1; i < data.length; i++) {
    var k = String(data[i][colKey] || "").trim();
    if (k.indexOf("AIza") === 0 && String(data[i][colStatus]).trim() === "Active") { apiKey = k; break; }
  }
  if (!apiKey) throw new Error('no Active AIza key in AI_Config');

  var models = [], pageToken = "", pages = 0;
  do {
    var url = "https://generativelanguage.googleapis.com/v1beta/models?key=" + encodeURIComponent(apiKey)
      + "&pageSize=200" + (pageToken ? "&pageToken=" + encodeURIComponent(pageToken) : "");
    var resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    if (resp.getResponseCode() !== 200) {
      throw new Error('models.list HTTP ' + resp.getResponseCode() + ': ' + String(resp.getContentText()).slice(0, 300));
    }
    var json = JSON.parse(resp.getContentText());
    (json.models || []).forEach(function(m) {
      models.push({
        id: String(m.name || "").replace(/^models\//, ""),
        displayName: m.displayName || "",
        methods: m.supportedGenerationMethods || []
      });
    });
    pageToken = json.nextPageToken || "";
  } while (pageToken && ++pages < 10);
  return models;
}

// Read-only discovery — GET ?action=discoverGeminiModels (model IDs ไม่ sensitive → public เหมือน aiConfigStatus)
function discoverGeminiModels() {
  function out(obj) {
    return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
  }
  try {
    var models = fetchLiveGeminiModels_();
    return out({ result: 'success', count: models.length, models: models });
  } catch (e) {
    return out({ result: 'error', message: e.message });
  }
}

// filter การ "append": เก็บเฉพาะ chat gemini-* ที่ทำ generateContent ได้ (ตัดสินใจแล้ว 2026-07-22)
// ตัด non-chat (image/tts/embedding/robotics/computer-use), preview ที่ churn, และ -latest alias ที่ลอย (พังการนับ per-model)
function isSyncableGeminiModel_(m) {
  var id = String(m.id || "");
  if (id.indexOf("gemini-") !== 0) return false;
  if ((m.methods || []).indexOf("generateContent") < 0) return false;
  if (/(image|tts|embedding|robotics|computer-use)/i.test(id)) return false;
  if (/preview|latest/i.test(id)) return false; // ตัด preview (รวมกลางสตริง เช่น -preview-customtools) + alias ลอย -latest
  return true;
}

/* =========================================================
   Phase 2 (P2-Q2/Q3): auto newer-first ranking on append — pure, testable, no GAS globals.
   parse id → (tierRank, major, minor); sort key = (tierRank asc, major desc, minor desc).
   Notes vocab (shared w/ probe auto-enable + admin panel — DO NOT drift):
     clean auto-rank : "auto-discovered <date>"
     rank fail-safe  : "auto-discovered <date> needs-manual-priority"
   ========================================================= */

// parse "gemini-X.Y-<tier>" → {tierRank, major, minor} หรือ null (curveball: preview/latest/date/-8b/customtools)
// flash-lite ต้องมาก่อน flash ใน alternation ไม่งั้น lite ไปแมตช์ prefix "flash"
function _parseGeminiRank_(id) {
  var m = String(id || "").match(/^gemini-(\d+)\.(\d+)-(flash-lite|flash|pro)$/);
  if (!m) return null;
  var tierRank = { "flash": 0, "flash-lite": 1, "pro": 2 }[m[3]];
  return { tierRank: tierRank, major: parseInt(m[1], 10), minor: parseInt(m[2], 10) };
}

// เทียบ 2 rank key: คืน <0 ถ้า a ดีกว่า (ลองก่อน = priority number น้อยกว่า)
// ดีกว่า = tierRank น้อยกว่า → major สูงกว่า → minor สูงกว่า (newer-first ในแต่ละ tier)
function _cmpGeminiRank_(a, b) {
  if (a.tierRank !== b.tierRank) return a.tierRank - b.tierRank;
  if (a.major !== b.major) return b.major - a.major;
  return b.minor - a.minor;
}

// P2-Q3 insert-without-renumber: คืน {priority} ให้ id ใหม่ตามตำแหน่ง sort-key เทียบ workingRegistry
// (existing non-Deprecated + ตัวที่เพิ่ง place รอบนี้), หรือ null → fail-safe (worst + needs-manual-priority)
// workingRegistry: [{id, priority}]. ไม่แตะ priority เดิม (เคารพ hand-tuning + monotonic lock)
function assignDiscoveredPriority_(id, workingRegistry) {
  var key = _parseGeminiRank_(id);
  if (!key) return null; // unparseable → fail-safe
  var allPrios = [], better = [], worse = [];
  for (var i = 0; i < (workingRegistry || []).length; i++) {
    var p = parseInt(workingRegistry[i].priority, 10);
    if (isNaN(p)) continue;
    allPrios.push(p);
    var rk = _parseGeminiRank_(workingRegistry[i].id);
    if (!rk) continue; // แถวเดิมที่ parse ไม่ได้ — กินสล็อต priority แต่เทียบ sort-key ไม่ได้
    var c = _cmpGeminiRank_(rk, key);
    if (c < 0) better.push(p);        // เดิมดีกว่า → ใหม่อยู่ต่อท้าย (priority มากกว่า)
    else if (c > 0) worse.push(p);    // เดิมแย่กว่า → ใหม่อยู่ก่อน (priority น้อยกว่า)
    else return null;                 // rank ชนพอดี (id ซ้ำ?) → fail-safe
  }
  if (allPrios.length === 0) return { priority: 1 };
  if (better.length === 0) return { priority: Math.min.apply(null, allPrios) - 1 }; // flagship
  if (worse.length === 0) return { priority: Math.max.apply(null, allPrios) + 1 };  // ท้ายสุด
  var prevPrio = Math.max.apply(null, better), nextPrio = Math.min.apply(null, worse);
  if (prevPrio + 1 >= nextPrio) return null; // ไม่มีช่อง integer → fail-safe
  return { priority: Math.floor((prevPrio + nextPrio) / 2) };
}

// Discovery + three-way reconcile (Q1/Q2) — GET ?action=reconcileGeminiModels
// live models.list ⇄ AI_Models: append ใหม่ Disabled/RPD=null · หายจาก API → Deprecated (ไม่ลบ คง RPD ที่ตั้งมือ) · มีทั้งคู่ → คงเดิม
// append เทียบ filter (chat gemini-* เท่านั้น); deprecation เทียบ RAW live list เต็ม → เคารพ preview ที่ owner เพิ่มมือ
// โครง monotonic: แถว append ท้ายเท่านั้น, คอลัมน์ผ่าน ensureAIConfigModelColumns_ (getLastColumn()+1) — index เดิมไม่ขยับ
function reconcileGeminiModels() {
  function out(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
  var live;
  try { live = fetchLiveGeminiModels_(); } catch (e) { return out({ result: 'error', step: 'discovery', message: e.message }); }
  var rawLiveSet = {};
  live.forEach(function(m) { rawLiveSet[m.id] = true; });
  var appendable = live.filter(isSyncableGeminiModel_);

  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(AI_MODELS_SHEET_NAME);
  if (!sheet) return out({ result: 'error', message: 'AI_Models sheet not found' });
  var data = sheet.getDataRange().getValues();
  var colModel = 0, colPrio = 2, colStatus = 3; // [Model, RPD_Limit, Priority, Status, Notes]
  var sheetModels = {}, maxPrio = 0, workingRegistry = [];
  for (var i = 1; i < data.length; i++) {
    var name = String(data[i][colModel] || "").trim();
    if (!name) continue;
    var status = String(data[i][colStatus] || "").trim();
    sheetModels[name] = { row: i + 1, status: status };
    var p = parseInt(data[i][colPrio], 10);
    if (!isNaN(p) && p > maxPrio) maxPrio = p;
    // ranker เทียบเฉพาะแถวที่ยังใช้งาน (Deprecated ไม่นับเป็นเพื่อนบ้าน)
    if (status !== "Deprecated" && !isNaN(p)) workingRegistry.push({ id: name, priority: p });
  }

  var today = Utilities.formatDate(new Date(), "Asia/Bangkok", "yyyy-MM-dd");
  var added = [], deprecated = [], leftCount = 0;

  // (1) append: appendable ที่ยังไม่มีในชีต → Disabled, RPD ว่าง (null)
  // P2-Q2/Q3: auto newer-first ranking. process ตาม sort-key (ดีสุดก่อน) แล้ว fold แต่ละตัวเข้า
  // workingRegistry ก่อน place ตัวถัดไป → โมเดลที่มาพร้อมกันรอบเดียว rank เทียบกันเองด้วย
  var pending = appendable.filter(function(m) {
    if (sheetModels[m.id]) { leftCount++; return false; }
    return true;
  });
  pending.sort(function(a, b) {
    var ka = _parseGeminiRank_(a.id), kb = _parseGeminiRank_(b.id);
    if (ka && kb) return _cmpGeminiRank_(ka, kb);
    if (ka) return -1; if (kb) return 1; return 0; // parse ไม่ได้ → ไปท้าย (place หลัง = worst)
  });
  var newRows = [];
  pending.forEach(function(m) {
    var r = assignDiscoveredPriority_(m.id, workingRegistry);
    var prio, notes;
    if (r === null) { // fail-safe: worst priority + flag ให้ owner ตั้งเอง
      maxPrio += 1; prio = maxPrio;
      notes = "auto-discovered " + today + " needs-manual-priority";
    } else {
      prio = r.priority;
      if (prio > maxPrio) maxPrio = prio;
      notes = "auto-discovered " + today;
    }
    newRows.push([m.id, "", prio, "Disabled", notes]);
    added.push(m.id);
    workingRegistry.push({ id: m.id, priority: prio }); // fold in ก่อน place ตัวถัดไป
  });
  if (newRows.length) {
    aiSheetRetry_(function() {
      sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, 5).setValues(newRows);
      SpreadsheetApp.flush();
    });
  }

  // (2) deprecate: ในชีตแต่หายจาก RAW live list → Deprecated (ไม่ลบ). ข้ามตัวที่ Deprecated อยู่แล้ว
  Object.keys(sheetModels).forEach(function(name) {
    if (rawLiveSet[name] || sheetModels[name].status === "Deprecated") return;
    aiSheetRetry_(function() { sheet.getRange(sheetModels[name].row, colStatus + 1).setValue("Deprecated"); });
    deprecated.push(name);
  });
  if (deprecated.length) SpreadsheetApp.flush();

  // (3) เติมคอลัมน์ <model>_Remaining สำหรับ registry ใหม่ (append-only ที่ getLastColumn()+1)
  var colsBefore = null, colsAfter = null;
  try {
    _aiModelRegistryCache = null; // reset cache → re-read หลัง mutate แถว
    var cfg = ss.getSheetByName(AI_CONFIG_SHEET_NAME);
    colsBefore = cfg.getLastColumn();
    ensureAIConfigModelColumns_(cfg, getAIModelRegistry_(ss));
    colsAfter = cfg.getLastColumn();
  } catch (e) { /* คอลัมน์เติมรอบหน้าได้ ไม่ critical */ }

  return out({ result: 'success', liveTotal: live.length, appendable: appendable.length,
               added: added, deprecated: deprecated, leftUnchanged: leftCount,
               configColsBefore: colsBefore, configColsAfter: colsAfter });
}

/* =========================================================
   429 policy (Q3/Q4): แยก RPD (PerDay) ออกจาก RPM (PerMinute)
   - PerDay  → zero โควต้าโมเดลนั้นทั้งวัน (markModelExhausted_ เดิม)
   - PerMinute/unknown → cooldown สั้นใน CacheService, ไม่แตะโควต้าวัน → ลองโมเดล/คีย์ถัดไป
   เหตุผล unknown→cooldown: ถ้า zero ผิดตอน RPM จะทำ flash รุ่น RPM ต่ำฆ่าตัวเอง แล้วตกไปใช้ subscription ทั้งที่ยังมีโควต้าวัน
   ground-truth shape (compat endpoint 429): [{error:{status:"RESOURCE_EXHAUSTED",message,details:[{QuotaFailure:{violations:[{quotaId:"...PerMinute..."}]}},{RetryInfo:{retryDelay:"59s"}}]}}]
   ========================================================= */

// จำแนก metric ของ 429 → { metric: 'perDay'|'perMinute'|'unknown', retrySec }
function parseGemini429_(body, headers) {
  var metric = 'unknown', retrySec = 0;
  try {
    var j = typeof body === 'string' ? JSON.parse(body) : body;
    if (Array.isArray(j)) j = j[0] || {};        // compat endpoint ครอบด้วย array
    var err = j.error || j;
    var blob = JSON.stringify(err.details || {}) + " " + String(err.message || "");
    if (/PerMinute|RequestsPerMinute|per[\s_-]?minute/i.test(blob)) metric = 'perMinute';
    else if (/PerDay|RequestsPerDay|per[\s_-]?day/i.test(blob)) metric = 'perDay';
    var m = JSON.stringify(err.details || {}).match(/"retryDelay"\s*:\s*"(\d+)(?:\.\d+)?s"/);
    if (!m) m = String(err.message || "").match(/retry in (\d+)(?:\.\d+)?\s*s/i);
    if (m) retrySec = parseInt(m[1], 10);
  } catch (e) {}
  if (!retrySec && headers) {
    var ra = headers["Retry-After"] || headers["retry-after"];
    if (ra) retrySec = parseInt(ra, 10) || 0;
  }
  if (!retrySec || retrySec < 1) retrySec = 60;   // ไม่มีสัญญาณ → default 60s
  return { metric: metric, retrySec: retrySec };
}

// cooldown แบบ ephemeral ต่อ (key,model) — GAS doPost ไม่มี state ข้าม call จึงเก็บใน CacheService
function _cooldownCacheKey_(apiKey, model) { return "cooldown:" + String(apiKey).slice(-4) + ":" + model; }
function setModelCooldown_(apiKey, model, ttlSec) {
  try {
    var ttl = Math.min(Math.max(parseInt(ttlSec, 10) || 60, 1), 21600); // CacheService cap = 6h
    CacheService.getScriptCache().put(_cooldownCacheKey_(apiKey, model), "1", ttl);
  } catch (e) { console.warn("setModelCooldown_ failed: " + e.message); }
}
function isModelCoolingDown_(apiKey, model) {
  try { return CacheService.getScriptCache().get(_cooldownCacheKey_(apiKey, model)) !== null; }
  catch (e) { return false; }
}

// ตัวจัดการ 429 เดียวที่ทุก path เรียก — คืน { action:'exhausted'|'cooldown', metric, retrySec }
function handleGemini429_(apiKeyInfo, model, body, headers) {
  var p = parseGemini429_(body, headers);
  if (p.metric === 'perDay') {
    markModelExhausted_(apiKeyInfo, model);                 // โควต้าวันหมดจริง → zero
    return { action: 'exhausted', metric: p.metric, retrySec: p.retrySec };
  }
  setModelCooldown_(apiKeyInfo.key, model, p.retrySec);      // perMinute/unknown → cooldown, ไม่แตะโควต้าวัน
  return { action: 'cooldown', metric: p.metric, retrySec: p.retrySec };
}

// Verify action — GET ?action=verifyRpmCooldown
// burst compat จน 429 จริง แล้วรัน handleGemini429_ → ยืนยัน perMinute: _Remaining ไม่ถูกแตะ + cooldown ถูกตั้ง
function verifyRpmCooldown() {
  function out(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
  var ss = SpreadsheetApp.openById(SHEET_ID), sheet = ss.getSheetByName(AI_CONFIG_SHEET_NAME);
  if (!sheet) return out({ result: 'error', message: 'AI_Config not found' });
  var data = sheet.getDataRange().getValues(), headers = data[0];
  var colKey = headers.indexOf("API_Key"), colStatus = headers.indexOf("Status");
  var model = "gemini-3.1-flash-lite", remCol = headers.indexOf(model + "_Remaining");
  var apiKey = "", rowIndex = -1;
  for (var i = 1; i < data.length; i++) {
    var k = String(data[i][colKey] || "").trim();
    if (k.indexOf("AIza") === 0 && String(data[i][colStatus]).trim() === "Active") { apiKey = k; rowIndex = i + 1; break; }
  }
  if (!apiKey) return out({ result: 'error', message: 'no active key' });
  var apiKeyInfo = { key: apiKey, index: rowIndex, model: model };
  var remainingBefore = remCol >= 0 ? Number(sheet.getRange(rowIndex, remCol + 1).getValue()) : null;

  for (var n = 0; n < 30; n++) {
    var resp = UrlFetchApp.fetch(GEMINI_OPENAI_COMPAT_ENDPOINT, {
      method: "post", contentType: "application/json",
      headers: { Authorization: "Bearer " + apiKey },
      payload: JSON.stringify({ model: model, messages: [{ role: "user", content: "ping" }], max_tokens: 4 }),
      muteHttpExceptions: true
    });
    if (resp.getResponseCode() === 429) {
      var r = handleGemini429_(apiKeyInfo, model, resp.getContentText(), resp.getAllHeaders());
      var remainingAfter = remCol >= 0 ? Number(sheet.getRange(rowIndex, remCol + 1).getValue()) : null;
      return out({ result: 'success', trippedAt: n + 1, handled: r,
                   remainingBefore: remainingBefore, remainingAfter: remainingAfter,
                   remainingUntouched: remainingBefore === remainingAfter,
                   cooldownSet: isModelCoolingDown_(apiKey, model) });
    }
  }
  return out({ result: 'no429', note: 'RPM not tripped within 30 calls' });
}

// helper: อ่าน Active AIza key ตัวแรกจาก AI_Config ตรงๆ (read-only) — ใช้ร่วมหลาย diagnostic/enable-gate
function firstActiveAiKey_(ss) {
  ss = ss || SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(AI_CONFIG_SHEET_NAME);
  if (!sheet) return "";
  var data = sheet.getDataRange().getValues(), headers = data[0];
  var colKey = headers.indexOf("API_Key"), colStatus = headers.indexOf("Status");
  for (var i = 1; i < data.length; i++) {
    var k = String(data[i][colKey] || "").trim();
    if (k.indexOf("AIza") === 0 && String(data[i][colStatus]).trim() === "Active") return k;
  }
  return "";
}

// Q6 tool-capability smoke test (fail-closed): ยิง compat endpoint ด้วย tool "ping" แล้วต้องได้ tool_calls กลับ
// (generateContent support ≠ tool-call support; Claude Code ต้องใช้ tool_calls) → คืน { pass, reason }
function geminiToolSmokeTest_(model, apiKey) {
  var payload = {
    model: model,
    messages: [{ role: "user", content: "Call the ping function now." }],
    tools: [{ type: "function", function: {
      name: "ping", description: "Returns pong.",
      parameters: { type: "object", properties: {}, required: [] }
    } }],
    // force ping — วัด "ทำ tool_calls ได้ไหม" (capability) ไม่ใช่ "อยากทำไหม" (propensity)
    // tool_choice:"auto" เสี่ยง false-negative: โมเดลที่ทำ tool ได้แต่เลือกตอบ text → deprecate ผิดใน weekly probe
    tool_choice: { type: "function", function: { name: "ping" } },
    // 64 ถูกพบว่าเตี้ยเกินไปสำหรับโมเดล thinking-by-default รุ่นใหม่ (เช่น gemini-3.6-flash) — reasoning
    // tokens กิน budget ก่อนถึง tool_calls จริง → finish=length false-negative ทั้งที่ capable
    // (2026-07-24, ยืนยันจาก gemini-3.6-flash ที่ fail ด้วย "no tool_calls (finish=length)")
    max_tokens: 1024
  };
  try {
    var resp = UrlFetchApp.fetch(GEMINI_OPENAI_COMPAT_ENDPOINT, {
      method: "post", contentType: "application/json",
      headers: { Authorization: "Bearer " + apiKey },
      payload: JSON.stringify(payload), muteHttpExceptions: true
    });
    var code = resp.getResponseCode();
    // transient = 429/5xx: ไม่ใช่ตัวชี้ว่า tool-incapable → weekly probe ต้องไม่ deprecate เพราะเหตุนี้
    if (code !== 200) return { pass: false, transient: (code === 429 || code >= 500),
                               reason: "HTTP " + code + ": " + String(resp.getContentText()).slice(0, 160) };
    var j = JSON.parse(resp.getContentText());
    var msg = j.choices && j.choices[0] && j.choices[0].message;
    var tc = msg && msg.tool_calls;
    if (tc && tc.length && tc[0].function && tc[0].function.name === "ping") return { pass: true, reason: "tool_calls ok" };
    return { pass: false, transient: false, reason: "no tool_calls (finish=" + (j.choices && j.choices[0] && j.choices[0].finish_reason) + ")" };
  } catch (e) { return { pass: false, transient: true, reason: e.message }; }
}

// Q6 enable-gate — GET ?action=enableGeminiModel&model=<id>
// Disabled → Active ต้องผ่าน smoke test ก่อน (fail-closed). ผ่าน → Status=Active (owner ตั้ง RPD_Limit เองตาม Q1)
// ไม่ผ่าน → คง Disabled, เขียน Notes "tool-incapable: <reason>"
function enableGeminiModel(modelParam) {
  function out(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
  var model = String(modelParam || "").trim();
  if (!model) return out({ result: 'error', message: 'missing model param' });
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(AI_MODELS_SHEET_NAME);
  if (!sheet) return out({ result: 'error', message: 'AI_Models not found' });
  var data = sheet.getDataRange().getValues();
  var statusCol = 3, notesCol = 4, row = -1;
  for (var i = 1; i < data.length; i++) { if (String(data[i][0]).trim() === model) { row = i + 1; break; } }
  if (row < 0) return out({ result: 'error', message: 'model not in AI_Models: ' + model });
  var apiKey = firstActiveAiKey_(ss);
  if (!apiKey) return out({ result: 'error', message: 'no active key' });

  var today = Utilities.formatDate(new Date(), "Asia/Bangkok", "yyyy-MM-dd");
  var t = geminiToolSmokeTest_(model, apiKey);
  if (!t.pass) {
    aiSheetRetry_(function() { sheet.getRange(row, notesCol + 1).setValue("tool-incapable: " + t.reason + " (" + today + ")"); });
    SpreadsheetApp.flush();
    return out({ result: 'refused', model: model, reason: t.reason });
  }
  aiSheetRetry_(function() { sheet.getRange(row, statusCol + 1).setValue("Active"); });
  SpreadsheetApp.flush();
  return out({ result: 'success', model: model, enabled: true, note: 'ตั้ง RPD_Limit > 0 เพื่อเริ่ม serve (Q1 human-confirmed)' });
}

// "ปุ่ม activate" — รันจาก editor เท่านั้น (ไม่ผูก doGet: deploy นี้เป็น public no-auth ที่นิสิตเรียก)
// Batch: ทุกแถว Disabled ที่ model ขึ้นต้น gemini- และ RPD_Limit>0 → smoke test → ผ่าน = Active + backfill
// <model>_Remaining = RPD ทุก key ที่ Active (serve ได้วันนี้ ไม่ต้องรอ daily reset) · fail จริง = Notes tool-incapable
// · transient (429/5xx) = ไม่แตะ Notes คง Disabled ให้ลองใหม่ (กัน deprecate ผิดจาก blip)
// เลือก scope ด้วย RPD: ตั้ง RPD>0 เฉพาะโมเดลที่ต้องการ (เว้น -001 alias/RPD=0 ไว้ = ข้าม)
function activateDisabledGeminiModels() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var mSheet = ss.getSheetByName(AI_MODELS_SHEET_NAME);
  if (!mSheet) return { result: 'error', message: 'AI_Models not found' };
  var apiKey = firstActiveAiKey_(ss);
  if (!apiKey) return { result: 'error', message: 'no active key for smoke test' };

  var cfg = getAIConfigSheet_(ss);
  var cfgData = cfg.getDataRange().getValues(), cfgHeaders = cfgData[0];
  var cfgStatusCol = cfgHeaders.indexOf("Status");
  var today = Utilities.formatDate(new Date(), "Asia/Bangkok", "yyyy-MM-dd");

  var data = mSheet.getDataRange().getValues();
  var statusCol = 3, notesCol = 4; // [Model, RPD_Limit, Priority, Status, Notes]
  var activated = [], refused = [], transient = [], skipped = [];

  for (var i = 1; i < data.length; i++) {
    var model = String(data[i][0] || "").trim();
    var rpd = parseInt(data[i][1], 10) || 0;
    var status = String(data[i][3] || "").trim();
    if (model.indexOf("gemini-") !== 0) continue;
    if (status !== "Disabled") continue;
    if (rpd <= 0) { skipped.push(model + " (RPD<=0)"); continue; }

    var t = geminiToolSmokeTest_(model, apiKey);
    if (!t.pass) {
      if (t.transient) { transient.push(model + ": " + t.reason); continue; } // blip → คง Disabled ไม่เขียน note
      aiSheetRetry_((function(r, reason) { return function() {
        mSheet.getRange(r, notesCol + 1).setValue("tool-incapable: " + reason + " (" + today + ")");
      }; })(i + 1, t.reason));
      refused.push(model + ": " + t.reason);
      continue;
    }
    // ผ่าน — flip Active + backfill _Remaining=RPD ทุก key ที่ Active (serve วันนี้เลย)
    aiSheetRetry_((function(r) { return function() {
      mSheet.getRange(r, statusCol + 1).setValue("Active");
    }; })(i + 1));
    var remCol = cfgHeaders.indexOf(model + "_Remaining");
    var filled = 0;
    if (remCol >= 0) {
      for (var k = 1; k < cfgData.length; k++) {
        if (String(cfgData[k][cfgStatusCol]).trim() !== "Active") continue;
        cfg.getRange(k + 1, remCol + 1).setValue(rpd);
        filled++;
      }
    }
    activated.push(model + " (RPD=" + rpd + ", backfilled " + filled + " keys" + (remCol < 0 ? ", NO _Remaining col — serves next reset" : "") + ")");
  }
  SpreadsheetApp.flush();
  var summary = { result: 'success', activated: activated, refused: refused, transient: transient, skipped: skipped };
  console.log("[activateDisabled] " + JSON.stringify(summary));
  return summary;
}

// Verify — GET ?action=verifyToolSmokeTest: known tool-capable (gemini-2.5-flash → pass) vs text-only (gemma-4-31b-it → refuse)
function verifyToolSmokeTest() {
  function out(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
  var apiKey = firstActiveAiKey_();
  if (!apiKey) return out({ result: 'error', message: 'no active key' });
  return out({ result: 'success',
    capable_gemini_2_5_flash: geminiToolSmokeTest_("gemini-2.5-flash", apiKey),   // → pass
    incapable_embedding_001: geminiToolSmokeTest_("gemini-embedding-001", apiKey) }); // no generateContent → fail-closed refuse
}

/* =========================================================
   Step 5: scheduled sync (Q5/Q6 cadence)
   - runGeminiModelSync : DAILY — quota reset (all Active-key rows) + models.list diff (reconcile)
   - runGeminiToolProbe : WEEKLY — tool smoke test บนโมเดล Active; genuine fail → Deprecated+Notes; transient → ข้าม
   "alert to MDKKUQUIZDATABASE" = console log + สถานะ Deprecated/Notes บนชีต ที่ dashboard อ่านผ่าน aiConfigStatus อยู่แล้ว
   (backend นี้ไม่มี mail/webhook infra — สถานะบนชีตคือช่องทาง alert)
   ========================================================= */

// Batch daily reset — เดิม getAvailableAIKey มี reset ต่อแถวแบบ lazy (บรรทัด ~252) แต่ทำงานเฉพาะ
// แถวที่ถูกสแกนตอนมี request เข้ามาเท่านั้น วันที่ไม่มี agentQuery เลย โควต้าจะไม่ refill จนกว่า request ถัดไป
// ฟังก์ชันนี้ทำ reset แบบ unconditional จาก trigger รายวัน ไม่ต้องพึ่ง traffic — เกณฑ์เดียวกับ lazy reset:
// Last_Reset_Date ของแถว (key ที่ Status=Active) ไม่ตรงวันนี้ (Asia/Bangkok) → เติม <model>_Remaining
// กลับเป็น RPD_Limit ของทุกโมเดล Active ใน AI_Models แล้วเซ็ต Last_Reset_Date = วันนี้
// ไม่แตะโครงคอลัมน์ (เขียนแค่ value) — monotonic เดิมยังอยู่
function resetExpiredGeminiQuotas_() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = getAIConfigSheet_(ss);
  var models = getAIModelRegistry_(ss).filter(function(m) { return m.active; });
  if (!models.length) return { checked: 0, reset: 0 };

  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var colStatus = headers.indexOf("Status");
  var colLastReset = headers.indexOf("Last_Reset_Date");
  var tz = "Asia/Bangkok"; // อย่าใช้ timezone ของ script (อาจเป็น UTC — reset ช้า 7 ชม.)
  var todayStr = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd");

  var checked = 0, reset = 0;
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!String(row[0] || "").trim()) continue;
    if (row[colStatus] !== "Active") continue;
    checked++;
    var lastResetStr = row[colLastReset]
      ? Utilities.formatDate(new Date(row[colLastReset]), tz, "yyyy-MM-dd") : "";
    if (lastResetStr === todayStr) continue;

    (function(rowIndex) {
      aiSheetRetry_(function() {
        for (var m = 0; m < models.length; m++) {
          var rc = headers.indexOf(models[m].model + "_Remaining");
          if (rc >= 0) sheet.getRange(rowIndex + 1, rc + 1).setValue(models[m].limit);
        }
        sheet.getRange(rowIndex + 1, colLastReset + 1).setValue(todayStr);
      });
    })(i);
    reset++;
  }
  if (reset) SpreadsheetApp.flush();
  console.log("[geminiQuotaReset] checked=" + checked + " reset=" + reset);
  return { checked: checked, reset: reset };
}

// DAILY job — reset โควต้าแถวที่ข้ามวัน (ไม่พึ่ง traffic) + reconcile รายการโมเดลจาก live models.list
function runGeminiModelSync() {
  var resetRes = resetExpiredGeminiQuotas_();
  var res = JSON.parse(reconcileGeminiModels().getContent());
  res.quotaReset = resetRes;
  console.log("[geminiSync] daily reconcile: added=" + JSON.stringify(res.added || [])
    + " deprecated=" + JSON.stringify(res.deprecated || []) + " left=" + res.leftUnchanged
    + " quotaReset=" + JSON.stringify(resetRes));
  return res;
}

// WEEKLY job — tool-capability smoke test. สองบทบาท:
//  (a) Active gemini-* → re-probe; genuine fail → Deprecated, transient (429/5xx) → ข้าม
//  (b) P2-Q4 auto-enable: Disabled ที่ Notes มี "auto-discovered" → smoke test; pass → Active (RPD ยังว่าง = dormant
//      ตาม P2-Q5 interlock จน owner ตั้ง RPD), genuine fail → คง Disabled + Notes tool-incapable, transient → ข้าม
//  ข้าม owner-hand-disabled (Notes ไม่มี "auto-discovered" เช่น gemini-2.5-pro "free tier RPD = 0") และ Deprecated
function runGeminiToolProbe() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(AI_MODELS_SHEET_NAME);
  if (!sheet) { console.warn("[geminiProbe] ไม่พบ AI_Models"); return { checked: 0, deprecated: 0 }; }
  var apiKey = firstActiveAiKey_(ss);
  if (!apiKey) { console.warn("[geminiProbe] ไม่มี active key"); return { checked: 0, deprecated: 0 }; }
  var data = sheet.getDataRange().getValues();
  var prioCol = 2, statusCol = 3, notesCol = 4;
  var today = Utilities.formatDate(new Date(), "Asia/Bangkok", "yyyy-MM-dd");
  var checked = 0, deprecated = 0, enabled = 0, refused = 0, skipped = 0, results = [];
  for (var i = 1; i < data.length; i++) {
    var name = String(data[i][0] || "").trim();
    if (!name || name.indexOf("gemini-") !== 0) continue;
    var status = String(data[i][statusCol]).trim();
    var notes = String(data[i][notesCol] || "");
    var isActive = status === "Active";
    var isAutoDisabled = status === "Disabled" && /auto-discovered/i.test(notes);
    if (!isActive && !isAutoDisabled) continue; // owner-hand-disabled + Deprecated → ข้าม
    checked++;
    var t = geminiToolSmokeTest_(name, apiKey);

    if (isAutoDisabled) { // (b) auto-enable path
      if (t.pass) {
        var prio = parseInt(data[i][prioCol], 10);
        var keepFlag = /needs-manual-priority/i.test(notes) ? "needs-manual-priority, " : ""; // ต้องคง flag
        sheet.getRange(i + 1, statusCol + 1).setValue("Active");
        sheet.getRange(i + 1, notesCol + 1).setValue(keepFlag + "auto-discovered, tool-OK, priority "
          + (isNaN(prio) ? "?" : prio) + " — SET RPD TO GO LIVE (" + today + ")");
        enabled++; results.push(name + ":ENABLED");
        console.log("[geminiProbe] " + name + " → Active (auto-enabled, รอ RPD)");
      } else if (t.transient) {
        skipped++; results.push(name + ":transient");
      } else {
        // เก็บ "auto-discovered" ไว้ในข้อความเสมอ — ไม่งั้น isAutoDisabled (บรรทัด 928) จะไม่ match รอบถัดไป
        // แล้วแถวนี้หลุดออกจาก retry pool ถาวร (แยกไม่ออกจาก owner-hand-disabled) ทั้งที่ fail อาจเป็น false
        // negative ชั่วคราว (เช่น max_tokens ของ smoke test เตี้ยไปสำหรับโมเดล thinking — พบ 2026-07-24)
        var keepFlag2 = /needs-manual-priority/i.test(notes) ? "needs-manual-priority, " : "";
        sheet.getRange(i + 1, notesCol + 1).setValue(keepFlag2 + "auto-discovered, tool-incapable: " + t.reason + " (" + today + ")");
        refused++; results.push(name + ":tool-incapable");
        console.warn("[geminiProbe] " + name + " tool-incapable: " + t.reason);
      }
      continue;
    }

    // (a) Active re-probe path
    if (t.pass) { results.push(name + ":pass"); continue; }
    if (t.transient) { skipped++; results.push(name + ":transient"); console.log("[geminiProbe] " + name + " transient, skip: " + t.reason); continue; }
    sheet.getRange(i + 1, statusCol + 1).setValue("Deprecated");
    sheet.getRange(i + 1, notesCol + 1).setValue("weekly-probe-failed: " + t.reason + " (" + today + ")");
    deprecated++; results.push(name + ":DEPRECATED");
    console.warn("[geminiProbe] " + name + " → Deprecated: " + t.reason);
  }
  SpreadsheetApp.flush();
  console.log("[geminiProbe] checked=" + checked + " enabled=" + enabled + " refused=" + refused
    + " deprecated=" + deprecated + " skippedTransient=" + skipped);
  return { checked: checked, enabled: enabled, refused: refused, deprecated: deprecated, skippedTransient: skipped, results: results };
}

// idempotent trigger installer — daily reconcile ~ตี 6, weekly probe อาทิตย์ ~ตี 6 (เว้นตี 3-5 ให้ job เดิม)
// รันจาก editor ได้เสมอ; ผ่าน web-app ต้องมี scope script.scriptapp (execute-as-owner)
function installGeminiSyncTriggers() {
  function out(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
  var handlers = { runGeminiModelSync: true, runGeminiToolProbe: true };
  try {
    var triggers = ScriptApp.getProjectTriggers();
    for (var i = 0; i < triggers.length; i++) {
      if (handlers[triggers[i].getHandlerFunction()]) ScriptApp.deleteTrigger(triggers[i]);
    }
    ScriptApp.newTrigger('runGeminiModelSync').timeBased().everyDays(1).atHour(6).create();
    ScriptApp.newTrigger('runGeminiToolProbe').timeBased().onWeekDay(ScriptApp.WeekDay.SUNDAY).atHour(6).create();
    var now = ScriptApp.getProjectTriggers().filter(function(t) { return handlers[t.getHandlerFunction()]; })
      .map(function(t) { return t.getHandlerFunction() + "/" + t.getEventType(); });
    return out({ result: 'success', installed: now });
  } catch (e) {
    return out({ result: 'error', message: e.message, hint: 'รัน installGeminiSyncTriggers() จาก Apps Script editor แทน' });
  }
}

// Read-only probe — GET ?action=probeGeminiTier
// ยิง endpoint + model เดียวกับ agentQuery Gemini tier (GEMINI_OPENAI_COMPAT_ENDPOINT + AGENT_QUERY_MODEL_OVERRIDE.Gemini)
// เพื่อยืนยันว่า tier คืน 200 จริงหลัง cleanup — ไม่หัก quota (ไม่เรียก updateAIUsage), ใช้ key ตรงจากชีต
function probeGeminiTier() {
  function out(obj) {
    return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
  }
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(AI_CONFIG_SHEET_NAME);
  if (!sheet) return out({ result: 'error', message: 'AI_Config sheet not found' });
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var colKey = headers.indexOf("API_Key");
  var colStatus = headers.indexOf("Status");
  var apiKey = "";
  for (var i = 1; i < data.length; i++) {
    var k = String(data[i][colKey] || "").trim();
    if (k.indexOf("AIza") === 0 && String(data[i][colStatus]).trim() === "Active") { apiKey = k; break; }
  }
  if (!apiKey) return out({ result: 'error', message: 'no Active AIza key in AI_Config' });

  var model = AGENT_QUERY_MODEL_OVERRIDE["Gemini"];
  var resp = UrlFetchApp.fetch(GEMINI_OPENAI_COMPAT_ENDPOINT, {
    method: "post", contentType: "application/json",
    headers: { "Authorization": "Bearer " + apiKey },
    payload: JSON.stringify({ model: model, messages: [{ role: "user", content: "ping" }], max_tokens: 8 }),
    muteHttpExceptions: true
  });
  var code = resp.getResponseCode();
  return out({ result: 'success', model: model, code: code, ok: code === 200,
               snippet: String(resp.getContentText()).slice(0, 300) });
}

// One-off cleanup — GET ?action=purgeFakeGeminiModelRows
// ลบแถวโมเดลปลอมออกจาก AI_Models (ไม่มีใน models.list จริง): gemini-3-flash, gemini-3.1-pro
// idempotent — รันซ้ำได้ (ไม่เจอก็ข้าม). ไม่แตะคอลัมน์ orphan _Remaining ใน AI_Config (คงโครง monotonic)
function purgeFakeGeminiModelRows() {
  function out(obj) {
    return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
  }
  var FAKE = { "gemini-3-flash": true, "gemini-3.1-pro": true,
               "gemini-3.1-pro-preview-customtools": true }; // ตัวหลัง = leaked จาก filter รอบแรก (preview กลางสตริง) — ลบทิ้ง
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(AI_MODELS_SHEET_NAME);
  if (!sheet) return out({ result: 'error', message: 'AI_Models sheet not found' });
  var data = sheet.getDataRange().getValues();
  var removed = [];
  for (var i = data.length - 1; i >= 1; i--) { // bottom-up: index ของแถวที่เหลือไม่เลื่อน
    var name = String(data[i][0] || "").trim();
    if (FAKE[name]) { sheet.deleteRow(i + 1); removed.push(name); }
  }
  SpreadsheetApp.flush();
  return out({ result: 'success', removed: removed, rowsLeft: sheet.getLastRow() - 1 });
}

// ลบคอลัมน์ <model>_Remaining ใน AI_Config ที่ไม่มีแถวใน AI_Models แล้ว (orphan จาก rename/purge — เช่นคู่กับ purgeFakeGeminiModelRows
// ที่ลบเฉพาะแถว AI_Models ทิ้งคอลัมน์ค้าง). Registry = แหล่งความจริง; Disabled ก็ยังอยู่ใน registry จึงคงคอลัมน์ไว้ (ลบเฉพาะที่หายจริง).
// Editor/trigger-only (ไม่ผูก doGet — deployment public no-auth). ลบขวา→ซ้ายเพื่อ index ไม่เลื่อน.
function purgeOrphanAiConfigColumns() {
  function out(obj) {
    return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
  }
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = getAIConfigSheet_(ss);
  var known = {};
  getAIModelRegistry_(ss).forEach(function(m) { known[m.model] = true; });
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var removed = [];
  for (var c = headers.length - 1; c >= 0; c--) {
    var h = String(headers[c] || "");
    if (h.slice(-10) !== "_Remaining") continue; // เว้น 5 คอลัมน์ fixed (ไม่ลงท้าย _Remaining)
    var model = h.slice(0, -10);
    if (known[model]) continue; // ยังอยู่ใน registry (รวม Disabled) — คงไว้
    aiSheetRetry_(function() { sheet.deleteColumn(c + 1); });
    removed.push(h);
  }
  SpreadsheetApp.flush();
  return out({ result: 'success', removed: removed, colsLeft: sheet.getLastColumn() });
}

// One-off/idempotent: สร้าง AI_Models + migrate AI_Config เป็นโครงใหม่ — GET ?action=setupAIConfig
function setupAIConfigSheet() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var hadModels = !!ss.getSheetByName(AI_MODELS_SHEET_NAME);
  var legacy = false;
  var cfg = ss.getSheetByName(AI_CONFIG_SHEET_NAME);
  if (cfg) {
    var h = cfg.getRange(1, 1, 1, Math.max(cfg.getLastColumn(), 1)).getValues()[0];
    legacy = h.indexOf("Last_Reset_Date") < 0;
  }
  var sheet = getAIConfigSheet_(ss);
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  return ContentService.createTextOutput(JSON.stringify({
    result: 'success',
    modelsSheetCreated: !hadModels,
    migratedFromLegacy: legacy,
    keyRows: Math.max(sheet.getLastRow() - 1, 0),
    headers: headers
  })).setMimeType(ContentService.MimeType.JSON);
}


/**
 * ฟังก์ชันหลักในการเรียก Gemini API พร้อมระบบ Google Search Grounding และรับส่งไฟล์รูปภาพ
 */
function callGeminiAI(prompt, apiKeyInfo, images) {
  // ใช้ v1beta เสมอเพื่อรองรับฟีเจอร์ Google Search Grounding และ Thinking ในการอัปเดตเกณฑ์การรักษาล่าสุด
  var apiVersion = 'v1beta';
  var url = "https://generativelanguage.googleapis.com/" + apiVersion + "/models/" + apiKeyInfo.model + ":generateContent?key=" + apiKeyInfo.key;
  
  var parts = [{ "text": prompt }];
  
  // แปลง URL รูปภาพหรือข้อมูล Base64 ให้เป็นอินพุตประเภท inlineData สำหรับ Gemini API
  if (images && Array.isArray(images)) {
    images.forEach(function(imgUrl) {
      try {
        var base64Data = "";
        var mimeType = "image/png"; // ค่าเริ่มต้น
        
        if (imgUrl.indexOf("data:image") === 0) {
          // กรณีรูปภาพเป็นข้อมูล Base64 ตรงจากฝั่งหน้าบ้าน
          var partsBase64 = imgUrl.split(",");
          mimeType = partsBase64[0].match(/:(.*?);/)[1];
          base64Data = partsBase64[1];
        } else if (imgUrl.indexOf("drive.google.com") > -1 || imgUrl.indexOf("googleusercontent.com") > -1) {
          // กรณีรูปภาพเก็บไว้บน Google Drive
          var fileId = "";
          var match = imgUrl.match(/\/d\/([^\/\?]+)/) || imgUrl.match(/id=([^&]+)/);
          if (match) fileId = match[1];
          
          if (fileId) {
            var file = DriveApp.getFileById(fileId);
            mimeType = file.getMimeType();
            base64Data = Utilities.base64Encode(file.getBlob().getBytes());
          }
        } else if (imgUrl.indexOf("http") === 0) {
          // กรณีรูปภาพเก็บอยู่บนอินเทอร์เน็ตภายนอกทั่วไป
          var response = UrlFetchApp.fetch(imgUrl, { "muteHttpExceptions": true });
          if (response.getResponseCode() === 200) {
            mimeType = response.getHeaders()["Content-Type"] || "image/png";
            base64Data = Utilities.base64Encode(response.getBlob().getBytes());
          }
        }
        
        if (base64Data) {
          parts.push({
            "inlineData": {
              "mimeType": mimeType,
              "data": base64Data
            }
          });
        }
      } catch (e) {
        console.warn("Failed to attach image to Gemini payload: " + imgUrl + " - " + e.message);
      }
    });
  }
  
  var payload = {
    "contents": [{
      "parts": parts
    }],
    "systemInstruction": {
      "parts": [{ "text": "You are a Medical Education Expert. Focus on Pathophysiology and Clinical Reasoning. If you need to verify medical guidelines (like AHA, GINA, GOLD, KDIGO) to formulate the response, use the search tool to find the most accurate and up-to-date recommendations." }]
    },
    // "tools": [{
    //   "google_search": {} // เปิดใช้งานระบบสืบค้นข้อมูล Google Search Grounding ของจริง
    // }],
    "generationConfig": {
      "temperature": 1.0, 
      "maxOutputTokens": 8192,
      "thinkingConfig": {
        "includeThoughts": true // เปิดใช้งาน Thinking ในการประเมินวิเคราะห์ข้อสอบ
      }
    }
  };
  
  var options = {
    "method": "post",
    "contentType": "application/json",
    "payload": JSON.stringify(payload),
    "muteHttpExceptions": true
  };

  try {
    var response = UrlFetchApp.fetch(url, options);
    var resJson = JSON.parse(response.getContentText());
    
    if (response.getResponseCode() == 200) {
      var candidate = resJson.candidates && resJson.candidates[0];
      if (candidate && candidate.content && candidate.content.parts) {
        updateAIUsage(apiKeyInfo, apiKeyInfo.model);

        // กรองเอาเฉพาะเนื้อหาคำตอบจริง (ข้ามส่วนที่เป็นกระบวนการคิดหรือ "thought": true)
        var respParts = candidate.content.parts;
        var aiText = "";
        for (var i = 0; i < respParts.length; i++) {
          if (!respParts[i].thought && respParts[i].text) {
            aiText += respParts[i].text;
          }
        }
        if (!aiText.trim()) {
          var finishReason = candidate.finishReason || "UNKNOWN";
          throw new Error("AI ไม่ส่งคำตอบกลับมา (finishReason: " + finishReason + ")");
        }
        return aiText.trim();
      } else {
        var finishReason = (resJson.candidates && resJson.candidates[0] && resJson.candidates[0].finishReason) || "NO_CONTENT";
        throw new Error("Gemini ไม่ส่งเนื้อหากลับมา (finishReason: " + finishReason + ")");
      }
    } else {
      var errorMsg = resJson.error ? resJson.error.message : "Unknown error";
      throw new Error("Gemini Error: " + errorMsg);
    }
  } catch (e) {
    throw new Error("AI Assistant Error: " + e.message);
  }
}


/*
   =========================================
   Student PDF Converter — Phase 1 (Gemini proxy)
   แผน: Idea/active/student-pdf-converter-plan.md
   =========================================
*/

// D13: fallback chain สำรองกรณีทะเบียน AI_Models อ่านไม่ได้ (ปกติ chain มาจาก apiKeyInfo.fallbackModels)
// 2026-08-09: ตัด flash-lite ออกทั้งหมด — คุณภาพแปลงข้อสอบต่ำเกินรับได้
// converter ใช้ full flash เท่านั้น (3.6 → 3.5 → 2.5); การกรองจริงอยู่ใน callGeminiConverter
// เพราะ chain ที่ใช้จริงมาจากทะเบียน AI_Models ซึ่งยังมี lite อยู่ (โมดูลอื่นยังใช้ lite ได้ตามเดิม)
var CONVERTER_FALLBACK_MODELS = ["gemini-3.6-flash", "gemini-3.5-flash", "gemini-2.5-flash"];
// กันชน 6-min execution limit: จำกัดจำนวนครั้งที่ยิง Gemini จริงต่อ 1 POST
var CONVERTER_MAX_ATTEMPTS = 3;

/**
 * Gemini call สำหรับแปลงข้อสอบ (คนละ tuning กับ callGeminiAI ของ chatbot):
 * JSON mode, temperature ต่ำ, maxOutputTokens สูง, ปิด thinking, ไม่มี systemInstruction
 * pdfB64 = base64 ของ PDF ทั้งไฟล์ (batch เดียว) หรือ images = dataURL ต่อหน้า (batch ใหญ่)
 * คืน { raw, finishReason, model } — ฝั่ง client เป็นคน parse (มี recovery logic ครบอยู่แล้ว)
 */
function callGeminiConverter(prompt, apiKeyInfo, pdfB64, images) {
  var chain = (apiKeyInfo.fallbackModels && apiKeyInfo.fallbackModels.length)
    ? apiKeyInfo.fallbackModels : CONVERTER_FALLBACK_MODELS;
  var models = [apiKeyInfo.model].concat(chain)
    .filter(function (m, i, arr) { return m && arr.indexOf(m) === i; })
    // converter ห้ามใช้ flash-lite เด็ดขาด — กรองทั้ง apiKeyInfo.model และ chain จากทะเบียน AI_Models
    // (ทะเบียนคืนโมเดล Active ทุกตัวรวม lite; โมดูลอื่น เช่น askAIExpert/IntelSphere ยังใช้ lite ได้ตามเดิม)
    .filter(function (m) { return !/flash-lite/i.test(m); });
  if (models.length === 0) {
    throw new Error("แปลงไม่สำเร็จ: โควต้าโมเดล flash เต็มแล้ว (ตัวแปลง PDF ไม่ใช้ flash-lite) กรุณาลองใหม่ภายหลัง");
  }

  var attempts = 0;
  var lastErr = "";
  var convTemp = 0.1; // ถูก bump เป็น 0.8 เมื่อเจอ RECITATION — temp ต่ำทำให้ retry ซ้ำผลเดิมเป๊ะ
  for (var mi = 0; mi < models.length; mi++) {
    // ลองแบบปิด thinking ก่อน (thinkingBudget:0) — บางรุ่น reject หรือคืนคำตอบว่าง จึงมี variant ไม่ส่ง thinkingConfig สำรอง
    var variants = [true, false];
    for (var vi = 0; vi < variants.length; vi++) {
      if (attempts >= CONVERTER_MAX_ATTEMPTS) {
        throw new Error("แปลงไม่สำเร็จ (ครบจำนวนครั้งที่ลองได้): " + lastErr + converterErrHint_(lastErr));
      }
      attempts++;
      var res = tryConverterCall_(prompt, apiKeyInfo.key, models[mi], variants[vi], pdfB64, images, convTemp);
      if (res.ok) {
        updateAIUsage(apiKeyInfo, models[mi]); // หักโควต้าโมเดลที่ใช้จริง (อาจเป็น fallback ไม่ใช่ตัวที่เลือกตอนแรก)
        return { raw: res.raw, finishReason: res.finishReason, model: models[mi] };
      }
      lastErr = models[mi] + ": " + res.error;
      if (res.fatal) throw new Error("แปลงไม่สำเร็จ: " + lastErr);
      // RECITATION: ยืนยันจากการรันจริง 2 รอบว่าสลับ thinking variant ไม่ช่วย — bump temp แล้วข้ามไปโมเดลถัดไปเลย
      if (res.recitation) { convTemp = 0.8; break; }
      if (res.quota) handleGemini429_(apiKeyInfo, models[mi], res.body429, res.headers429); // 429: perDay→zero, perMinute/unknown→cooldown
      if (res.nextModel) break; // 429/404 → ข้ามไปโมเดลถัดไปเลย ไม่ต้องลอง variant
      // อื่นๆ (400/คำตอบว่าง) → วนไป variant ไม่ส่ง thinkingConfig; ถ้าหมด variant ก็ตกไปโมเดลถัดไป
    }
  }
  throw new Error("แปลงไม่สำเร็จ (ทุกโมเดลใช้งานไม่ได้): " + lastErr + converterErrHint_(lastErr));
}

// ข้อความช่วยอธิบายให้ผู้ใช้ เมื่อ error สุดท้ายคือ RECITATION (ตัวกรองการคัดลอกเนื้อหาของ Gemini)
function converterErrHint_(msg) {
  return msg.indexOf("RECITATION") >= 0
    ? " — เนื้อหาไปตรงกับตัวกรอง recitation ของ Gemini ลองกดแปลงซ้ำอีกครั้ง หรือแบ่งช่วงหน้าให้เล็กลง"
    : "";
}

// ยิง Gemini 1 ครั้ง — คืน {ok,raw,finishReason} หรือ {ok:false,error,nextModel?,fatal?}
function tryConverterCall_(prompt, apiKey, model, disableThinking, pdfB64, images, temperature) {
  var url = "https://generativelanguage.googleapis.com/v1beta/models/" + model + ":generateContent?key=" + apiKey;

  var parts = [{ "text": prompt }];
  if (pdfB64) {
    parts.push({ "inlineData": { "mimeType": "application/pdf", "data": pdfB64 } });
  }
  if (images && Array.isArray(images)) {
    images.forEach(function (dataUrl) {
      var s = String(dataUrl);
      if (s.indexOf("data:") !== 0) return;
      var comma = s.indexOf(",");
      var mimeMatch = s.match(/^data:(.*?);/);
      if (comma < 0 || !mimeMatch) return;
      parts.push({ "inlineData": { "mimeType": mimeMatch[1], "data": s.substring(comma + 1) } });
    });
  }

  var genConfig = {
    "responseMimeType": "application/json",
    "temperature": (typeof temperature === "number" ? temperature : 0.1),
    "maxOutputTokens": 65536
  };
  if (disableThinking) {
    genConfig.thinkingConfig = { "thinkingBudget": 0 };
  }

  var options = {
    "method": "post",
    "contentType": "application/json",
    "payload": JSON.stringify({ "contents": [{ "parts": parts }], "generationConfig": genConfig }),
    "muteHttpExceptions": true
  };

  try {
    var response = UrlFetchApp.fetch(url, options);
    var code = response.getResponseCode();
    var resJson;
    try { resJson = JSON.parse(response.getContentText()); } catch (pe) { resJson = {}; }

    if (code === 200) {
      var candidate = resJson.candidates && resJson.candidates[0];
      var raw = "";
      if (candidate && candidate.content && candidate.content.parts) {
        var rp = candidate.content.parts;
        for (var i = 0; i < rp.length; i++) {
          if (!rp[i].thought && rp[i].text) raw += rp[i].text;
        }
      }
      if (!raw.trim()) {
        // silent-empty (เจอได้กับ JSON mode + thinkingBudget:0 บางรุ่น) → ให้ caller ลอง variant/โมเดลถัดไป
        var fr = (candidate && candidate.finishReason) || "NO_CONTENT";
        return { ok: false, error: "คำตอบว่าง (finishReason: " + fr + ")", recitation: fr === "RECITATION" };
      }
      return { ok: true, raw: raw, finishReason: (candidate && candidate.finishReason) || "STOP" };
    }

    var msg = (resJson.error && resJson.error.message) || ("HTTP " + code);
    if (code === 429) return { ok: false, error: msg, nextModel: true, quota: true, body429: response.getContentText(), headers429: response.getAllHeaders() }; // 429 → caller แยก perDay/perMinute
    if (code === 404 || code >= 500) return { ok: false, error: msg, nextModel: true }; // ไม่มีโมเดล/overloaded → โมเดลถัดไป
    if (code === 400) return { ok: false, error: msg }; // เช่น reject thinkingConfig → ลอง variant ถัดไป
    return { ok: false, error: msg, fatal: true }; // 401/403 — key ใช้ไม่ได้ เปลี่ยนโมเดลก็ไม่ช่วย
  } catch (e) {
    return { ok: false, error: e.message, fatal: true };
  }
}

/* =========================================================
   บริจาค Gemini API Key เข้า AI_Config pool (converter ดึงผ่าน getAvailableAIKey("Gemini"))
   UI: MDKKUQUIZDATABASE converter panel — ปุ่ม "บริจาค Gemini Key"
   ========================================================= */

// ตรวจ Gemini key แบบ tri-state: 'valid' | 'invalid' | 'unknown'
// - 200/429(quota)/404(model missing แต่ auth ผ่าน) → valid
// - 400/403 ที่เป็น key ปลอม/ถูกเพิกถอน → invalid
// - network/5xx/อื่นๆ → unknown (ไม่บันทึก บอกผู้ใช้ลองใหม่)
function validateGeminiKeyLive(apiKey) {
  try {
    var url = "https://generativelanguage.googleapis.com/v1beta/models/gemini-3.5-flash:generateContent?key=" + encodeURIComponent(apiKey);
    var res = UrlFetchApp.fetch(url, {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify({
        contents: [{ parts: [{ text: "hi" }] }],
        generationConfig: { maxOutputTokens: 1 }
      }),
      muteHttpExceptions: true
    });
    var code = res.getResponseCode();
    if (code === 200 || code === 429 || code === 404) return 'valid'; // auth ผ่าน (throttle/model-missing ไม่ใช่ปัญหาความถูกต้องของ key)
    var body = res.getContentText() || "";
    if ((code === 400 || code === 403) &&
        /API_KEY_INVALID|API key not valid|PERMISSION_DENIED|API_KEY_SERVICE_BLOCKED|API_KEY_HTTP_REFERRER_BLOCKED/i.test(body)) {
      return 'invalid';
    }
    return 'unknown';
  } catch (e) {
    return 'unknown';
  }
}

// บันทึก key ลง AI_Config โครงใหม่ (idempotent by API_Key) — เรียกภายใต้ localized lock จาก router
// prefill โควต้าเต็มทุกโมเดลจากทะเบียน AI_Models (แบบเดียวกับ seedIntelSphereKey)
function seedGeminiKey(apiKey, donorName) {
  if (!apiKey) {
    return ContentService.createTextOutput(JSON.stringify({
      result: 'error', message: 'กรุณากรอก API Key'
    })).setMimeType(ContentService.MimeType.JSON);
  }
  apiKey = String(apiKey).trim();

  var verdict = validateGeminiKeyLive(apiKey);
  if (verdict === 'invalid') {
    return ContentService.createTextOutput(JSON.stringify({
      result: 'error', message: 'API Key ไม่ถูกต้องหรือถูกเพิกถอนแล้ว กรุณาตรวจสอบจาก aistudio.google.com/apikey อีกครั้ง'
    })).setMimeType(ContentService.MimeType.JSON);
  }
  if (verdict !== 'valid') {
    return ContentService.createTextOutput(JSON.stringify({
      result: 'error', message: 'ตรวจสอบ Key ไม่ได้ในขณะนี้ (Google อาจขัดข้องชั่วคราว) กรุณาลองใหม่ภายหลัง'
    })).setMimeType(ContentService.MimeType.JSON);
  }

  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = getAIConfigSheet_(ss); // auto-สร้าง/migrate โครงใหม่
  var models = getAIModelRegistry_(ss);

  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var colStatus = headers.indexOf("Status");
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === apiKey) {
      sheet.getRange(i + 1, colStatus + 1).setValue("Active"); // ปลุก key เดิมที่อาจถูกปิดไว้
      SpreadsheetApp.flush();
      return ContentService.createTextOutput(JSON.stringify({
        result: 'success', updatedExisting: true,
        message: 'Key นี้มีอยู่ในระบบแล้ว — เปิดใช้งานอีกครั้งเรียบร้อย ขอบคุณครับ'
      })).setMimeType(ContentService.MimeType.JSON);
    }
  }

  // แถวใหม่ตาม header order: prefill โควต้าเต็มทุกโมเดล
  var rowValues = new Array(headers.length).fill("");
  function setColV(name, val) { var c = headers.indexOf(name); if (c >= 0) rowValues[c] = val; }
  setColV("API_Key", apiKey);
  setColV("Donor_Name", donorName || "");
  setColV("Status", "Active");
  setColV("Last_Used", new Date());
  setColV("Last_Reset_Date", Utilities.formatDate(new Date(), "Asia/Bangkok", "yyyy-MM-dd"));
  models.forEach(function(m) { setColV(m.model + "_Remaining", m.limit); });
  sheet.appendRow(rowValues);
  SpreadsheetApp.flush();
  return ContentService.createTextOutput(JSON.stringify({
    result: 'success', appended: true,
    message: 'ขอบคุณสำหรับการบริจาค! Key ผ่านการตรวจสอบและพร้อมใช้งานแล้ว'
  })).setMimeType(ContentService.MimeType.JSON);
}


