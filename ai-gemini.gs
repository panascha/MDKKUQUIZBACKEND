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
var AI_MODELS_DEFAULTS = [
  // [Model, RPD_Limit, Priority, Status, Notes]
  ["gemini-3.5-flash",      20,  1, "Active",   "text-out หลัก"],
  ["gemini-3-flash",        20,  2, "Active",   ""],
  ["gemini-2.5-flash",      20,  3, "Active",   ""],
  ["gemini-3.1-flash-lite", 500, 4, "Active",   "RPD สูงสุดใน free tier"],
  ["gemini-2.5-flash-lite", 20,  5, "Active",   ""],
  ["gemini-3.1-pro",        0,  90, "Disabled", "free tier RPD = 0"],
  ["gemini-2.5-pro",        0,  91, "Disabled", "free tier RPD = 0"]
];

// Retry helper — เอกสารใหญ่ เจอ "บริการ สเปรดชีต หมดเวลา" เป็นพักๆ ระหว่าง write ติดกัน
function aiSheetRetry_(fn) {
  var lastErr;
  for (var a = 0; a < 3; a++) {
    try { return fn(); } catch (e) { lastErr = e; Utilities.sleep(1500 * (a + 1)); }
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
function getAvailableAIKey(provider, preferredModel, reserveCount) {
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
    if (remCol >= 0) sheet.getRange(apiKeyInfo.index, remCol + 1).setValue(0);
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
      "maxOutputTokens": 2048,
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
var CONVERTER_FALLBACK_MODELS = ["gemini-3.5-flash", "gemini-2.5-flash", "gemini-3.1-flash-lite"];
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
    .filter(function (m, i, arr) { return m && arr.indexOf(m) === i; });

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
      if (res.quota) markModelExhausted_(apiKeyInfo, models[mi]); // 429 → โมเดลนี้หมดโควต้าวันนี้สำหรับ key นี้
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
    if (code === 429 || code === 404 || code >= 500) return { ok: false, error: msg, nextModel: true, quota: code === 429 }; // quota/ไม่มีโมเดล/overloaded → โมเดลถัดไป
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


