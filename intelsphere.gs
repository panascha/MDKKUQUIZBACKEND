// จำแนก provider จาก prefix ของ model ID — ทนต่อโมเดลใหม่ในตระกูลเดิมที่ KKU เพิ่มภายหลัง
function inferProviderFromModel(modelId) {
  if (/^claude-/i.test(modelId))                          return "Claude";
  if (/^deepseek-/i.test(modelId))                         return "Deepseek";
  if (/^gemini-/i.test(modelId))                           return "Gemini";
  if (/^llama-/i.test(modelId))                            return "Meta";
  if (/^minimax-/i.test(modelId))                          return "MiniMax";
  if (/^(mistral-|codestral-|devstral-)/i.test(modelId))   return "Mistral";
  if (/^nova-/i.test(modelId))                             return "Nova";
  if (/^gpt-/i.test(modelId))                              return "OpenAI";
  if (/^qwen/i.test(modelId))                              return "Qwen";
  if (/^grok-/i.test(modelId))                             return "xAI";
  if (/^sonar-/i.test(modelId))                            return "Perplexity"; // classified for display only — still excluded from rotation
  return null; // unrecognized prefix — log it, don't silently drop
}

// หา key ของแถว Active แถวแรกเพื่อใช้ fetch catalog (catalog เป็น account-agnostic)
function getAnyActiveIntelSphereKeyForCatalogFetch() {
  var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(INTELSPHERE_SHEET_NAME);
  if (!sheet) return null;
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return null;
  var headers = data[0];
  var colKey = headers.indexOf("API_Key");
  var colStatus = headers.indexOf("Status");
  if (colKey < 0 || colStatus < 0) return null;
  for (var i = 1; i < data.length; i++) {
    if (data[i][colStatus] === "Active" && data[i][colKey]) return data[i][colKey];
  }
  return null;
}

// Live catalog จาก GET /models, cache 6 ชม. — fallback เป็น PROVIDER_MODELS_FALLBACK เมื่อ fetch ล้มเหลว
function getIntelSphereModelCatalog() {
  var cache = CacheService.getScriptCache();
  var cached = cache.get("intelsphere_catalog");
  if (cached) return JSON.parse(cached);

  var catalog = {};
  var unknownModels = [];

  try {
    var anyKey = getAnyActiveIntelSphereKeyForCatalogFetch();
    if (!anyKey) throw new Error("no active key available to fetch catalog");

    var response = UrlFetchApp.fetch("https://gen.ai.kku.ac.th/api/v1/models", {
      method: "get",
      headers: { "Authorization": "Bearer " + anyKey },
      muteHttpExceptions: true
    });
    if (response.getResponseCode() !== 200) throw new Error("catalog fetch HTTP " + response.getResponseCode());

    var body = JSON.parse(response.getContentText());
    body.data.forEach(function(m) {
      // ยืนยันกับ live response 2026-07-03: model ID อยู่ที่ m.id; m.owned_by = ชื่อ provider แบบ display
      // ("Meta AI","Nova (AWS)") ที่ไม่ตรง internal key — จึงจำแนกด้วย prefix ของ m.id แทน
      var modelId = m.id;
      var provider = inferProviderFromModel(modelId);
      if (!provider) { unknownModels.push(modelId); return; }
      if (!catalog[provider]) catalog[provider] = [];
      catalog[provider].push(modelId);
    });

    if (unknownModels.length > 0) {
      console.warn("[IntelSphere] Unclassified model IDs: " + unknownModels.join(", "));
    }
    if (Object.keys(catalog).length === 0) throw new Error("catalog parsed but empty — check response shape assumption");

  } catch (e) {
    console.warn("[IntelSphere] Live catalog fetch failed (" + e.message + ") — using fallback list");
    catalog = PROVIDER_MODELS_FALLBACK;
  }

  cache.put("intelsphere_catalog", JSON.stringify(catalog), 21600);
  return catalog;
}

// ตรวจ key กับ IntelSphere จริงก่อนบันทึก — คืน 'valid' | 'invalid' | 'unavailable'
// GET /models ใช้ไม่ได้ (endpoint สาธารณะ ตอบ 200 แม้ไม่มี auth — ยืนยัน 2026-07-03) ต้องยิง chat 1 token แทน
// 401 "reached daily limit" = key จริงแต่โควต้าวันนี้หมด → ถือว่า valid (พรุ่งนี้ใช้ได้)
function validateIntelSphereKeyLive(apiKey) {
  var model = PROVIDER_MODEL_MAP["Deepseek"];
  try {
    var catalog = getIntelSphereModelCatalog();
    for (var p in catalog) {
      if (catalog[p] && catalog[p].length > 0) { model = catalog[p][0]; break; }
    }
  } catch (e) { /* ใช้ flagship fallback */ }

  try {
    var resp = UrlFetchApp.fetch(INTELSPHERE_ENDPOINT, {
      method: "post",
      contentType: "application/json",
      headers: { "Authorization": "Bearer " + apiKey },
      payload: JSON.stringify({ model: model, messages: [{ role: "user", content: "hi" }], max_tokens: 1 }),
      muteHttpExceptions: true
    });
    var code = resp.getResponseCode();
    if (code === 200) return 'valid';
    if (code === 401) {
      var errText = "";
      try { errText = JSON.parse(resp.getContentText()).error || resp.getContentText(); }
      catch (e2) { errText = resp.getContentText(); }
      if (errText.indexOf("reached daily limit") >= 0) return 'valid';
      if (errText.indexOf("Invalid model") >= 0) {
        CacheService.getScriptCache().remove("intelsphere_catalog"); // catalog drift — ให้รอบหน้าดึงใหม่
        return 'unavailable';
      }
      return 'invalid';
    }
    return 'unavailable';
  } catch (e3) {
    return 'unavailable';
  }
}

// ── Daily key sweep: probe key ทุกใบใน IntelSphere_Keys ด้วย validateIntelSphereKeyLive ──
// Active → probe 'invalid' → mark Invalid (จับ key ถูกเพิกถอนก่อนชน traffic จริง)
// Invalid → probe 'valid' → คืน Active (แก้ false positive จาก 401 ตอน traffic)
// 'unavailable' = endpoint/catalog ขัดข้อง ไม่ใช่ตัว key → ไม่แตะสถานะ; สถานะ manual อื่นๆ ไม่ยุ่ง
function runIntelSphereKeySweep() {
  var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(INTELSPHERE_SHEET_NAME);
  if (!sheet) { console.warn("[keySweep] ไม่พบ sheet " + INTELSPHERE_SHEET_NAME); return { checked: 0, changed: 0 }; }

  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var colKey = headers.indexOf("API_Key");
  var colStatus = headers.indexOf("Status");
  var colDonor = headers.indexOf("Donor_Name");

  var checked = 0, changed = 0;
  for (var i = 1; i < data.length; i++) {
    var key = String(data[i][colKey] || "").trim();
    var status = data[i][colStatus];
    if (!key) continue;
    if (status !== "Active" && status !== "Invalid") continue;

    var verdict = validateIntelSphereKeyLive(key);
    checked++;
    if (verdict === 'unavailable') continue;

    var want = (verdict === 'valid') ? "Active" : "Invalid";
    if (want !== status) {
      sheet.getRange(i + 1, colStatus + 1).setValue(want);
      changed++;
      console.warn("[keySweep] row " + (i + 1) + " (" + (colDonor >= 0 ? data[i][colDonor] : "?") + "): " + status + " → " + want);
    }
  }
  SpreadsheetApp.flush();
  console.log("[keySweep] checked=" + checked + " changed=" + changed);
  return { checked: checked, changed: changed };
}

// ติดตั้ง time-driven trigger รันทุกวัน ~ตี 5 (idempotent) — เว้นตี 3-4 ให้ batch jobs เดิม
function installKeySweepTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'runIntelSphereKeySweep') ScriptApp.deleteTrigger(triggers[i]);
  }
  ScriptApp.newTrigger('runIntelSphereKeySweep').timeBased().everyDays(1).atHour(5).create();
  return 'installed';
}

// เลือก (key, provider) ที่ยังมีโควต้า — weighted-random ตาม remaining tokens
// (กับ key เดียวในระบบ พฤติกรรมเทียบเท่า top-down scan; รองรับหลาย key อัตโนมัติเมื่อมีผู้บริจาคเพิ่ม)
// reservedKeySet: ชุดของ API_Key ที่ห้ามใช้ (จองไว้ให้ public) — {} = ไม่มี reserve
function getActiveIntelSphereKey(requestedProvider, reservedKeySet) {
  reservedKeySet = reservedKeySet || {};
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(INTELSPHERE_SHEET_NAME);
  if (!sheet) throw new Error("ไม่พบ sheet IntelSphere_Keys — ตรวจสอบชื่อ sheet");

  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var colKey = headers.indexOf("API_Key");
  var colStatus = headers.indexOf("Status");
  var colLastReset = headers.indexOf("Last_Reset_Date");

  var tz = "Asia/Bangkok"; // hardcoded — อย่าใช้ timezone ของ account เจ้าของ script (อาจเป็น UTC ทำให้ reset ช้า 7 ชม.)
  var todayStr = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd");

  // 1. Per-row daily reset แล้วรวบรวมทุก (key, provider) pair ที่โควต้าเหลือเกิน floor
  var candidates = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (row[colStatus] !== "Active") continue;

    var lastResetStr = row[colLastReset]
      ? Utilities.formatDate(new Date(row[colLastReset]), tz, "yyyy-MM-dd") : "";
    if (lastResetStr !== todayStr) {
      for (var r = 0; r < INTELSPHERE_PROVIDER_PRIORITY.length; r++) {
        var pr = INTELSPHERE_PROVIDER_PRIORITY[r];
        var rc = headers.indexOf(pr + "_Remaining");
        if (rc >= 0) sheet.getRange(i + 1, rc + 1).setValue(INTELSPHERE_LIMITS[pr]);
      }
      sheet.getRange(i + 1, colLastReset + 1).setValue(todayStr);
      SpreadsheetApp.flush();
      data = sheet.getDataRange().getValues();
      row = data[i];
    }

    var keyId = String(row[colKey] || "").trim();
    if (reservedKeySet[keyId]) continue; // จองไว้ให้ public — agentQuery ห้ามใช้

    for (var j = 0; j < INTELSPHERE_PROVIDER_PRIORITY.length; j++) {
      var provider = INTELSPHERE_PROVIDER_PRIORITY[j];
      var remCol = headers.indexOf(provider + "_Remaining");
      if (remCol < 0) continue; // header หาย/พิมพ์ผิด — provider นั้นหลุดจาก rotation เงียบๆ ตรวจตอน seed sheet
      var remaining = Number(row[remCol]);
      var floor = INTELSPHERE_LIMITS[provider] * INTELSPHERE_QUOTA_FLOOR;
      if (isNaN(remaining) || remaining <= floor) continue;
      candidates.push({
        key: row[colKey], provider: provider, rowIndex: i + 1,
        remaining: remaining, remainingCol: remCol + 1
      });
    }
  }
  if (candidates.length === 0) return null; // ไม่มี key ที่ใช้ได้เลย

  // 2. เคารพ provider ที่นิสิตเลือกถ้ายังมีโควต้า ไม่งั้น draw จาก pool ทั้งหมด
  var pool = candidates;
  if (requestedProvider) {
    var preferred = candidates.filter(function(c) { return c.provider === requestedProvider; });
    if (preferred.length > 0) pool = preferred;
  }

  // 3. Weighted-random pick, weight = remaining tokens
  var total = pool.reduce(function(s, c) { return s + c.remaining; }, 0);
  var dart = Math.random() * total;
  for (var k = 0; k < pool.length; k++) {
    dart -= pool[k].remaining;
    if (dart <= 0) return pool[k];
  }
  return pool[pool.length - 1]; // float-rounding safety
}

// ยิงคำถามไป IntelSphere พร้อม rotation: quota-exhausted → provider ถัดไป, invalid key → key ถัดไป
function executeChatbotQuery(prompt, requestedModel, attempt, maxTokens) {
  attempt = attempt || 1;
  var maxAttempts = INTELSPHERE_PROVIDER_PRIORITY.length; // exhaust รอบ rotation เต็มก่อนยอมแพ้
  if (attempt > maxAttempts) throw new Error("ขออภัย ระบบ AI ไม่พร้อมใช้งานชั่วคราว กรุณาลองใหม่ในอีกสักครู่");

  var requestedProvider = inferProviderFromModel(requestedModel);
  if (!requestedProvider) throw new Error("ไม่รู้จักโมเดลนี้ กรุณาเลือกโมเดลใหม่จากรายการ");
  if (requestedProvider === "Perplexity") throw new Error("โมเดลนี้ยังไม่พร้อมใช้งานในระบบ กรุณาเลือกโมเดลอื่น");

  var keyObj = getActiveIntelSphereKey(requestedProvider);

  if (!keyObj) {
    // Legacy fallback: ใช้ Gemini pool เดิมถ้า IntelSphere หมดทุก key
    if (typeof getAvailableAIKey === "function" && typeof callGeminiAI === "function") {
      var fallbackKey = getAvailableAIKey("Gemini");
      if (!fallbackKey) throw new Error("โควต้า AI หมดแล้วสำหรับวันนี้ กรุณารอจนถึงเที่ยงคืนเพื่อรีเซ็ตโควต้า");
      return { content: callGeminiAI(prompt, fallbackKey, null), servedModel: "gemini (legacy pool)", switched: true };
    }
    throw new Error("โควต้า AI หมดแล้วสำหรับวันนี้ กรุณารอจนถึงเที่ยงคืนเพื่อรีเซ็ตโควต้า");
  }

  // ใช้โมเดลที่นิสิตเลือกเป๊ะๆ ถ้า rotation ยังอยู่ provider เดิม
  // ถ้า rotation ย้าย provider เราไม่รู้ว่านิสิตอยากได้โมเดลไหนของเจ้านั้น — ใช้ flagship
  var actualModel = (keyObj.provider === requestedProvider) ? requestedModel : PROVIDER_MODEL_MAP[keyObj.provider];
  var switched = (actualModel !== requestedModel);

  // maxTokens optional (default 2000)
  var payload = { model: actualModel, messages: [{ role: "user", content: prompt }], max_tokens: maxTokens || 2000, temperature: 0.3 };
  var options = {
    method: "post", contentType: "application/json",
    headers: { "Authorization": "Bearer " + keyObj.key },
    payload: JSON.stringify(payload), muteHttpExceptions: true
  };

  var response = UrlFetchApp.fetch(INTELSPHERE_ENDPOINT, options);
  var code = response.getResponseCode();

  if (code === 200) {
    var body;
    try { body = JSON.parse(response.getContentText()); }
    catch (parseErr) { throw new Error("เกิดข้อผิดพลาดในการอ่านคำตอบจาก AI API กรุณาลองใหม่อีกครั้ง"); }

    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName(INTELSPHERE_SHEET_NAME);
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    sheet.getRange(keyObj.rowIndex, headers.indexOf("Last_Used") + 1).setValue(new Date());
    if (body.model_quota && typeof body.model_quota.daily_remaining_tokens === "number") {
      sheet.getRange(keyObj.rowIndex, keyObj.remainingCol).setValue(body.model_quota.daily_remaining_tokens);
    }
    return { content: body.choices[0].message.content, servedModel: actualModel, switched: switched };
  }

  if (code === 401) {
    var errText = "";
    try { errText = JSON.parse(response.getContentText()).error || response.getContentText(); }
    catch (e) { errText = response.getContentText(); }

    var ss2 = SpreadsheetApp.openById(SHEET_ID);
    var sheet2 = ss2.getSheetByName(INTELSPHERE_SHEET_NAME);
    var headers2 = sheet2.getRange(1, 1, 1, sheet2.getLastColumn()).getValues()[0];

    if (errText.indexOf("reached daily limit") >= 0) {
      // Quota หมดของ provider นี้ — zero column แล้ว rotate ต่อ (ไม่แตะ Status)
      sheet2.getRange(keyObj.rowIndex, keyObj.remainingCol).setValue(0);
      SpreadsheetApp.flush();
      return executeChatbotQuery(prompt, requestedModel, attempt + 1, maxTokens);
    }

    if (errText.indexOf("Invalid model") >= 0) {
      // Catalog drift — bust cache แล้ว retry ด้วย flagship ของ provider เดิม (ไม่แตะ donor key)
      CacheService.getScriptCache().remove("intelsphere_catalog");
      console.warn("[IntelSphere] Invalid model at request time: " + actualModel + " — catalog cache cleared");
      if (actualModel !== PROVIDER_MODEL_MAP[keyObj.provider]) {
        return executeChatbotQuery(prompt, PROVIDER_MODEL_MAP[keyObj.provider], attempt + 1, maxTokens);
      }
      throw new Error("เกิดข้อผิดพลาดในการตั้งค่าโมเดล AI กรุณาแจ้งทีม IT");
    }

    // Key เสีย/ถูกเพิกถอนจริงๆ
    sheet2.getRange(keyObj.rowIndex, headers2.indexOf("Status") + 1).setValue("Invalid");
    SpreadsheetApp.flush();
    return executeChatbotQuery(prompt, requestedModel, attempt + 1, maxTokens);
  }

  if (code === 400) throw new Error("เกิดข้อผิดพลาดในการส่งคำขอ กรุณาลองใหม่อีกครั้ง");

  // 500/503/gateway — ไม่พยายาม JSON.parse body ที่อาจไม่ใช่ JSON
  throw new Error("เกิดข้อผิดพลาดจาก AI API (HTTP " + code + ") กรุณาลองใหม่อีกครั้ง");
}

// ==== agentQuery: เสิร์ฟ claude-kkuintelsphere-router (Claude Code fallback proxy) ====

var AGENT_QUERY_OWNER_SECRET_HASH_PROP = 'AGENT_QUERY_OWNER_SECRET_HASH';

// Non-expiring owner auth for agentQuery. Only the SHA-256 hash is stored in
// Script Properties — the plaintext secret lives only in the proxy owner's
// localhost config.local.json. Returns true iff the presented secret matches.
function verifyAgentQueryOwnerSecret(secret) {
  if (!secret) return false;
  var stored = PropertiesService.getScriptProperties().getProperty(AGENT_QUERY_OWNER_SECRET_HASH_PROP);
  if (!stored) return false;
  var got = hashPasswordInternal(String(secret));
  // constant-time-ish compare (avoid early-exit length/char leaks)
  if (got.length !== stored.length) return false;
  var diff = 0;
  for (var i = 0; i < got.length; i++) diff |= got.charCodeAt(i) ^ stored.charCodeAt(i);
  return diff === 0;
}

// One-off setup: run this once in the Apps Script editor. It mints a random
// secret, stores only its hash, and logs the plaintext ONCE — copy that into
// the proxy's config.local.json (gas.ownerSecret), then clear the execution log.
// Re-running rotates the secret (invalidates the old one).
function generateAgentQueryOwnerSecret() {
  var secret = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
  PropertiesService.getScriptProperties().setProperty(
    AGENT_QUERY_OWNER_SECRET_HASH_PROP, hashPasswordInternal(secret)
  );
  Logger.log('AGENT_QUERY OWNER SECRET (paste into config.local.json gas.ownerSecret; shown once):\n' + secret);
  return secret;
}

// Priority แยกจาก INTELSPHERE_PROVIDER_PRIORITY โดยเจตนา — agent ต้องการโมเดลแรงสุดก่อน ไม่ใช่ถูกสุดก่อน
var AGENT_QUERY_PROVIDER_PRIORITY = ["Claude", "Deepseek", "Qwen", "OpenAI"];
var AGENT_QUERY_MAX_OUTPUT_TOKENS = 8192; // Claude Code ส่ง max_tokens สูง (เช่น 32000) — clamp กัน 400 จาก provider ที่ cap ต่ำกว่า
// Context window โดยประมาณ (tokens) ของ flagship ต่อ provider — ตัวเลข conservative, ปรับเมื่อ KKU เปลี่ยนรุ่น
var AGENT_PROVIDER_CONTEXT = { "Claude": 200000, "Deepseek": 128000, "Qwen": 131072, "OpenAI": 128000 };
// จอง key ไว้สำหรับผู้ใช้สาธารณะ — agentQuery (owner proxy) จะไม่ใช้ key ที่มีโควต้าคงเหลือรวมมากที่สุด N อันดับแรก
// (key = 1 API_Key ใช้ได้ทุก provider → reserve ทั้ง key ไม่ใช่แยก provider)
var AGENT_QUERY_KEY_RESERVE_COUNT = 2;

// รวมโควต้าคงเหลือรายวันต่อ provider (ทุก donor key ที่ Active) — ใช้จัดลำดับ chain แบบ load-balance
// นับเฉพาะ key ที่เกิน quota floor (เกณฑ์เดียวกับ getActiveIntelSphereKey) — ต่ำกว่า floor คือ serve ไม่ได้จริง
// reserveCount > 0 → ข้าม key ที่มี totalRemaining มากที่สุด N อันดับแรก (จองไว้ให้ public) — ใช้เฉพาะ agentQuery
// ข้อจำกัดที่ยอมรับ: อ่านค่าก่อน daily reset (reset เกิดใน getActiveIntelSphereKey ทีหลัง) —
// request แรกของวันอาจเรียงด้วยค่าค้างของเมื่อวาน แล้วหายเองใน request ถัดไป
function getIntelSphereQuotaTotals(reserveCount) {
  reserveCount = reserveCount || 0;
  var totals = {};
  var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(INTELSPHERE_SHEET_NAME);
  if (!sheet) return totals;
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var colStatus = headers.indexOf("Status");
  var colKey = headers.indexOf("API_Key");
  // สร้าง list ของ (apiKey, rowIndex, {provider: remaining}) — ไว้คำนวณ reserve
  var keySnapshots = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i][colStatus] !== "Active") continue;
    var snapshot = { apiKey: String(data[i][colKey] || "").trim(), totalRemaining: 0 };
    for (var j = 0; j < AGENT_QUERY_PROVIDER_PRIORITY.length; j++) {
      var p = AGENT_QUERY_PROVIDER_PRIORITY[j];
      var col = headers.indexOf(p + "_Remaining");
      if (col < 0) continue;
      var rem = Number(data[i][col]);
      var floor = INTELSPHERE_LIMITS[p] * INTELSPHERE_QUOTA_FLOOR;
      if (!isNaN(rem) && rem > floor) {
        snapshot.totalRemaining += rem;
        totals[p] = (totals[p] || 0) + rem;
      }
    }
    if (snapshot.apiKey) keySnapshots.push(snapshot);
  }
  // จอง: เรียง key ตาม totalRemaining มาก→น้อย แล้วหักเฉพาะ top-N ออก (reserveCount = AGENT_QUERY_KEY_RESERVE_COUNT)
  if (reserveCount > 0 && keySnapshots.length > 0) {
    keySnapshots.sort(function(a, b) { return b.totalRemaining - a.totalRemaining; });
    var reservedKeys = {};
    for (var r = 0; r < reserveCount && r < keySnapshots.length; r++) {
      reservedKeys[keySnapshots[r].apiKey] = true;
    }
    // สร้างชุดที่สองของ keySnapshots สำหรัับอ่านค่า remaining แบบรวม reserve ด้วย (key level)
    // อันนี้คือ "active non-reserved totals" — เอาไปให้ agentQuery orderProviders
    // แต่เรายังต้องรู้ reserve key ใน getActiveIntelSphereKey → เลย refactor แยกฟังก์ชัน
    var nrTotals = {};
    for (var i = 1; i < data.length; i++) {
      if (data[i][colStatus] !== "Active") continue;
      if (reservedKeys[String(data[i][colKey] || "").trim()]) continue;
      for (var j = 0; j < AGENT_QUERY_PROVIDER_PRIORITY.length; j++) {
        var ap = AGENT_QUERY_PROVIDER_PRIORITY[j];
        var acol = headers.indexOf(ap + "_Remaining");
        if (acol < 0) continue;
        var arem = Number(data[i][acol]);
        var afloor = INTELSPHERE_LIMITS[ap] * INTELSPHERE_QUOTA_FLOOR;
        if (!isNaN(arem) && arem > afloor) nrTotals[ap] = (nrTotals[ap] || 0) + arem;
      }
    }
    return { totals: nrTotals, reservedKeySet: reservedKeys };
  }
  return { totals: totals, reservedKeySet: {} };
}

// จัดลำดับ provider ต่อ request: (1) ตัด provider ที่ context window ไม่พอ (est input + output budget),
// (2) haiku-tier → เจ้าถูกก่อน เก็บ Claude ไว้ท้าย, sonnet/opus-tier → Claude ก่อน,
// (3) เจ้าที่ไม่ใช่ Claude เรียงตามโควต้าคงเหลือมาก→น้อย (กระจาย load ไม่ drain เจ้าเดียว)
// คืน [] ได้เมื่อ input ใหญ่เกินทุกเจ้า (est >~190k) — executeAgentQuery จะข้ามไป Gemini (1M window) เอง
function orderAgentProviders(request, quotaTotals) {
  var estTokens = Math.ceil(JSON.stringify(request.messages || []).length / 4)
                + Math.ceil(JSON.stringify(request.tools || []).length / 4);
  var need = estTokens + AGENT_QUERY_MAX_OUTPUT_TOKENS;
  var eligible = AGENT_QUERY_PROVIDER_PRIORITY.filter(function(p) {
    // provider ที่ไม่มีใน map → ใช้ 128k conservative แทนการหลุด chain เงียบๆ (NaN filter)
    return need <= (AGENT_PROVIDER_CONTEXT[p] || 128000) * 0.95;
  });
  var hasClaude = eligible.indexOf("Claude") >= 0;
  var rest = eligible.filter(function(p) { return p !== "Claude"; }).sort(function(a, b) {
    return (quotaTotals[b] || 0) - (quotaTotals[a] || 0);
  });
  var order = /haiku/i.test(request.model || "")
    ? rest.concat(hasClaude ? ["Claude"] : [])
    : (hasClaude ? ["Claude"] : []).concat(rest);
  console.log("[agentQuery] est~" + estTokens + " tokens, model=" + (request.model || "?")
    + " → order: " + (order.join(" → ") || "(none fit — Gemini only)"));
  return order;
}
// Endpoint OpenAI-compat ของ Google — ใช้แทน callGeminiAI เพราะรับ payload พร้อม tools ได้ตรงๆ (agentic ขาด tools ไม่ได้)
var GEMINI_OPENAI_COMPAT_ENDPOINT = "https://generativelanguage.googleapis.com/v1beta/openai/chat/completions";

// เลือกโมเดลต่อ provider: Claude ใช้โมเดลที่ proxy ขอมาถ้ามีใน catalog (Claude Code ส่งชื่อรุ่นจริงเช่น claude-sonnet-5), ที่เหลือใช้ flagship
function pickAgentModel(provider, requestedModel) {
  if (provider === "Claude" && requestedModel) {
    var claudeModels = getIntelSphereModelCatalog()["Claude"] || [];
    if (claudeModels.indexOf(requestedModel) >= 0) return requestedModel;
  }
  return PROVIDER_MODEL_MAP[provider];
}

// วิ่ง priority chain ตามลำดับจาก orderAgentProviders (per-request: context-fit + tier + quota balance)
// → personal Gemini pool → throw (terminal, ไม่ retry-loop)
function executeAgentQuery(request) {
  var skip = {};
  var attempts = 0;
  var qSnapshot = getIntelSphereQuotaTotals(AGENT_QUERY_KEY_RESERVE_COUNT);
  var order = orderAgentProviders(request, qSnapshot.totals);
  var reservedKeySet = qSnapshot.reservedKeySet;
  var maxAttempts = order.length * 2; // เผื่อหลาย donor key ต่อ provider

  while (attempts < maxAttempts) {
    attempts++;

    var keyObj = null, provider = null;
    for (var i = 0; i < order.length; i++) {
      var p = order[i];
      if (skip[p]) continue;
      var candidate = getActiveIntelSphereKey(p, reservedKeySet);
      // getActiveIntelSphereKey draw provider อื่นเมื่อ provider ที่ขอไม่มีโควต้า — ตีความเป็น "หมด" แล้วไล่ตัวถัดไปเอง
      if (candidate && candidate.provider === p) { keyObj = candidate; provider = p; break; }
      skip[p] = true;
    }
    if (!keyObj) break; // IntelSphere หมดทั้ง chain → Gemini ส่วนตัวด้านล่าง

    var payload = JSON.parse(JSON.stringify(request));
    payload.model = pickAgentModel(provider, request.model);
    delete payload.stream;
    if (!payload.max_tokens || payload.max_tokens > AGENT_QUERY_MAX_OUTPUT_TOKENS) {
      payload.max_tokens = AGENT_QUERY_MAX_OUTPUT_TOKENS;
    }

    var response = UrlFetchApp.fetch(INTELSPHERE_ENDPOINT, {
      method: "post", contentType: "application/json",
      headers: { "Authorization": "Bearer " + keyObj.key },
      payload: JSON.stringify(payload), muteHttpExceptions: true
    });
    var code = response.getResponseCode();

    if (code === 200) {
      var body = JSON.parse(response.getContentText());
      var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(INTELSPHERE_SHEET_NAME);
      var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      sheet.getRange(keyObj.rowIndex, headers.indexOf("Last_Used") + 1).setValue(new Date());
      if (body.model_quota && typeof body.model_quota.daily_remaining_tokens === "number") {
        sheet.getRange(keyObj.rowIndex, keyObj.remainingCol).setValue(body.model_quota.daily_remaining_tokens);
      }
      return { provider: "intelsphere:" + provider, completion: body };
    }

    if (code === 401) {
      var errText = "";
      try { errText = JSON.parse(response.getContentText()).error || response.getContentText(); }
      catch (e401) { errText = response.getContentText(); }
      var sheet2 = SpreadsheetApp.openById(SHEET_ID).getSheetByName(INTELSPHERE_SHEET_NAME);
      var headers2 = sheet2.getRange(1, 1, 1, sheet2.getLastColumn()).getValues()[0];

      if (errText.indexOf("reached daily limit") >= 0) {
        // โควต้า provider นี้ของ key นี้หมด — zero column แล้ววนใหม่ (donor key อื่นของ provider เดิมยังมีสิทธิ์)
        sheet2.getRange(keyObj.rowIndex, keyObj.remainingCol).setValue(0);
        SpreadsheetApp.flush();
        continue;
      }
      if (errText.indexOf("Invalid model") >= 0) {
        CacheService.getScriptCache().remove("intelsphere_catalog");
        console.warn("[agentQuery] Invalid model " + payload.model + " — catalog cache cleared, skipping " + provider);
        skip[provider] = true;
        continue;
      }
      // Key เสีย/ถูกเพิกถอน
      sheet2.getRange(keyObj.rowIndex, headers2.indexOf("Status") + 1).setValue("Invalid");
      SpreadsheetApp.flush();
      continue;
    }

    // 4xx/5xx อื่น — ปัญหาฝั่ง provider/payload ไม่ใช่ตัว key: ข้าม provider นี้ ไม่แตะสถานะ donor key
    console.warn("[agentQuery] " + provider + " HTTP " + code + ": " + String(response.getContentText()).slice(0, 200));
    skip[provider] = true;
  }

  // ---- Tier สุดท้าย: personal Gemini pool (AI_Config sheet เดิม — rotation/daily-reset ในตัว) ----
  var geminiKey = getAvailableAIKey("Gemini");
  if (geminiKey) {
    var gPayload = JSON.parse(JSON.stringify(request));
    gPayload.model = geminiKey.model || "gemini-2.5-flash";
    delete gPayload.stream;
    var gResp = UrlFetchApp.fetch(GEMINI_OPENAI_COMPAT_ENDPOINT, {
      method: "post", contentType: "application/json",
      headers: { "Authorization": "Bearer " + geminiKey.key },
      payload: JSON.stringify(gPayload), muteHttpExceptions: true
    });
    if (gResp.getResponseCode() === 200) {
      updateAIUsage(geminiKey.index, geminiKey.usage);
      return { provider: "gemini:" + gPayload.model, completion: JSON.parse(gResp.getContentText()) };
    }
    console.warn("[agentQuery] Gemini HTTP " + gResp.getResponseCode() + ": " + String(gResp.getContentText()).slice(0, 200));
  }

  // Terminal failure — ตั้งใจให้ fail ทันที proxy ฝั่ง client จะแสดง error ชัดๆ ไม่ retry
  throw new Error("agentQuery: all tiers exhausted — IntelSphere (" + AGENT_QUERY_PROVIDER_PRIORITY.join(" → ") + ") + personal Gemini pool");
}

// One-off setup: สร้าง tab IntelSphere_Keys พร้อม headers A–Q (idempotent — เรียกซ้ำไม่ทำลายข้อมูล)
function setupIntelSphereSheet() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(INTELSPHERE_SHEET_NAME);
  var created = false;

  if (!sheet) {
    sheet = ss.insertSheet(INTELSPHERE_SHEET_NAME);
    created = true;
  }

  var headers = [
    "Timestamp", "Donor_Name", "API_Key", "Status", "Last_Used", "Last_Reset_Date",
    "Deepseek_Remaining", "Gemini_Remaining", "Meta_Remaining", "Nova_Remaining",
    "xAI_Remaining", "Qwen_Remaining", "OpenAI_Remaining", "Claude_Remaining",
    "Mistral_Remaining", "MiniMax_Remaining", "Notes"
  ];

  // เขียน headers เฉพาะเมื่อแถว 1 ยังว่าง (ไม่ทับของเดิม)
  var firstCell = sheet.getRange(1, 1).getValue();
  var headersWritten = false;
  if (!firstCell) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#e6f7ff");
    sheet.setFrozenRows(1);
    headersWritten = true;
  }

  // Startup column check (plan §2.5): ทุก {Provider}_Remaining header ต้อง resolve ได้
  var liveHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var missing = [];
  for (var i = 0; i < INTELSPHERE_PROVIDER_PRIORITY.length; i++) {
    if (liveHeaders.indexOf(INTELSPHERE_PROVIDER_PRIORITY[i] + "_Remaining") < 0) {
      missing.push(INTELSPHERE_PROVIDER_PRIORITY[i] + "_Remaining");
    }
  }

  return ContentService.createTextOutput(JSON.stringify({
    result: 'success',
    sheetCreated: created,
    headersWritten: headersWritten,
    missingProviderColumns: missing
  })).setMimeType(ContentService.MimeType.JSON);
}

// Seed/อัปเดต key หนึ่งใบใน IntelSphere_Keys — idempotent by API_Key, prefill remaining เต็มโควต้า, Status=Active
function seedIntelSphereKey(apiKey, donorName, notes) {
  if (!apiKey) {
    return ContentService.createTextOutput(JSON.stringify({
      result: 'error', message: 'apiKey required'
    })).setMimeType(ContentService.MimeType.JSON);
  }
  apiKey = String(apiKey).trim();

  // ตรวจ key กับ IntelSphere ก่อนบันทึก — กัน key ปลอม/พิมพ์ผิด/ถูกเพิกถอนเข้ามาปน pool
  var verdict = validateIntelSphereKeyLive(apiKey);
  if (verdict === 'invalid') {
    return ContentService.createTextOutput(JSON.stringify({
      result: 'error', message: 'API Key ไม่ถูกต้องหรือถูกเพิกถอนแล้ว กรุณาตรวจสอบ Key จาก gen.ai.kku.ac.th อีกครั้ง'
    })).setMimeType(ContentService.MimeType.JSON);
  }
  if (verdict !== 'valid') {
    return ContentService.createTextOutput(JSON.stringify({
      result: 'error', message: 'ไม่สามารถตรวจสอบ Key ได้ในขณะนี้ (ระบบ IntelSphere อาจขัดข้องชั่วคราว) กรุณาลองใหม่ภายหลัง'
    })).setMimeType(ContentService.MimeType.JSON);
  }

  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(INTELSPHERE_SHEET_NAME);
  if (!sheet) { setupIntelSphereSheet(); sheet = ss.getSheetByName(INTELSPHERE_SHEET_NAME); }

  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var colKey = headers.indexOf("API_Key");
  var tz = "Asia/Bangkok";
  var todayStr = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd");

  // หาแถวเดิมที่มี key นี้อยู่แล้ว (idempotent)
  var rowIndex = -1;
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][colKey]).trim() === String(apiKey).trim()) { rowIndex = i + 1; break; }
  }

  // สร้าง row object ตาม header order
  var rowValues = new Array(headers.length).fill("");
  function setCol(name, val) { var c = headers.indexOf(name); if (c >= 0) rowValues[c] = val; }
  setCol("Timestamp", new Date());
  setCol("Donor_Name", donorName || "");
  setCol("API_Key", apiKey);
  setCol("Status", "Active");
  setCol("Last_Reset_Date", todayStr);
  setCol("Notes", notes || "");
  for (var p = 0; p < INTELSPHERE_PROVIDER_PRIORITY.length; p++) {
    setCol(INTELSPHERE_PROVIDER_PRIORITY[p] + "_Remaining", INTELSPHERE_LIMITS[INTELSPHERE_PROVIDER_PRIORITY[p]]);
  }

  var appended = false;
  if (rowIndex > 0) {
    sheet.getRange(rowIndex, 1, 1, headers.length).setValues([rowValues]);
  } else {
    sheet.appendRow(rowValues);
    appended = true;
  }
  SpreadsheetApp.flush();
  CacheService.getScriptCache().remove("intelsphere_catalog"); // ให้ดึง live catalog ใหม่ด้วย key นี้
  CacheService.getScriptCache().remove("donor_credits"); // รายชื่อผู้บริจาคเปลี่ยน

  return ContentService.createTextOutput(JSON.stringify({
    result: 'success', appended: appended, updatedExisting: (rowIndex > 0),
    message: 'ขอบคุณสำหรับการบริจาค! Key ของคุณผ่านการตรวจสอบและพร้อมใช้งานแล้ว'
  })).setMimeType(ContentService.MimeType.JSON);
}

// Rate limit: 15 คำถาม/ชม. ต่อ session token (rolling window ผ่าน CacheService TTL)
function checkRateLimit(userToken) {
  var cache = CacheService.getScriptCache();
  var cacheKey = "rl_" + userToken;
  var count = parseInt(cache.get(cacheKey) || "0");
  if (count >= 15) return false;
  cache.put(cacheKey, String(count + 1), 3600);
  return true;
}

// Rate limit ทั่วไป: prefix แยกต่อ action, limit ต่อชั่วโมง (checkRateLimit เดิมคงไว้ — askAIExpert ใช้อยู่)
function checkActionRateLimit(prefix, token, limit) {
  var cache = CacheService.getScriptCache();
  var cacheKey = prefix + token;
  var count = parseInt(cache.get(cacheKey) || "0");
  if (count >= limit) return false;
  cache.put(cacheKey, String(count + 1), 3600);
  return true;
}

// รายชื่อผู้บริจาค key ที่ Active (ชื่ออย่างเดียว ไม่มี key) — cache 6 ชม., bust ตอน seed
function getDonorCredits() {
  var cache = CacheService.getScriptCache();
  var cached = cache.get("donor_credits");
  if (cached) return JSON.parse(cached);

  var donors = [];
  try {
    var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(INTELSPHERE_SHEET_NAME);
    if (sheet) {
      var data = sheet.getDataRange().getValues();
      var headers = data[0];
      var colName = headers.indexOf("Donor_Name");
      var colStatus = headers.indexOf("Status");
      for (var i = 1; i < data.length; i++) {
        var name = String(data[i][colName] || "").trim();
        if (data[i][colStatus] === "Active" && name && donors.indexOf(name) < 0) donors.push(name);
      }
    }
    cache.put("donor_credits", JSON.stringify(donors), 21600); // cache เฉพาะตอนอ่านสำเร็จ — error ชั่วคราวไม่ควรค้าง 6 ชม.
  } catch (e) {
    console.warn("[IntelSphere] getDonorCredits failed: " + e.message);
  }
  return donors;
}

// One-off setup: สร้าง tab AI_Feedback (idempotent — เรียกซ้ำไม่ทับข้อมูลเดิม)
