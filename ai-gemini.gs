/* 
   =========================================
   ส่วนที่ 4: AI Expert & API Quota
   =========================================
*/

/**
 * ฟังก์ชันดึง API Key ที่พร้อมใช้งานและจัดการโควต้าต่อวัน
 */
function getAvailableAIKey(provider) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName("AI_Config");
  if (!sheet) return null;

  var data = sheet.getDataRange().getValues();
  var today = new Date().toDateString();

  for (var i = 1; i < data.length; i++) {
    // ลำดับคอลัมน์ A=0, B=1, C=2, D=3, E=4, F=5, G=6
    var key    = data[i][0]; // A: API_Key
    var prov   = data[i][1]; // B: Provider
    var model  = data[i][2]; // C: Model
    var limit  = parseInt(data[i][3]) || 0; // D: Daily_Limit
    
    // E: Usage_Count - บังคับให้เป็นตัวเลขเสมอ (กันความผิดพลาดถ้าในช่องเป็นวันที่)
    var usage = parseInt(data[i][4]);
    if (isNaN(usage)) usage = 0; 

    var lastUsedVal = data[i][5]; // F: Last_Used
    var lastDate = lastUsedVal ? new Date(lastUsedVal).toDateString() : "";
    var status = data[i][6]; // G: Status

    if (prov === provider) {
      // ตรรกะรีเซ็ตเมื่อขึ้นวันใหม่
      if (lastDate !== today) {
        usage = 0;
        sheet.getRange(i + 1, 5).setValue(0);         // Reset คอลัมน์ E (Usage_Count)
        sheet.getRange(i + 1, 6).setValue(new Date()); // Update คอลัมน์ F (Last_Used)
        sheet.getRange(i + 1, 7).setValue("Active");   // Update คอลัมน์ G (Status)
        status = "Active";
      }

      // ตรวจสอบว่า Status เป็น Active และ Usage ยังไม่เต็ม
      if (status === "Active" && usage < limit) {
        return { 
          key: key, 
          model: model || "gemini-1.5-flash", 
          index: i + 1, 
          usage: usage, 
          limit: limit 
        };
      } else if (usage >= limit && status !== "Exhausted") {
        sheet.getRange(i + 1, 7).setValue("Exhausted");
      }
    }
  }
  return null;
}

/**
 * อัปเดตจำนวนการใช้งานหลังจากเรียก AI สำเร็จ
 */
function updateAIUsage(index, currentUsage) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName("AI_Config");
  
  // บังคับให้ currentUsage เป็นตัวเลข
  var count = parseInt(currentUsage);
  if (isNaN(count)) count = 0;
  var newUsage = count + 1;

  // อัปเดตช่องให้ตรงคอลัมน์
  sheet.getRange(index, 5).setValue(newUsage);    // คอลัมน์ E: Usage_Count
  sheet.getRange(index, 6).setValue(new Date());  // คอลัมน์ F: Last_Used

  // ตรวจสอบ Limit จากคอลัมน์ D
  var limit = parseInt(sheet.getRange(index, 4).getValue()) || 0;
  if (newUsage >= limit) {
    sheet.getRange(index, 7).setValue("Exhausted"); // คอลัมน์ G: Status
  }
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
        updateAIUsage(apiKeyInfo.index, apiKeyInfo.usage);

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

// D13: fallback chain เมื่อโมเดลจากคอลัมน์ C ของ AI_Config ใช้ไม่ได้ (quota/deprecated)
var CONVERTER_FALLBACK_MODELS = ["gemini-3.5-flash", "gemini-2.5-flash", "gemini-3.1-pro"];
// กันชน 6-min execution limit: จำกัดจำนวนครั้งที่ยิง Gemini จริงต่อ 1 POST
var CONVERTER_MAX_ATTEMPTS = 3;

/**
 * Gemini call สำหรับแปลงข้อสอบ (คนละ tuning กับ callGeminiAI ของ chatbot):
 * JSON mode, temperature ต่ำ, maxOutputTokens สูง, ปิด thinking, ไม่มี systemInstruction
 * pdfB64 = base64 ของ PDF ทั้งไฟล์ (batch เดียว) หรือ images = dataURL ต่อหน้า (batch ใหญ่)
 * คืน { raw, finishReason, model } — ฝั่ง client เป็นคน parse (มี recovery logic ครบอยู่แล้ว)
 */
function callGeminiConverter(prompt, apiKeyInfo, pdfB64, images) {
  var models = [apiKeyInfo.model].concat(CONVERTER_FALLBACK_MODELS)
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
        updateAIUsage(apiKeyInfo.index, apiKeyInfo.usage);
        return { raw: res.raw, finishReason: res.finishReason, model: models[mi] };
      }
      lastErr = models[mi] + ": " + res.error;
      if (res.fatal) throw new Error("แปลงไม่สำเร็จ: " + lastErr);
      // RECITATION: ยืนยันจากการรันจริง 2 รอบว่าสลับ thinking variant ไม่ช่วย — bump temp แล้วข้ามไปโมเดลถัดไปเลย
      if (res.recitation) { convTemp = 0.8; break; }
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
    if (code === 429 || code === 404 || code >= 500) return { ok: false, error: msg, nextModel: true }; // quota/ไม่มีโมเดล/overloaded → โมเดลถัดไป
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

// บันทึก key ลง AI_Config (idempotent by API_Key คอลัมน์ A) — เรียกภายใต้ localized lock จาก router
// Model = "gemini-3.5-flash" (ตัวแรกของ CONVERTER_FALLBACK_MODELS) กันเปลือง 1 ใน 3 attempts กับรุ่นตาย
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
  var sheet = ss.getSheetByName("AI_Config");
  if (!sheet) {
    return ContentService.createTextOutput(JSON.stringify({
      result: 'error', message: 'ไม่พบชีต AI_Config ในระบบ'
    })).setMimeType(ContentService.MimeType.JSON);
  }

  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === apiKey) {
      sheet.getRange(i + 1, 7).setValue("Active"); // G: Status — ปลุก key เดิมที่อาจ Exhausted
      SpreadsheetApp.flush();
      return ContentService.createTextOutput(JSON.stringify({
        result: 'success', updatedExisting: true,
        message: 'Key นี้มีอยู่ในระบบแล้ว — เปิดใช้งานอีกครั้งเรียบร้อย ขอบคุณครับ'
      })).setMimeType(ContentService.MimeType.JSON);
    }
  }

  // A=API_Key B=Provider C=Model D=Daily_Limit E=Usage_Count F=Last_Used G=Status
  sheet.appendRow([apiKey, "Gemini", "gemini-3.5-flash", 200, 0, new Date(), "Active"]);
  SpreadsheetApp.flush();
  return ContentService.createTextOutput(JSON.stringify({
    result: 'success', appended: true,
    message: 'ขอบคุณสำหรับการบริจาค! Key ผ่านการตรวจสอบและพร้อมใช้งานแล้ว'
  })).setMimeType(ContentService.MimeType.JSON);
}


