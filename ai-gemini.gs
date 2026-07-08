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


