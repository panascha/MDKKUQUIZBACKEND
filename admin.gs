function hashPasswordInternal(password) {
  if (!password) return "";
  var signature = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, password);
  return signature.map(function(byte) {
    var v = (byte < 0) ? (byte + 256) : byte;
    return ("0" + v.toString(16)).slice(-2);
  }).join("");
}

function verifyAdmin(username, password) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName("Admins");
  if (!sheet) return null;
  
  var hashedInput = hashPasswordInternal(password);
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] == username && (data[i][1] == password || data[i][1] == hashedInput)) {
      return {
        username: data[i][0],
        displayName: data[i][2],
        avatar: data[i][3] || "https://api.dicebear.com/7.x/avataaars/svg?seed=" + data[i][2],
        role: data[i][4],
        prefix: data[i][6],
        fullName: data[i][7],
        studentId: data[i][8],
        year: data[i][9],
        contact: data[i][10]
      };
    }
  }
  return null;
}

// Verify Google id_token (JWT) safely with Google tokeninfo endpoint
function verifyGoogleToken(idToken) {
  if (!idToken || String(idToken).trim() === "" || String(idToken) === "undefined" || String(idToken) === "null") {
    console.error("[AUTH] verifyGoogleToken failed: Token string is empty or invalid.");
    return null;
  }
  
  // 1. ตรวจสอบโครงสร้าง JWT และความถูกต้องของเวลาหมดอายุแบบถอดรหัสภายในเบื้องต้นก่อนยิงขอสิทธิ์
  try {
    var parts = idToken.split('.');
    if (parts.length === 3) {
      var decodedPayload = Utilities.newBlob(Utilities.base64DecodeWebSafe(parts[1])).getDataAsString();
      var localPayload = JSON.parse(decodedPayload);
      var now = Date.now() / 1000;
      
      // ป้องกันกรณี Token หมดอายุไปแล้วจากฝั่งเครื่องผู้ใช้งาน
      if (localPayload.exp && now >= localPayload.exp) {
        console.warn("[AUTH] verifyGoogleToken local check: Token exp claim expired (" + localPayload.exp + " < now: " + now + ")");
        return null;
      }
    }
  } catch (e) {
    console.warn("[AUTH] verifyGoogleToken local pre-check warning: " + e.message);
  }

  // 2. ส่งข้อมูลยืนยันความปลอดภัยทางอ้อมกับ Tokeninfo Endpoint ด้วยระบบ URL-safe
  var maxRetries = 2;
  var lastError = "";
  
  for (var i = 0; i < maxRetries; i++) {
    try {
      var url = "https://oauth2.googleapis.com/tokeninfo?id_token=" + encodeURIComponent(idToken);
      var response = UrlFetchApp.fetch(url, { "muteHttpExceptions": true });
      var responseCode = response.getResponseCode();
      var content = response.getContentText();
      
      if (responseCode == 200) {
        return JSON.parse(content);
      } else {
        lastError = "HTTP " + responseCode + " - " + content;
        console.warn("[AUTH] verifyGoogleToken remote attempt " + (i + 1) + " failed: " + lastError);
        Utilities.sleep(500); // ดีเลย์ก่อนสแกนซ้ำกรณีเกิดความหน่วงเครือข่ายชั่วคราว
      }
    } catch (e) {
      lastError = e.message;
      console.error("[AUTH] verifyGoogleToken remote attempt " + (i + 1) + " threw error: " + e.message);
      Utilities.sleep(500);
    }
  }
  
  console.error("[AUTH] verifyGoogleToken failed completely. Last error: " + lastError);
  return null;
}

// Lookup Admin from sheet matching verified Google email address
function findAdminByEmail(email) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName("Admins");
  if (!sheet) return null;

  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    // Column Index 5 contains KKUMail address
    if (String(data[i][5]).trim().toLowerCase() === String(email).trim().toLowerCase()) {
      return {
        username: data[i][0],
        displayName: data[i][2],
        avatar: data[i][3] || "https://api.dicebear.com/7.x/avataaars/svg?seed=" + data[i][2],
        role: data[i][4],
        prefix: data[i][6],
        fullName: data[i][7],
        studentId: data[i][8],
        year: data[i][9],
        contact: data[i][10],
        email: data[i][5]
      };
    }
  }
  return null;
}

function verifyUser(data) {
  if (data && data.sessionToken) {
    return verifySessionToken(data.sessionToken);
  } else if (data && data.googleIdToken) {
    var payload = verifyGoogleToken(data.googleIdToken);
    if (payload) return findAdminByEmail(payload.email);
  } else if (data) {
    return verifyAdmin(data.username, data.adminPass);
  }
  return null;
}

function uploadToDrive(base64Data, filename, mimeType) {
  try {
    var parentFolder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
    var targetFolder;
    
    var folders = parentFolder.getFoldersByName("User");
    if (folders.hasNext()) {
      targetFolder = folders.next();
    } else {
      targetFolder = parentFolder.createFolder("User");
    }
    
    var cleanBase64 = base64Data.split(',')[1] || base64Data;
    var fileBlob = Utilities.newBlob(Utilities.base64Decode(cleanBase64), mimeType, filename);
    
    var existingFiles = targetFolder.getFilesByName(filename);
    while (existingFiles.hasNext()) {
      existingFiles.next().setTrashed(true);
    }
    
    var file = targetFolder.createFile(fileBlob);
    //file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    return 'https://drive.google.com/uc?export=view&id=' + file.getId();
    
  } catch (e) {
    Logger.log("Drive Upload Error: " + e.message);
    throw new Error("Cannot upload file: " + e.message);
  }
}

function getPendingReportCount(filterSubject, startTime) {
  var v = getVotesVersionCached();
  var cleanFilter = filterSubject ? String(filterSubject).trim().toUpperCase() : "all";
  var cacheKey = "pending_report_count_" + v + "_" + cleanFilter;
  var cached = getLargeCache(cacheKey);
  if (cached != null) {
    return ContentService.createTextOutput(cached).setMimeType(ContentService.MimeType.JSON);
  }

  if (startTime) assertNotTimedOut_(startTime, 'getPendingReportCount');
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName("Report");
  if (!sheet) {
    var responseStr = JSON.stringify({ count: 0, samples: [] });
    return ContentService.createTextOutput(responseStr).setMimeType(ContentService.MimeType.JSON);
  }

  var data = sheet.getDataRange().getValues();
  var pendingCount = 0;
  var samples = [];

  for (var i = 1; i < data.length; i++) {
    var subjectRef = String(data[i][0]).trim().toUpperCase();
    var status = String(data[i][9]).trim();

    if ((status === "Pending" || status === "") && (cleanFilter === "all" || subjectRef === cleanFilter)) {
      pendingCount++;

      if (samples.length < 2) {
        samples.push({
          category: data[i][1],
          question: data[i][3].substring(0, 80) + "..."
        });
      }
    }
  }

  var responseObj = {
    count: pendingCount,
    samples: samples,
    subject: cleanFilter === "all" ? "ALL" : cleanFilter
  };
  var responseStr = JSON.stringify(responseObj);
  putLargeCache(cacheKey, responseStr, 300, startTime); // 5 minutes cache
  return ContentService.createTextOutput(responseStr).setMimeType(ContentService.MimeType.JSON);
}

