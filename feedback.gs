/*
   =========================================================================
   APP FEEDBACK (feature/bug reporting จากผู้ใช้ REAL) — คนละระบบกับ Report ข้อสอบ
   submitFeedback (localized tier, ใน router-doPost) / getFeedback (lock-free, admin)
   แผน: Idea/active/user-feedback-reporting.md
   ========================================================================= */

// idempotent: สร้างชีต Feedback ถ้ายังไม่มี (mirror setupAiFeedbackSheet)
function setupFeedbackSheet() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName("Feedback");
  if (!sheet) sheet = ss.insertSheet("Feedback");

  var headers = ["Timestamp", "Type", "Description", "Email", "ClientId", "Context", "Images", "Status", "AdminNote"];
  if (!sheet.getRange(1, 1).getValue()) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#e6f7ff");
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// เซฟรูป feedback หนึ่งรูปลง subfolder "Feedback" ใต้ DRIVE_FOLDER_ID (แยกโฟลเดอร์ = ล้างทิ้งง่าย)
// รับ data URI base64 → คืน URL แบบ uc?export=view เหมือน uploadQuestionImageToDrive
function saveFeedbackImageToDrive(base64Data) {
  var rootFolder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
  var fbFolder = getOrCreateFolder(rootFolder, "Feedback");

  var mimeType = "image/jpeg";
  var fileExtension = "jpg";
  if (base64Data.indexOf("data:") === 0) {
    var partsMime = base64Data.split(";")[0].split(":");
    if (partsMime.length > 1 && partsMime[1].indexOf("image/") === 0) {
      mimeType = partsMime[1];
      fileExtension = mimeType.split("/")[1] || "jpg";
    }
  }

  var cleanBase64 = base64Data.split(',')[1] || base64Data;
  var fileName = "FB_" + new Date().getTime() + "_" + Math.floor(Math.random() * 10000) + "." + fileExtension;
  var blob = Utilities.newBlob(Utilities.base64Decode(cleanBase64), mimeType, fileName);
  var file = fbFolder.createFile(blob);
  return 'https://drive.google.com/uc?export=view&id=' + file.getId();
}

// อ่าน Feedback ทุกแถวเป็น JSON (เรียกหลังผ่าน admin auth แล้วเท่านั้น — แถวมี PII: email/free text)
function getFeedbackRows() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName("Feedback");
  var rows = [];
  if (sheet && sheet.getLastRow() > 1) {
    var values = sheet.getDataRange().getValues();
    for (var i = 1; i < values.length; i++) {
      rows.push({
        timestamp: values[i][0] instanceof Date ? values[i][0].toISOString() : String(values[i][0]),
        type: values[i][1],
        description: values[i][2],
        email: values[i][3],
        clientId: values[i][4],
        context: values[i][5],
        images: values[i][6],
        status: values[i][7],
        adminNote: values[i][8]
      });
    }
  }
  return ContentService.createTextOutput(JSON.stringify({ result: 'success', feedback: rows }))
    .setMimeType(ContentService.MimeType.JSON);
}
