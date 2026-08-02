
/* 
   ========================================
   ส่วนที่ 1: การดึงข้อมูล (GET)
   =========================================
*/

function onSheetEdit(e) {
  if (!e) return;
  var sheet = e.source.getActiveSheet();
  var sheetName = sheet.getName();

  var watchSheetsQuestions = ['Questions', 'Structure', 'Category', 'Admins', 'Announcements'];
  var watchSheetsVotes = ['Report', 'Votes'];

  if (watchSheetsQuestions.indexOf(sheetName) > -1) {
    updateVersion();
  }
  if (watchSheetsVotes.indexOf(sheetName) > -1) {
    updateVotesVersion();
  }

  // ระบบตรวจสอบอัตโนมัติเมื่อคอลัมน์ Category (G) ของชีต Questions มีการแก้ไข
  if (sheetName === 'Questions') {
    var range = e.range;
    var startCol = range.getColumn();
    var endCol = range.getLastColumn();

    if (startCol <= 7 && endCol >= 7) {
      var startRow = Math.max(2, range.getRow()); // ข้ามหัวตาราง (Header)
      var endRow = range.getLastRow();
      var sortedNeeded = false;

      for (var r = startRow; r <= endRow; r++) {
        var qId = sheet.getRange(r, 1).getValue().toString().trim();
        var catRaw = sheet.getRange(r, 7).getValue().toString().trim();
        if (qId && catRaw !== "") {
          try {
            var categories = [];
            if (catRaw.indexOf("[") > -1) {
              categories = JSON.parse(catRaw.replace(/'/g, '"'));
            } else {
              categories = [catRaw];
            }
            if (categories.length >= 2) {
              autoCreateSplitCategories(qId, categories, true);
              sortedNeeded = true;
            }
          } catch (err) {
            console.error("Split error in onSheetEdit for row " + r + ": " + err.message);
          }
        }
      }

      if (sortedNeeded) {
        sortCategorySheet();
      }
    }
  }
}

function updateVersion() {
  var cache = PropertiesService.getScriptProperties();
  // ใช้ Timestamp ปัจจุบันเป็น Version ID (แม่นยำกว่าการนับเลข)
  var newVer = new Date().getTime().toString();
  cache.setProperty('v', newVer);
  try {
    CacheService.getScriptCache().put("v_cache", newVer, 60);
  } catch(e) {
    console.warn("Version cache put failed: " + e.message);
  }
  return newVer;
}

// --- SESSION TOKEN SYSTEM (30-day admin sessions) ---

var SESSION_EXPIRY_DAYS = 30;

function generateSessionToken() {
  var chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789';
  var token = '';
  for (var i = 0; i < 64; i++) {
    token += chars.charAt(Math.floor(Math.random() * chars.length));
  }
  return token;
}

function createSession(email, userObj) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName("Sessions") || ss.insertSheet("Sessions");
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(["Token", "Email", "CreatedAt", "ExpiresAt", "LastUsed"]);
    sheet.getRange(1, 1, 1, 5).setFontWeight("bold");
  }
  // Multi-device: เก็บได้สูงสุด 5 session ต่ออีเมล (ลบอันเก่าสุดเกินโควต้า ไม่ลบทั้งหมดแบบเดิม)
  capSessionsByEmail(sheet, email, 4);
  var token = generateSessionToken();
  var now = new Date();
  var expiry = new Date(now.getTime() + SESSION_EXPIRY_DAYS * 24 * 60 * 60 * 1000);
  sheet.appendRow([token, email, now.toISOString(), expiry.toISOString(), now.toISOString()]);
  return token;
}

function verifySessionToken(token) {
  if (!token) return null;
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName("Sessions");
  if (!sheet) return null;
  var data = sheet.getDataRange().getValues();
  var now = new Date();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === token) {
      var expiry = new Date(data[i][3]);
      if (now > expiry) {
        sheet.deleteRow(i + 1);
        return null;
      }
      // T2.6: เขียน LastUsed เฉพาะเมื่อค่าเดิมเก่ากว่า 1 ชั่วโมง (ลด Sheets write ทุก request ของแอดมิน)
      var lastUsed = data[i][4] ? new Date(data[i][4]) : null;
      if (!lastUsed || isNaN(lastUsed.getTime()) || (now.getTime() - lastUsed.getTime()) > 3600000) {
        sheet.getRange(i + 1, 5).setValue(now.toISOString());
      }
      return findAdminByEmail(data[i][1]);
    }
  }
  return null;
}

// ตรวจ token ของ "ใครก็ได้" (Admin หรือ Student) — ใช้เฉพาะ endpoint sync ความคืบหน้า + verifySession
// ห้ามใช้แทน verifySessionToken ใน action ฝั่งแอดมิน: token ของ Student ต้องผ่านไม่ได้
function verifyAnySession(token) {
  if (!token) return null;
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName("Sessions");
  if (!sheet) return null;
  var data = sheet.getDataRange().getValues();
  var now = new Date();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === token) {
      var expiry = new Date(data[i][3]);
      if (now > expiry) {
        sheet.deleteRow(i + 1);
        return null;
      }
      // T2.6: เขียน LastUsed เฉพาะเมื่อค่าเดิมเก่ากว่า 1 ชั่วโมง
      var lastUsed = data[i][4] ? new Date(data[i][4]) : null;
      if (!lastUsed || isNaN(lastUsed.getTime()) || (now.getTime() - lastUsed.getTime()) > 3600000) {
        sheet.getRange(i + 1, 5).setValue(now.toISOString());
      }
      var email = data[i][1];
      var adminUser = findAdminByEmail(email);
      if (adminUser) return adminUser;
      return { email: email, role: "Student", displayName: String(email).split("@")[0] };
    }
  }
  return null;
}

// เก็บ session ล่าสุดไว้ไม่เกิน keep รายการต่ออีเมล (ลบอันเก่าสุดออก) — CreatedAt คือคอลัมน์ 3
function capSessionsByEmail(sheet, email, keep) {
  var data = sheet.getDataRange().getValues();
  var rows = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i][1] === email) rows.push({ row: i + 1, createdAt: new Date(data[i][2]).getTime() || 0 });
  }
  if (rows.length <= keep) return;
  rows.sort(function (a, b) { return a.createdAt - b.createdAt; }); // เก่าสุดก่อน
  var toDelete = rows.slice(0, rows.length - keep).map(function (r) { return r.row; });
  toDelete.sort(function (a, b) { return b - a; }); // ลบจากล่างขึ้นบน กัน index เลื่อน
  toDelete.forEach(function (rowIdx) { sheet.deleteRow(rowIdx); });
}

// --- PROGRESS SYNC (cross-device continue) ---

function getOrCreateProgressSheet(ss) {
  var sheet = ss.getSheetByName("Progress");
  if (!sheet) {
    sheet = ss.insertSheet("Progress");
    sheet.appendRow(["Email", "Subject", "Timestamp", "Blob"]);
    sheet.getRange(1, 1, 1, 4).setFontWeight("bold").setBackground("#d0e0f0");
  }
  return sheet;
}

function cleanupSessionsByEmail(sheet, email) {
  var data = sheet.getDataRange().getValues();
  for (var i = data.length - 1; i >= 1; i--) {
    if (data[i][1] === email) sheet.deleteRow(i + 1);
  }
}

function cleanupExpiredSessions() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName("Sessions");
  if (!sheet) return;
  var data = sheet.getDataRange().getValues();
  var now = new Date();
  for (var i = data.length - 1; i >= 1; i--) {
    if (new Date(data[i][3]) < now) sheet.deleteRow(i + 1);
  }
}

// ติดตั้ง time-driven trigger รันทุกวัน ~ตี 2 (idempotent) — เว้นตี 3-6 ให้ batch jobs เดิม
// รันเองครั้งเดียวจาก Apps Script editor (ไม่เรียกอัตโนมัติตอนโหลด — เหมือน install*Trigger ตัวอื่น)
function installSessionCleanupTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'cleanupExpiredSessions') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  ScriptApp.newTrigger('cleanupExpiredSessions').timeBased().everyDays(1).atHour(2).create();
  return 'installed';
}

function getOrCreateAnnouncementsSheet(ss) {
  var sheet = ss.getSheetByName("Announcements");
  if (!sheet) {
    sheet = ss.insertSheet("Announcements");
    sheet.appendRow(["Id", "Text", "Type", "Active", "Order"]);
    sheet.getRange(1, 1, 1, 5).setFontWeight("bold").setBackground("#d9ead3");
    // Add default values
    sheet.appendRow([
      "ANN_1", 
      "<strong><i class=\"fas fa-bullhorn\"></i> ยินดีต้อนรับสู่ MDKKUQUIZ!</strong> ระบบคลังข้อสอบและวิเคราะห์จุดอ่อนสำหรับเตรียมตัวสอบ", 
      "info", 
      "TRUE", 
      1
    ]);
    sheet.appendRow([
      "ANN_2", 
      "<strong style=\"color: #ea580c;\"><i class=\"fas fa-star\"></i> อัปเดตใหม่!</strong> ระบบแก้ไขข้อสอบ & Rich Explanation พร้อมระบบ AI Assistant เรียบร้อยแล้ว", 
      "warning", 
      "TRUE", 
      2
    ]);
  }
  return sheet;
}
