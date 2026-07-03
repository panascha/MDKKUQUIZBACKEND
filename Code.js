var SHEET_ID = '12rN8vcykEwgcPFK4LoOj18PEhj7JPhwMfz6uUkKrhJU';
var DRIVE_FOLDER_ID = '1nzLH2ia2lL2TMxfrr6Kv-5fhsWwOWSCm'; 

var VOTE_THRESHOLD_CONFIRM = 2;

var REPORT_VOTE_THRESHOLD = 5;

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('🛠️ MDKKU Tools')
      .addItem('🔄 สั่งเรียงลำดับ Category ใหม่', 'sortCategorySheet')
      .addItem('✅ Verify ทั้งหมด (เฉพาะที่มี 1 Cat & ไม่ใช่ AI)', 'verifyAllSingleCategoryVotes')
      .addItem('📂 แยกกลุ่มวิชาอัตโนมัติ (Extracted) ทั้งหมด', 'runManualSplitExtraction')
      .addItem('📃 ตรวจสอบ Image url', 'generateImageVerificationReport')
      .addToUi();
}

// --- Caching Engine for Large Value Chunking (Apps Script 100KB Limit Workaround) ---

function getVersionCached() {
  var cache = CacheService.getScriptCache();
  var v = cache.get("v_cache");
  if (v == null) {
    v = PropertiesService.getScriptProperties().getProperty('v') || "0";
    try {
      cache.put("v_cache", v, 60); // Cache version check for 60 seconds
    } catch(e) {
      console.warn("Version cache write error: " + e.message);
    }
  }
  return v;
}

function getVotesVersionCached() {
  var cache = CacheService.getScriptCache();
  var v = cache.get("v_votes_cache");
  if (v == null) {
    v = PropertiesService.getScriptProperties().getProperty('v_votes') || "0";
    try {
      cache.put("v_votes_cache", v, 60); // Cache votes version for 60 seconds
    } catch (e) {
      console.warn("Votes version cache write error: " + e.message);
    }
  }
  return v;
}

function updateVotesVersion() {
  var cache = PropertiesService.getScriptProperties();
  var newVer = new Date().getTime().toString();
  cache.setProperty('v_votes', newVer);
  try {
    CacheService.getScriptCache().put("v_votes_cache", newVer, 60);
  } catch (e) {
    console.warn("Votes version cache put failed: " + e.message);
  }
  return newVer;
}

function putLargeCache(key, value, ttl) {
  if (!value) return;
  var cache = CacheService.getScriptCache();
  // Downsized to 25KB character slices to protect against multi-byte (Thai) UTF-8 expansion (up to 3x bytes per char)
  var chunkSize = 25 * 1024;
  var chunks = Math.ceil(value.length / chunkSize);

  try {
    cache.put(key + "_chunks", String(chunks), ttl);
    for (var i = 0; i < chunks; i++) {
      cache.put(key + "_chunk_" + i, value.substring(i * chunkSize, (i + 1) * chunkSize), ttl);
    }
  } catch (e) {
    console.warn("putLargeCache failed for key " + key + ": " + e.message);
  }
}

function getLargeCache(key) {
  var cache = CacheService.getScriptCache();
  var chunksStr = cache.get(key + "_chunks");
  if (!chunksStr) return null;
  
  var chunks = parseInt(chunksStr, 10);
  var value = "";
  for (var i = 0; i < chunks; i++) {
    var chunk = cache.get(key + "_chunk_" + i);
    if (chunk == null) return null; // If any chunk is lost, treat as cache miss
    value += chunk;
  }
  return value;
}

function getStructureDataCached(filterSubject) {
  var v = getVersionCached();
  var cleanFilter = filterSubject ? String(filterSubject).trim().toUpperCase() : "all";
  var cacheKey = "struct_" + v + "_" + cleanFilter;
  
  var cachedStr = getLargeCache(cacheKey);
  if (cachedStr != null) {
    return ContentService.createTextOutput(cachedStr).setMimeType(ContentService.MimeType.JSON);
  }
  
  var response = getStructureData(filterSubject);
  var responseStr = response.getContent();
  putLargeCache(cacheKey, responseStr, 1800); // 30 minutes
  return response;
}

function getQuestionsDataCached(filterSubject, ss) {
  var v = getVersionCached();
  var cleanFilter = filterSubject ? String(filterSubject).trim().toUpperCase() : "all";
  var cacheKey = "questions_" + v + "_" + cleanFilter;

  var cachedStr = getLargeCache(cacheKey);
  if (cachedStr != null) {
    return ContentService.createTextOutput(cachedStr).setMimeType(ContentService.MimeType.JSON);
  }

  var response = getQuestionsData(filterSubject, ss);
  var responseStr = response.getContent();
  putLargeCache(cacheKey, responseStr, 1800); // ขยาย Cache เป็น 30 นาที เนื่องจากคีย์ผูกกับเวอร์ชันอยู่แล้ว (เมื่อมีข้อมูลใหม่แคชจะรีเซ็ตอัตโนมัติ)
  return response;
}

function getAllDataForAdminCached() {
  var v = getVersionCached();
  var cacheKey = "admin_all_data_" + v;

  var cachedStr = getLargeCache(cacheKey);
  if (cachedStr != null) {
    return ContentService.createTextOutput(cachedStr).setMimeType(ContentService.MimeType.JSON);
  }

  var response = getAllDataForAdmin();
  var responseStr = response.getContent();
  putLargeCache(cacheKey, responseStr, 300); // Cache complete dataset for 5 minutes
  return response;
}

// ────────────────────────────────────────────────────────────────────
// NEW: ADVANCED CACHED SHEET LOADERS (Bypasses Sheets API contention)
// ────────────────────────────────────────────────────────────────────

function getCategorySheetDataCached(ss) {
  var v = getVersionCached();
  var cacheKey = "category_sheet_raw_" + v;
  var cached = getLargeCache(cacheKey);
  if (cached) {
    return JSON.parse(cached);
  }
  if (!ss) ss = SpreadsheetApp.openById(SHEET_ID);
  var catSheet = ss.getSheetByName('Category');
  if (!catSheet) return [];
  var rows = catSheet.getDataRange().getValues();
  putLargeCache(cacheKey, JSON.stringify(rows), 1800); // 30 minutes
  return rows;
}

function getStructureSheetDataCached(ss) {
  var v = getVersionCached();
  var cacheKey = "structure_sheet_raw_" + v;
  var cached = getLargeCache(cacheKey);
  if (cached) {
    return JSON.parse(cached);
  }
  if (!ss) ss = SpreadsheetApp.openById(SHEET_ID);
  var structSheet = ss.getSheetByName('Structure');
  if (!structSheet) return [];
  var rows = structSheet.getDataRange().getValues();
  putLargeCache(cacheKey, JSON.stringify(rows), 1800); // 30 minutes
  return rows;
}

function getAllQuestionsCached(ss) {
  var v = getVersionCached();
  var cacheKey = "all_questions_raw_" + v;
  var cached = getLargeCache(cacheKey);
  if (cached) {
    return JSON.parse(cached);
  }
  if (!ss) ss = SpreadsheetApp.openById(SHEET_ID);
  var qSheet = ss.getSheetByName('Questions');
  var qLastRow = qSheet.getLastRow();
  if (qLastRow <= 1) return [];

  var qData = qSheet.getRange(2, 1, qLastRow - 1, 7).getValues();
  putLargeCache(cacheKey, JSON.stringify(qData), 1800); // 30 minutes
  return qData;
}

function getCategoryToSubjectMapCached(ss) {
  var v = getVersionCached();
  var cacheKey = "cat_to_subj_map_" + v;
  var cached = getLargeCache(cacheKey);
  if (cached) {
    return JSON.parse(cached);
  }
  var catRows = getCategorySheetDataCached(ss);
  var categoryToSubjectMap = {};
  for (var i = 1; i < catRows.length; i++) {
    categoryToSubjectMap[String(catRows[i][0]).trim()] = String(catRows[i][1]).trim().toUpperCase();
  }
  putLargeCache(cacheKey, JSON.stringify(categoryToSubjectMap), 1800); // 30 minutes
  return categoryToSubjectMap;
}

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
  cleanupSessionsByEmail(sheet, email);
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

function doGet(e) {
  var action = e.parameter.action;
  var clientVer = e.parameter.clientVer;
  var serverVer = getVersionCached();

  // คืนค่าสถานะ NOT_MODIFIED ทันทีเพื่อประหยัด Round-trip หากเวอร์ชันของไคลเอนต์ล่าสุดตรงกับเซิร์ฟเวอร์ (ช่วยประหยัดโหลดและลดการสปินอัพคอนเทนเนอร์)
  if (clientVer && clientVer === serverVer) {
    return ContentService.createTextOutput(JSON.stringify({ status: 'NOT_MODIFIED', v: serverVer })).setMimeType(ContentService.MimeType.JSON);
  }

  if (action == 'checkVersion') {
    return ContentService.createTextOutput(JSON.stringify({ v: serverVer })).setMimeType(ContentService.MimeType.JSON);
  }
  if (action == 'getStructure') return getStructureDataCached(e.parameter.subject);
  if (action == 'getQuestions') return getQuestionsDataCached(e.parameter.subject);
  if (action == 'getPendingVotes') return getPendingVotesData(e.parameter.qid);
  if (action == 'getPendingReports') return getPendingReportsData(e.parameter.qid);
  if (action == 'getPendingVotesReports') return getPendingVotesReportsData(e.parameter.subject);
  if (action == 'getAllData') return getAllDataForAdminCached();
  if (action == 'getLogsPage') return getLogsPageData(e.parameter.offset, e.parameter.limit);
  if (action == 'getPendingReportCount') return getPendingReportCount(e.parameter.subject);
  if (action == 'getChangedSince') return getChangedSinceTimestamp(e.parameter.since, e.parameter.subject);
  if (action == 'setupIntelSphere') return setupIntelSphereSheet(); // idempotent one-off: สร้าง tab IntelSphere_Keys ถ้ายังไม่มี


  return ContentService.createTextOutput("Action not defined").setMimeType(ContentService.MimeType.TEXT);
}

function getAdminsList() {
    return getSheetDataJSON('Admins');
}

function getAllDataForAdmin() {
  var ss = SpreadsheetApp.openById(SHEET_ID);

  // ดึงข้อมูล Admins แบบเร็ว
  var adminsRaw = getSheetDataJSON('Admins', ss);
  var adminsSafe = adminsRaw.map(function (admin) {
    var safeAdmin = {};
    for (var key in admin) {
      if (key !== 'Password') safeAdmin[key] = admin[key];
    }
    return safeAdmin;
  });

  getOrCreateAnnouncementsSheet(ss); // Ensure sheet exists

  var data = {
    v: getVersionCached(), // แทรกเวอร์ชันปัจจุบันเพื่อให้ฝั่งไคลเอนต์ใช้ซิงค์ในรอบเดี่ยวได้โดยไม่ต้องยิง checkVersion แยก
    questions: JSON.parse(getQuestionsData('', ss).getContent()), // ส่ง ss เข้าไปด้วย
    structure: getSheetDataJSON('Structure', ss),
    category: getSheetDataJSON('Category', ss),
    report: getSheetDataJSON('Report', ss),
    votes: getSheetDataJSON('Votes', ss),
    logs: getLogsTailJSON(ss, 300), // จำกัดเฉพาะ 300 แถวล่าสุด (Logs โตไม่จำกัด) — โหลดเต็มผ่าน action=getLogsPage
    admins: adminsSafe,
    announcements: getSheetDataJSON('Announcements', ss)
  };
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
}

function getSheetDataJSON(sheetName, ss) {
    if (!ss) ss = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) return [];

    var data = sheet.getDataRange().getValues();
    if (data.length < 2) return [];

    var headers = data[0];
    var result = [];

    for (var i = 1; i < data.length; i++) {
        var obj = {};
        for (var j = 0; j < headers.length; j++) {
            var value = data[i][j];
            obj[headers[j]] = (value instanceof Date) ? value.toISOString() : value;
        }
        result.push(obj);
    }
    return result;
}

// อ่านเฉพาะ Log ท้ายสุด (tail) จำนวน limit แถว โดยไม่อ่านทั้งชีต (ป้องกัน Logs ที่โตไม่จำกัด)
// คืนค่า object array แบบเดียวกับ getSheetDataJSON('Logs') เรียงจากเก่า→ใหม่ ตามลำดับในชีต
function getLogsTailJSON(ss, limit) {
    if (!ss) ss = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName('Logs');
    if (!sheet) return [];

    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];

    var numCols = sheet.getLastColumn();
    var headers = sheet.getRange(1, 1, 1, numCols).getValues()[0];

    var maxRows = limit || 300;
    var numRows = Math.min(maxRows, lastRow - 1);
    var startRow = lastRow - numRows + 1;

    var data = sheet.getRange(startRow, 1, numRows, numCols).getValues();
    var result = [];
    for (var i = 0; i < data.length; i++) {
        var obj = {};
        for (var j = 0; j < headers.length; j++) {
            var value = data[i][j];
            obj[headers[j]] = (value instanceof Date) ? value.toISOString() : value;
        }
        result.push(obj);
    }
    return result;
}

// Server-side pagination ของชีต Logs สำหรับหน้า "ประวัติทั้งหมด" ในแดชบอร์ดแอดมิน
// offset นับจากแถวใหม่สุด (offset=0 = ชุดล่าสุด), limit = จำนวนแถวต่อหน้า (ค่าเริ่มต้น 300)
// คืนค่า logs เป็น object array รูปแบบเดียวกับ entry ใน getAllDataForAdmin().logs
function getLogsPageData(offsetStr, limitStr) {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName('Logs');
    if (!sheet) {
        return ContentService.createTextOutput(JSON.stringify({ status: 'success', logs: [], total: 0, offset: 0, limit: 0 })).setMimeType(ContentService.MimeType.JSON);
    }

    var lastRow = sheet.getLastRow();
    var total = lastRow > 1 ? lastRow - 1 : 0;

    var offset = parseInt(offsetStr) || 0;
    if (offset < 0) offset = 0;
    var limit = parseInt(limitStr) || 300;
    if (limit < 1) limit = 300;

    if (total === 0 || offset >= total) {
        return ContentService.createTextOutput(JSON.stringify({ status: 'success', logs: [], total: total, offset: offset, limit: limit })).setMimeType(ContentService.MimeType.JSON);
    }

    var numCols = sheet.getLastColumn();
    var headers = sheet.getRange(1, 1, 1, numCols).getValues()[0];

    // แถวใหม่สุดอยู่ล่างสุด (lastRow). ข้าม offset แถวใหม่สุด แล้วดึงถัดไป limit แถว
    var numRows = Math.min(limit, total - offset);
    var startRow = lastRow - offset - numRows + 1;

    var data = sheet.getRange(startRow, 1, numRows, numCols).getValues();
    var result = [];
    for (var i = 0; i < data.length; i++) {
        var obj = {};
        for (var j = 0; j < headers.length; j++) {
            var value = data[i][j];
            obj[headers[j]] = (value instanceof Date) ? value.toISOString() : value;
        }
        result.push(obj);
    }

    return ContentService.createTextOutput(JSON.stringify({
        status: 'success',
        logs: result,
        total: total,
        offset: offset,
        limit: limit
    })).setMimeType(ContentService.MimeType.JSON);
}

function getPendingVotesData(qid) {
  var v = getVotesVersionCached();
  var cacheKey = "pending_votes_" + v + "_" + qid;
  var cached = getLargeCache(cacheKey);
  if (cached != null) {
    return ContentService.createTextOutput(cached).setMimeType(ContentService.MimeType.JSON);
  }

  var ss = SpreadsheetApp.openById(SHEET_ID);
  var voteSheet = ss.getSheetByName("Votes");
  var result = [];
  if (voteSheet) {
    var data = voteSheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      var status = data[i][5];
      if (data[i][0] == qid && (status == "Pending" || status == "Approved")) {
        result.push({
          categoryId: data[i][2],
          count: data[i][3],
          status: status
        });
      }
    }
  }

  var responseObj = {
    votes: result,
    thresholds: {
      confirm: VOTE_THRESHOLD_CONFIRM
    }
  };
  var responseStr = JSON.stringify(responseObj);
  putLargeCache(cacheKey, responseStr, 300); // Cache for 5 minutes
  return ContentService.createTextOutput(responseStr).setMimeType(ContentService.MimeType.JSON);
}

function getPendingReportsData(qid) {
  var v = getVotesVersionCached();
  var cacheKey = "pending_reports_" + v + "_" + qid;
  var cached = getLargeCache(cacheKey);
  if (cached != null) {
    return ContentService.createTextOutput(cached).setMimeType(ContentService.MimeType.JSON);
  }

  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName("Report");
  var result = [];
  if (sheet) {
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][2]).trim() !== String(qid) || String(data[i][9]).trim() !== "Pending") continue;
      result.push({
        timestamp: data[i][8],
        suggestedChoice: data[i][6],
        suggestedExplain: data[i][12] || "",
        reportDetail: data[i][7],
        voteCount: parseInt(data[i][13]) || 0
      });
    }
  }
  result.sort(function (a, b) { return b.voteCount - a.voteCount; });

  var responseObj = { reports: result, threshold: REPORT_VOTE_THRESHOLD };
  var responseStr = JSON.stringify(responseObj);
  putLargeCache(cacheKey, responseStr, 300); // Cache for 5 minutes
  return ContentService.createTextOutput(responseStr).setMimeType(ContentService.MimeType.JSON);
}

// T1.1: Bulk endpoint — คืน pending votes + reports ของทุกข้อในวิชาเดียว (sparse map ตาม qid)
// แคชทั้งก้อนต่อวิชาโดยผูกกับ votes-version key (โหวต/รายงาน 1 ครั้ง = ล้างแคชครั้งเดียว)
// value ต่อ qid มีรูปแบบเดียวกับ endpoint per-qid เดิม (votes -> {votes,thresholds}, reports -> {reports,threshold})
function getPendingVotesReportsData(subjectParam) {
  var v = getVotesVersionCached();
  var cleanFilter = subjectParam ? String(subjectParam).trim().toUpperCase() : "all";
  var cacheKey = "pending_vr_" + v + "_" + cleanFilter;
  var cached = getLargeCache(cacheKey);
  if (cached != null) {
    return ContentService.createTextOutput(cached).setMimeType(ContentService.MimeType.JSON);
  }

  var ss = SpreadsheetApp.openById(SHEET_ID);

  // สร้างเซ็ตของ qid ที่อยู่ในวิชานี้ (ผ่าน category->subject map + คำถามที่แคชไว้)
  // ใช้ตัวคำถามเป็นเกณฑ์เพื่อความถูกต้อง โดยไม่ขึ้นกับความกำกวมของคอลัมน์ subject ในชีต Votes/Report
  var subjectQids = null; // null = รับทุก qid (กรณี subject = all)
  if (cleanFilter !== "all") {
    var catToSubj = getCategoryToSubjectMapCached(ss);
    var qData = getAllQuestionsCached(ss);
    subjectQids = {};
    for (var qi = 0; qi < qData.length; qi++) {
      var cats = [];
      try {
        var craw = String(qData[qi][6]).trim();
        if (craw !== "") cats = (craw.indexOf("[") > -1) ? JSON.parse(craw.replace(/'/g, '"')) : [craw];
      } catch (e) { cats = []; }
      for (var ci = 0; ci < cats.length; ci++) {
        if ((catToSubj[cats[ci]] || "") === cleanFilter) {
          subjectQids[String(qData[qi][0]).trim()] = true;
          break;
        }
      }
    }
  }

  // --- Votes (สถานะ Pending หรือ Approved) ---
  var votesMap = {};
  var voteSheet = ss.getSheetByName("Votes");
  if (voteSheet) {
    var vv = voteSheet.getDataRange().getValues();
    for (var i = 1; i < vv.length; i++) {
      var status = vv[i][5];
      if (!(status == "Pending" || status == "Approved")) continue;
      var qid = String(vv[i][0]).trim();
      if (subjectQids && !subjectQids[qid]) continue;
      if (!votesMap[qid]) votesMap[qid] = { votes: [], thresholds: { confirm: VOTE_THRESHOLD_CONFIRM } };
      votesMap[qid].votes.push({
        categoryId: vv[i][2],
        count: vv[i][3],
        status: status
      });
    }
  }

  // --- Reports (สถานะ Pending) ---
  var reportsMap = {};
  var reportSheet = ss.getSheetByName("Report");
  if (reportSheet) {
    var rv = reportSheet.getDataRange().getValues();
    for (var j = 1; j < rv.length; j++) {
      if (String(rv[j][9]).trim() !== "Pending") continue;
      var rqid = String(rv[j][2]).trim();
      if (subjectQids && !subjectQids[rqid]) continue;
      if (!reportsMap[rqid]) reportsMap[rqid] = { reports: [], threshold: REPORT_VOTE_THRESHOLD };
      reportsMap[rqid].reports.push({
        timestamp: rv[j][8],
        suggestedChoice: rv[j][6],
        suggestedExplain: rv[j][12] || "",
        reportDetail: rv[j][7],
        voteCount: parseInt(rv[j][13]) || 0
      });
    }
  }
  // เรียง reports ต่อ qid ตาม voteCount มาก→น้อย (ให้เหมือน endpoint per-qid เดิม)
  Object.keys(reportsMap).forEach(function (k) {
    reportsMap[k].reports.sort(function (a, b) { return b.voteCount - a.voteCount; });
  });

  var responseObj = {
    status: 'success',
    data: {
      votes: votesMap,
      reports: reportsMap
    }
  };
  var responseStr = JSON.stringify(responseObj);
  putLargeCache(cacheKey, responseStr, 300); // Cache for 5 minutes (ผูกกับ votes-version key)
  return ContentService.createTextOutput(responseStr).setMimeType(ContentService.MimeType.JSON);
}

function getAnnouncementsDataCached(ss) {
  var v = getVersionCached();
  var cacheKey = "announcements_data_raw_" + v;
  var cached = getLargeCache(cacheKey);
  if (cached) {
    return JSON.parse(cached);
  }
  if (!ss) ss = SpreadsheetApp.openById(SHEET_ID);
  getOrCreateAnnouncementsSheet(ss); // Ensure sheet exists
  var data = getSheetDataJSON('Announcements', ss);
  putLargeCache(cacheKey, JSON.stringify(data), 1800); // 30 minutes
  return data;
}

function getStructureData(filterSubject) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var cleanFilter = filterSubject ? String(filterSubject).trim().toUpperCase() : "";

  var rows = getStructureSheetDataCached(ss);
  var structData = [];
  for (var i = 1; i < rows.length; i++) {
    if (cleanFilter !== "" && String(rows[i][1]).trim().toUpperCase() !== cleanFilter) continue;
    structData.push({
      year: rows[i][0],
      subjectId: rows[i][1],
      subjectName: rows[i][2],
      accordionGroup: rows[i][3]
    });
  }

  var catRows = getCategorySheetDataCached(ss);
  var categoryData = [];
  for (var i = 1; i < catRows.length; i++) {
    if (cleanFilter !== "" && String(catRows[i][1]).trim().toUpperCase() !== cleanFilter) continue;
    categoryData.push({
      categoryId: catRows[i][0],
      subjectRef: catRows[i][1],
      accordionGroup: catRows[i][2],
      categoryName: catRows[i][3]
    });
  }

  var announcementsData = getAnnouncementsDataCached(ss); // Optimized to use cache!

  return ContentService.createTextOutput(JSON.stringify({
    subjects: structData,
    category: categoryData,
    announcements: announcementsData
  })).setMimeType(ContentService.MimeType.JSON);
}

function getQuestionsData(filterSubject, ss) {
  if (!ss) ss = SpreadsheetApp.openById(SHEET_ID);
  var cleanFilter = filterSubject ? String(filterSubject).trim().toUpperCase() : "";

  var categoryToSubjectMap = getCategoryToSubjectMapCached(ss);

  var qData = getAllQuestionsCached(ss);
  if (qData.length === 0) return ContentService.createTextOutput("[]").setMimeType(ContentService.MimeType.JSON);

  var questions = qData.map(function (row) {
    var categories = [];
    try {
      var catRaw = row[6].toString().trim();
      if (catRaw !== "") {
        categories = (catRaw.indexOf("[") > -1) ? JSON.parse(catRaw.replace(/'/g, '"')) : [catRaw];
      }
    } catch (err) { categories = ["Uncategorized"]; }

    return {
      questionId: row[0],
      problem: row[1],
      img: row[2],
      choices: row[3],
      answer: row[4],
      explain: row[5],
      category: categories
    };
  }).filter(function (q) {
    if (cleanFilter === "") return true;
    return q.category.some(function (catId) {
      return (categoryToSubjectMap[catId] || "") === cleanFilter;
    });
  });

  return ContentService.createTextOutput(JSON.stringify(questions)).setMimeType(ContentService.MimeType.JSON);
}

function getChangedSinceTimestamp(sinceStr, filterSubject) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sinceMs = parseInt(sinceStr) || 0;
  var cleanFilter = filterSubject ? String(filterSubject).trim().toUpperCase() : "";

  // --- Build Category → Subject map ---
  var catToSubjectMap = getCategoryToSubjectMapCached(ss);

  // --- Scan Logs sheet using cached or fresh data ---
  var logDataJson = getLargeCache("logs_data_cache");
  var logData;
  if (logDataJson) {
    logData = JSON.parse(logDataJson);
  } else {
    var logSheet = ss.getSheetByName('Logs');
    if (logSheet) {
      var lastRow = logSheet.getLastRow();
      if (lastRow > 1) {
        // Optimistically read only the last 1000 rows first (milliseconds read)
        var numRowsToRead = Math.min(1000, lastRow - 1);
        var startRow = lastRow - numRowsToRead + 1;
        var sampleData = logSheet.getRange(startRow, 1, numRowsToRead, 10).getValues();

        var oldestSampleTime = sampleData.length > 0 ? new Date(sampleData[0][0]).getTime() : 0;

        if (sinceMs > 0 && oldestSampleTime <= sinceMs) {
          // Excellent! The last 1000 rows fully cover the timeframe since 'sinceMs'
          logData = [[]].concat(sampleData); // Prepend dummy header to match 1-based index offsets
        } else {
          // Fallback to full read only if client has no sinceMs or is extremely outdated
          logData = logSheet.getDataRange().getValues();
        }
        putLargeCache("logs_data_cache", JSON.stringify(logData), 15); // Cache for 15s to block stamps
      } else {
        logData = [];
      }
    } else {
      logData = [];
    }
  }

  var changedIds = {};

  if (logData.length > 1) {
    for (var i = 1; i < logData.length; i++) {
      var logTime = new Date(logData[i][0]).getTime();
      var actionGroup = String(logData[i][3]);
      var targetId = String(logData[i][5]).trim();

      if (logTime > sinceMs && actionGroup === 'QUESTION' && targetId) {
        changedIds[targetId] = true;
      }
    }
  }

  var changedIdKeys = Object.keys(changedIds);

  if (changedIdKeys.length === 0) {
    return ContentService.createTextOutput(JSON.stringify({
      changed: [],
      serverTime: new Date().getTime(),
      count: 0
    })).setMimeType(ContentService.MimeType.JSON);
  }

  // --- Fetch changed rows from cached Questions instead of sheet ---
  var qData = getAllQuestionsCached(ss);
  var changedQuestions = [];

  for (var i = 0; i < qData.length; i++) {
    var qId = String(qData[i][0]).trim();
    if (!changedIds[qId]) continue;

    var categories = [];
    try {
      var catRaw = qData[i][6].toString().trim();
      if (catRaw !== "") {
        categories = (catRaw.indexOf("[") > -1)
          ? JSON.parse(catRaw.replace(/'/g, '"'))
          : [catRaw];
      }
    } catch (e) { categories = ["Uncategorized"]; }

    if (cleanFilter !== "") {
      var inSubject = categories.some(function (catId) {
        return (catToSubjectMap[catId] || "") === cleanFilter;
      });
      if (!inSubject) continue;
    }

    changedQuestions.push({
      questionId: qData[i][0],
      problem: qData[i][1],
      img: qData[i][2],
      choices: qData[i][3],
      answer: qData[i][4],
      explain: qData[i][5],
      category: categories
    });
  }

  return ContentService.createTextOutput(JSON.stringify({
    changed: changedQuestions,
    serverTime: new Date().getTime(),
    count: changedQuestions.length
  })).setMimeType(ContentService.MimeType.JSON);
}

/* 
   =========================================
   ส่วนที่ 2: การบันทึกข้อมูล (POST)
   =========================================
*/
function doPost(e) {
  try {
    var doc = SpreadsheetApp.openById(SHEET_ID);
    var contents = e.postData.contents;
    var data = JSON.parse(contents);
    var action = data.action;

    // ----------------------------------------------------
    // LOCK-FREE GROUP (Processes instantly, no write queue overhead)
    // ----------------------------------------------------
    if (action === 'verifySession') {
      var userObj = verifySessionToken(data.sessionToken);
      if (userObj) {
        return ContentService.createTextOutput(JSON.stringify({
          'result': 'success',
          'user': userObj
        })).setMimeType(ContentService.MimeType.JSON);
      } else {
        return ContentService.createTextOutput(JSON.stringify({
          'result': 'error',
          'message': 'session_expired'
        })).setMimeType(ContentService.MimeType.JSON);
      }
    }

    // Seed/อัปเดต IntelSphere key หนึ่งใบ (idempotent by API_Key) — one-off admin setup, lock-free
    if (action === 'seedIntelSphereKey') {
      return seedIntelSphereKey(data.apiKey, data.donorName, data.notes);
    }

    // อ่าน catalog โมเดล IntelSphere (read-only, ไม่ต้อง auth) — lock-free
    if (action === 'listModels') {
      return ContentService.createTextOutput(JSON.stringify({
        result: 'success',
        catalog: getIntelSphereModelCatalog()
      })).setMimeType(ContentService.MimeType.JSON);
    }

    if (action === 'askAIExpert') {
      // --- IntelSphere shared-pool branch: public, rate-limited, ไม่ใช้ admin auth ---
      if (data.provider === "IntelSphere") {
        var rlToken = data.sessionToken || "guest_user";
        if (!checkRateLimit(rlToken)) {
          return ContentService.createTextOutput(JSON.stringify({
            result: 'error',
            message: 'คุณส่งคำถามถึง AI เกินกำหนด (สูงสุด 15 ครั้งต่อชั่วโมง) กรุณารอสักครู่แล้วลองใหม่ เพื่อช่วยแบ่งโควต้าให้เพื่อนๆ ด้วยนะครับ'
          })).setMimeType(ContentService.MimeType.JSON);
        }

        var isModel = data.model; // ไม่มี default — frontend ส่งโมเดลที่เลือกจาก catalog เสมอ
        if (!isModel) {
          return ContentService.createTextOutput(JSON.stringify({
            result: 'error', message: 'กรุณาเลือกโมเดล AI ก่อนส่งคำถาม'
          })).setMimeType(ContentService.MimeType.JSON);
        }

        try {
          var aiResult = executeChatbotQuery(data.prompt, isModel, 1);
          return ContentService.createTextOutput(JSON.stringify({
            result: 'success', answer: aiResult.content, servedModel: aiResult.servedModel, switched: aiResult.switched
          })).setMimeType(ContentService.MimeType.JSON);
        } catch (isErr) {
          return ContentService.createTextOutput(JSON.stringify({
            result: 'error', message: isErr.message
          })).setMimeType(ContentService.MimeType.JSON);
        }
      }

      // --- เส้นทาง admin เดิม (Gemini pool) ---
      var userObj = null;
      if (data.sessionToken) {
        userObj = verifySessionToken(data.sessionToken);
      } else if (data.googleIdToken) {
        var payload = verifyGoogleToken(data.googleIdToken);
        if (payload) userObj = findAdminByEmail(payload.email);
      } else {
        userObj = verifyAdmin(data.username, data.adminPass);
      }

      if (!userObj) {
        return ContentService.createTextOutput(JSON.stringify({
          'result': 'error',
          'message': 'Session หมดอายุ กรุณาล็อกอินใหม่'
        })).setMimeType(ContentService.MimeType.JSON);
      }

      var provider = data.provider || "Gemini";
      var apiKeyInfo = getAvailableAIKey(provider);

      if (!apiKeyInfo) {
        return ContentService.createTextOutput(JSON.stringify({
          'result': 'error', 'message': 'ขณะนี้ไม่มี API Key ที่พร้อมใช้งาน (โควต้าเต็มทุก Key หรือยังไม่ได้ตั้งค่า)'
        })).setMimeType(ContentService.MimeType.JSON);
      }

      try {
        var aiResponse = callGeminiAI(data.prompt, apiKeyInfo, data.images);
        return ContentService.createTextOutput(JSON.stringify({
          'result': 'success',
          'answer': aiResponse,
          'quota': (apiKeyInfo.usage + 1) + "/" + apiKeyInfo.limit
        })).setMimeType(ContentService.MimeType.JSON);
      } catch (err) {
        return ContentService.createTextOutput(JSON.stringify({
          'result': 'error', 'message': err.message
        })).setMimeType(ContentService.MimeType.JSON);
      }
    }

    // Helper function for mapping batch logs to rows
    function toActivityRow(entry) {
      return [
        entry.timestamp ? new Date(entry.timestamp) : new Date(),
        entry.session || "N/A",
        entry.action || "",
        entry.target || "",
        entry.result || "",
        entry.timeSpent || 0,
        entry.metadata || ""
      ];
    }

    // ----------------------------------------------------
    // T0.1: batchLog — จัดการ "ก่อน" ขอ Lock ใดๆ เพื่อไม่ให้ analytics (write ถี่สุดของนักเรียน) ไปบล็อกโหวต/รายงาน
    // ใช้ appendRow ต่อแถว (atomic ในตัว ไม่ต้องพึ่ง LockService) แทน getRange(getLastRow()+1).setValues()
    // ----------------------------------------------------
    if (action === 'batchLog') {
      var logs = data.logs || [];
      if (logs.length > 0) {
        var activitySheet = doc.getSheetByName("UserActivity") || doc.insertSheet("UserActivity");
        if (activitySheet.getLastRow() === 0) {
          activitySheet.appendRow(["Timestamp", "SessionID", "Action", "TargetID", "Result", "TimeSpent", "Metadata"]);
          activitySheet.getRange(1, 1, 1, 7).setFontWeight("bold").setBackground("#e6f7ff");
        }
        for (var li = 0; li < logs.length; li++) {
          activitySheet.appendRow(toActivityRow(logs[li]));
        }
      }
      return ContentService.createTextOutput(JSON.stringify({ 'result': 'success' })).setMimeType(ContentService.MimeType.JSON);
    }

    // ----------------------------------------------------
    // LOCALIZED LOCK GROUP (Locks briefly for writes, tryLock 15s)
    // ----------------------------------------------------
    var localizedActions = ['submitVote', 'submitReport', 'voteOnReport', 'deleteSession'];
    if (localizedActions.indexOf(action) > -1) {
      var lock = LockService.getScriptLock();
      var acquired = lock.tryLock(15000);
      if (!acquired) {
        return ContentService.createTextOutput(JSON.stringify({
          'result': 'error',
          'message': 'เซิร์ฟเวอร์ไม่ตอบสนองเนื่องจากโหลดสูง (Lock Timeout)'
        })).setMimeType(ContentService.MimeType.JSON);
      }
      var lockReleased = false; // T0.2: ให้ปลด Lock เองก่อนงานหนัก (reconciliation/Gemini) โดย finally ไม่ปลดซ้ำ
      try {
        if (action === 'deleteSession') {
          var token = data.sessionToken;
          if (token) {
            var ss = SpreadsheetApp.openById(SHEET_ID);
            var sheet = ss.getSheetByName("Sessions");
            if (sheet) {
              var rows = sheet.getDataRange().getValues();
              for (var i = rows.length - 1; i >= 1; i--) {
                if (rows[i][0] === token) {
                  sheet.deleteRow(i + 1);
                  break;
                }
              }
            }
          }
          return ContentService.createTextOutput(JSON.stringify({ 'result': 'success' })).setMimeType(ContentService.MimeType.JSON);
        }

        if (action === 'submitVote') {
          var voteSheet = doc.getSheetByName("Votes") || doc.insertSheet("Votes");
          var voteData = voteSheet.getDataRange().getValues();
          var suggestedCategory = data.suggestedCategory || [];
          var delta = data.delta || 1;
          var timestamp = new Date();

          suggestedCategory.forEach(function (category) {
            var foundRowIndex = -1;
            for (var j = 1; j < voteData.length; j++) {
              if (voteData[j][0] == data.questionId && voteData[j][2] == category) {
                foundRowIndex = j + 1;
                break;
              }
            }

            if (foundRowIndex !== -1) {
              var currentVote = parseInt(voteSheet.getRange(foundRowIndex, 4).getValue()) || 0;
              var newVote = currentVote + delta;

              if (newVote < 0) {
                voteSheet.deleteRow(foundRowIndex);
              } else {
                voteSheet.getRange(foundRowIndex, 4).setValue(newVote);
                voteSheet.getRange(foundRowIndex, 5).setValue(timestamp);
              }
            } else if (delta > 0) {
              voteSheet.appendRow([data.questionId, data.questionText, category, 1, timestamp, "Pending"]);
            }
          });
          updateVotesVersion();
          // T0.2: ปลด Lock ก่อนรัน processVotes() (full-sheet reconciliation + sort + updateVersion)
          // เพื่อไม่ให้ค้าง Lock ระหว่างงานหนัก — processVotes ไม่มีการเรียก UrlFetchApp จึงปลอดภัย
          lock.releaseLock();
          lockReleased = true;
          processVotes();
          return ContentService.createTextOutput(JSON.stringify({ 'result': 'success' })).setMimeType(ContentService.MimeType.JSON);
        }

        if (action === 'submitReport') {
          var sheet = doc.getSheetByName("Report") || doc.insertSheet("Report");
          if (sheet.getLastRow() == 0) {
            sheet.appendRow(["From", "Category", "QuestionID", "Question", "Image", "Choices", "SuggestedAnswer", "ReportDetail", "Time", "Status", "AdminNote", "Done", "SuggestedExplain", "VoteCount"]);
          } else {
            var headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
            if (headerRow.indexOf("SuggestedExplain") === -1) {
              sheet.getRange(1, 13).setValue("SuggestedExplain");
              sheet.getRange(1, 14).setValue("VoteCount");
            }
          }

          var qImg = (data.questionImages && data.questionImages.indexOf("http") === 0) ? data.questionImages.split("///")[0] : "";
          var ansSug = data.suggestedChoice || "";

          sheet.appendRow([
            data.from || "User",
            data.category || "",
            data.questionId || "",
            data.question || "",
            qImg,
            data.allChoices || "",
            ansSug,
            data.report || "",
            new Date().toISOString(),
            "Pending",
            "",
            "FALSE",
            data.suggestedExplain || "",
            1
          ]);

          updateVotesVersion();
          return ContentService.createTextOutput(JSON.stringify({ 'result': 'success' })).setMimeType(ContentService.MimeType.JSON);
        }

        if (action === 'voteOnReport') {
          var reportSheet = doc.getSheetByName("Report");
          if (!reportSheet) {
            return ContentService.createTextOutput(JSON.stringify({ result: 'error', message: 'Report sheet not found' })).setMimeType(ContentService.MimeType.JSON);
          }
          var rv = reportSheet.getDataRange().getValues();
          var targetTs = String(data.reportTimestamp || "").trim();
          var delta = parseInt(data.delta) || 1;
          var foundIdx = -1;
          for (var i = 1; i < rv.length; i++) {
            var sTime = rv[i][8] instanceof Date ? rv[i][8].toISOString() : String(rv[i][8]);
            if (sTime.trim() === targetTs) { foundIdx = i; break; }
          }
          if (foundIdx === -1) {
            return ContentService.createTextOutput(JSON.stringify({ result: 'error', message: 'Report not found' })).setMimeType(ContentService.MimeType.JSON);
          }

          // เขียนคะแนนโหวตใหม่ (งานเบาภายใต้ Lock)
          var newVotes = Math.max(0, (parseInt(rv[foundIdx][13]) || 0) + delta);
          reportSheet.getRange(foundIdx + 1, 14).setValue(newVotes);
          updateVotesVersion();

          // T0.2: threshold check เฉพาะแถวนี้ — รัน processReports (ซึ่งอาจเรียก Gemini/UrlFetchApp) เฉพาะเมื่อ
          // รายงานนี้ยัง Pending และแตะเกณฑ์แล้วเท่านั้น และต้องทำ "นอก Lock" เสมอ (ห้ามเรียก UrlFetchApp ใต้ LockService)
          var reportStatus = String(rv[foundIdx][9]).trim();
          var shouldProcess = (reportStatus === "Pending" && newVotes >= REPORT_VOTE_THRESHOLD);

          lock.releaseLock();
          lockReleased = true;
          if (shouldProcess) {
            processReports(doc);
          }
          return ContentService.createTextOutput(JSON.stringify({ result: 'success', newVoteCount: newVotes })).setMimeType(ContentService.MimeType.JSON);
        }
      } finally {
        if (!lockReleased) lock.releaseLock(); // อาจถูกปลดไปแล้วใน submitVote/voteOnReport
      }
    }

    // ----------------------------------------------------
    // ADMIN LOCK GROUP (Admin and write operations, tryLock 25s)
    // ----------------------------------------------------
    var adminLock = LockService.getScriptLock();
    var adminAcquired = adminLock.tryLock(25000);
    if (!adminAcquired) {
      return ContentService.createTextOutput(JSON.stringify({
        'result': 'error',
        'message': 'ระบบหลังบ้านทำงานหนักเนื่องจากมีการเขียนซ้อนกัน (Admin Lock Timeout)'
      })).setMimeType(ContentService.MimeType.JSON);
    }
    try {
      if (action === 'checkGoogleAuth') {
        var tokenPayload = verifyGoogleToken(data.idToken);
        if (!tokenPayload) {
          return ContentService.createTextOutput(JSON.stringify({
            'result': 'error',
            'message': 'Token ไม่ถูกต้องหรือหมดอายุการใช้งาน'
          })).setMimeType(ContentService.MimeType.JSON);
        }

        var email = tokenPayload.email;
        var hd = tokenPayload.hd;

        if (hd !== "kkumail.com" && hd !== "kku.ac.th") {
          return ContentService.createTextOutput(JSON.stringify({
            'result': 'error',
            'message': 'ต้องใช้บัญชี @kkumail.com หรือ @kku.ac.th ของทางมหาวิทยาลัยเท่านั้น'
          })).setMimeType(ContentService.MimeType.JSON);
        }

        var adminUser = findAdminByEmail(email);
        if (adminUser) {
          var sessionToken = createSession(email, adminUser);
          writeAdminLog(adminUser.displayName, adminUser.role, "AUTH", "LOGIN_SSO", "Session", "Google SSO Login Success", "", "", "");
          return ContentService.createTextOutput(JSON.stringify({
            'result': 'success',
            'user': adminUser,
            'sessionToken': sessionToken
          })).setMimeType(ContentService.MimeType.JSON);
        } else {
          writeAdminLog(email, "GUEST", "AUTH", "LOGIN_SSO_FAIL", "Session", "Google login blocked: Email not in whitelist", "", "", "");
          return ContentService.createTextOutput(JSON.stringify({
            'result': 'error',
            'message': 'บัญชีผู้ใช้นี้ไม่มีอยู่ในสิทธิ์การแก้ไขระบบ กรุณาติดต่อผู้ดูแลเพื่อเพิ่มรายชื่ออีเมลของคุณ'
          })).setMimeType(ContentService.MimeType.JSON);
        }
      }

      if (action === 'registerAdmin') {
        var sheet = doc.getSheetByName("Admins");
        var users = sheet.getDataRange().getValues();

        for (var i = 1; i < users.length; i++) {
          if (users[i][0] == data.userData.Username) {
            return ContentService.createTextOutput(JSON.stringify({ 'result': 'error', 'message': 'Username นี้ถูกใช้ไปแล้ว' })).setMimeType(ContentService.MimeType.JSON);
          }
          if (users[i][8] == data.userData.StudentID) {
            return ContentService.createTextOutput(JSON.stringify({ 'result': 'error', 'message': 'รหัสนักศึกษานี้ลงทะเบียนแล้ว' })).setMimeType(ContentService.MimeType.JSON);
          }
        }

        var avatarUrl = "https://api.dicebear.com/7.x/avataaars/svg?seed=" + data.userData.Username;
        if (data.userData.AvatarBase64) {
          try {
            avatarUrl = uploadToDrive(data.userData.AvatarBase64, data.userData.Username + "_avatar.png", "image/png");
          } catch (err) {
            return ContentService.createTextOutput(JSON.stringify({ 'result': 'error', 'message': 'Upload รูปไม่สำเร็จ: ' + err.message })).setMimeType(ContentService.MimeType.JSON);
          }
        }

        sheet.appendRow([
          data.userData.Username,
          data.userData.Password,
          data.userData.DisplayName,
          avatarUrl,
          data.userData.Role || "Admin",
          data.userData.KKUMail,
          data.userData.Prefix,
          data.userData.FullName,
          data.userData.StudentID,
          data.userData.Year,
          data.userData.Contact
        ]);

        updateVersion();
        writeAdminLog("System", "SYSTEM", "AUTH", "REGISTER", data.userData.Username, "New Admin Registered", "", "", "");

        return ContentService.createTextOutput(JSON.stringify({ 'result': 'success' })).setMimeType(ContentService.MimeType.JSON);
      }

      if (action === 'resetPassword') {
        var sheet = doc.getSheetByName("Admins");
        var users = sheet.getDataRange().getValues();
        var found = false;

        for (var i = 1; i < users.length; i++) {
          if (users[i][0] == data.verifyData.Username &&
            users[i][5] == data.verifyData.KKUMail &&
            users[i][8] == data.verifyData.StudentID &&
            users[i][7] == data.verifyData.FullName) {

            sheet.getRange(i + 1, 2).setValue(data.newPassword);
            updateVersion();
            writeAdminLog(data.verifyData.Username, "USER", "AUTH", "RESET_PWD", "Self", "Password Changed via Verification", "", "", "");
            found = true;
            break;
          }
        }

        if (found) {
          return ContentService.createTextOutput(JSON.stringify({ 'result': 'success' })).setMimeType(ContentService.MimeType.JSON);
        } else {
          return ContentService.createTextOutput(JSON.stringify({ 'result': 'error', 'message': 'ข้อมูลยืนยันตัวตนไม่ถูกต้อง' })).setMimeType(ContentService.MimeType.JSON);
        }
      }

      if (action === 'checkAuth') {
        var userObj = verifyAdmin(data.username, data.password);
        if (userObj) {
          writeAdminLog(userObj.username, userObj.role, "AUTH", "LOGIN", "Session", "Login Success", "", "", data.metadata || "");
          return ContentService.createTextOutput(JSON.stringify({
            'result': 'success',
            'user': userObj
          })).setMimeType(ContentService.MimeType.JSON);
        } else {
          writeAdminLog(data.username || "Unknown", "GUEST", "AUTH", "LOGIN_FAIL", "Session", "Login Failed", "", "", data.metadata || "");
          return ContentService.createTextOutput(JSON.stringify({
            'result': 'error',
            'message': 'Username หรือ Password ไม่ถูกต้อง'
          })).setMimeType(ContentService.MimeType.JSON);
        }
      }

      if (action === 'updateAdminProfile') {
        try {
          var ss = SpreadsheetApp.openById(SHEET_ID);
          var sheet = ss.getSheetByName("Admins");
          if (!sheet) throw new Error("ไม่พบแผ่นงาน 'Admins'");

          var values = sheet.getDataRange().getValues();
          var targetUsername = data.targetUsername ? data.targetUsername.toString().trim() : "";
          var foundRow = -1;
          var oldProfileData = {};

          for (var i = 1; i < values.length; i++) {
            if (values[i][0].toString().trim().toLowerCase() === targetUsername.toLowerCase()) {
              foundRow = i + 1;
              oldProfileData = {
                displayName: values[i][2],
                avatarUrl: values[i][3],
                prefix: values[i][6],
                fullName: values[i][7],
                year: values[i][9],
                contact: values[i][10]
              };
              break;
            }
          }

          if (foundRow !== -1) {
            var u = data.updateData;
            if (u.displayName !== undefined) sheet.getRange(foundRow, 3).setValue(u.displayName);
            if (u.avatarUrl !== undefined) sheet.getRange(foundRow, 4).setValue(u.avatarUrl);
            if (u.prefix !== undefined) sheet.getRange(foundRow, 7).setValue(u.prefix);
            if (u.fullName !== undefined) sheet.getRange(foundRow, 8).setValue(u.fullName);
            if (u.year !== undefined) sheet.getRange(foundRow, 10).setValue(u.year);
            if (u.contact !== undefined) sheet.getRange(foundRow, 11).setValue(u.contact);

            updateVersion();
            writeAdminLog(data.username || targetUsername, "ADMIN", "PROFILE", "UPDATE", targetUsername, "Updated profile details", oldProfileData, u, "");

            return ContentService.createTextOutput(JSON.stringify({
              'result': 'success',
              'message': 'Profile updated successfully'
            })).setMimeType(ContentService.MimeType.JSON);
          } else {
            throw new Error("ไม่พบชื่อผู้ใช้: " + targetUsername);
          }
        } catch (error) {
          return ContentService.createTextOutput(JSON.stringify({
            'result': 'error',
            'message': error.toString()
          })).setMimeType(ContentService.MimeType.JSON);
        }
      }

      if (action === 'uploadImage') {
        var userObj = null;
        if (data.sessionToken) {
          userObj = verifySessionToken(data.sessionToken);
        } else if (data.googleIdToken) {
          var payload = verifyGoogleToken(data.googleIdToken);
          if (payload) userObj = findAdminByEmail(payload.email);
        } else {
          userObj = verifyAdmin(data.username, data.adminPass);
        }

        if (!userObj) {
          return ContentService.createTextOutput(JSON.stringify({
            'result': 'error',
            'message': 'token_expired'
          })).setMimeType(ContentService.MimeType.JSON);
        }

        try {
          var subject = data.data.subject;
          var year = data.data.year;
          var fileUrl = uploadQuestionImageToDrive(data.data.base64, data.data.questionId, data.data.type, subject, year);
          writeAdminLog(userObj.username, userObj.role, "IMAGE", "UPLOAD", data.data.questionId, "Uploaded new " + data.data.type + " image", "", fileUrl, "");

          return ContentService.createTextOutput(JSON.stringify({
            'result': 'success',
            'url': fileUrl
          })).setMimeType(ContentService.MimeType.JSON);

        } catch (err) {
          return ContentService.createTextOutput(JSON.stringify({
            'result': 'error',
            'message': 'Drive Upload Error: ' + err.message
          })).setMimeType(ContentService.MimeType.JSON);
        }
      }

      // T2.5: อัปโหลดหลายรูปในการเรียกครั้งเดียว (สูงสุด 10 รูป) — auth และพารามิเตอร์ต่อรายการเหมือน uploadImage
      // แต่ละรายการอยู่ใน data.images[] = { base64, questionId, type, subject, year }
      // คืน urls[] เรียงตามลำดับ input; รายการที่ล้มเหลวจะเป็น { error: "..." } (ไม่ทำให้ทั้ง batch ล้ม)
      if (action === 'uploadImagesBatch') {
        var userObj = null;
        if (data.sessionToken) {
          userObj = verifySessionToken(data.sessionToken);
        } else if (data.googleIdToken) {
          var payload = verifyGoogleToken(data.googleIdToken);
          if (payload) userObj = findAdminByEmail(payload.email);
        } else {
          userObj = verifyAdmin(data.username, data.adminPass);
        }

        if (!userObj) {
          return ContentService.createTextOutput(JSON.stringify({
            'result': 'error',
            'message': 'token_expired'
          })).setMimeType(ContentService.MimeType.JSON);
        }

        var images = data.images || [];
        if (!Array.isArray(images) || images.length === 0) {
          return ContentService.createTextOutput(JSON.stringify({
            'result': 'error',
            'message': 'ไม่พบรายการรูปภาพ (images array is empty)'
          })).setMimeType(ContentService.MimeType.JSON);
        }
        if (images.length > 10) {
          return ContentService.createTextOutput(JSON.stringify({
            'result': 'error',
            'message': 'อัปโหลดได้สูงสุด 10 รูปต่อครั้ง (batch size exceeds 10)'
          })).setMimeType(ContentService.MimeType.JSON);
        }

        var urls = [];
        var successCount = 0;
        for (var bi = 0; bi < images.length; bi++) {
          var item = images[bi] || {};
          try {
            if (!item.base64) { urls.push({ error: 'missing base64' }); continue; }
            var fileUrl = uploadQuestionImageToDrive(item.base64, item.questionId, item.type, item.subject, item.year);
            urls.push(fileUrl);
            successCount++;
          } catch (err) {
            urls.push({ error: err.message });
          }
        }

        writeAdminLog(userObj.username, userObj.role, "IMAGE", "UPLOAD_BATCH",
          (images[0] && images[0].questionId) || "",
          "Batch uploaded " + successCount + "/" + images.length + " images", "", "", "");

        return ContentService.createTextOutput(JSON.stringify({
          'result': 'success',
          'urls': urls
        })).setMimeType(ContentService.MimeType.JSON);
      }

      if (action === 'deleteImage') {
        var userObj = verifyUser(data);
        if (!userObj) {
          return ContentService.createTextOutput(JSON.stringify({ 'result': 'error', 'message': 'Session หมดอายุ หรือไม่ได้รับอนุญาตให้เข้าถึง' })).setMimeType(ContentService.MimeType.JSON);
        }

        try {
          var fileUrl = data.data.url;
          var currentQid = String(data.data.currentQid || "").trim();

          var match = fileUrl.match(/id=([^&]+)/) || fileUrl.match(/\/d\/([^\/]+)/);
          if (!match) throw new Error("ไม่สามารถระบุ ID ของไฟล์จาก URL นี้ได้");
          var fileId = match[1];

          var qSheet = doc.getSheetByName("Questions");
          var qData = qSheet.getDataRange().getValues();
          var otherUsages = [];

          for (var i = 1; i < qData.length; i++) {
            var qid = String(qData[i][0]).trim();
            if (qid === currentQid) continue;

            var imgCol = String(qData[i][2]);
            var choiceCol = String(qData[i][3]);

            if (imgCol.indexOf(fileId) !== -1 || choiceCol.indexOf(fileId) !== -1) {
              otherUsages.push({ qid: qid, type: imgCol.indexOf(fileId) !== -1 ? 'Main' : 'Choice' });
            }
          }

          if (otherUsages.length > 0) {
            var nextOwner = otherUsages[0];
            var file = DriveApp.getFileById(fileId);
            var timestamp = new Date().getTime();
            var originalName = file.getName();
            var ext = originalName.substring(originalName.lastIndexOf('.')) || ".png";
            var newName = "Q_" + nextOwner.qid + "_" + nextOwner.type + "_" + timestamp + ext;
            file.setName(newName);

            writeAdminLog(userObj.username, userObj.role, "IMAGE", "TRANSFER", fileId, "ภาพยังถูกใช้โดยข้อ " + nextOwner.qid + " จึงแค่เปลี่ยนชื่อไฟล์", "", "", "");

            return ContentService.createTextOutput(JSON.stringify({
              'result': 'success', 'message': 'ปลดลิงก์สำเร็จ (ไฟล์ยังคงอยู่เพราะข้อ ' + nextOwner.qid + ' ใช้งานอยู่)'
            })).setMimeType(ContentService.MimeType.JSON);

          } else {
            var resultMessage = deleteImageToRecycleBin(fileUrl, userObj.username);
            writeAdminLog(userObj.username, userObj.role, "IMAGE", "DELETE", fileId, "ย้ายรูปลง Recycle Bin", "", "TRASHED", "");

            return ContentService.createTextOutput(JSON.stringify({
              'result': 'success', 'message': resultMessage
            })).setMimeType(ContentService.MimeType.JSON);
          }

        } catch (err) {
          return ContentService.createTextOutput(JSON.stringify({ 'result': 'error', 'message': 'ลบรูปภาพขัดข้อง: ' + err.message })).setMimeType(ContentService.MimeType.JSON);
        }
      }

      if (action === 'restoreImage') {
        var userObj = verifyUser(data.username, data.adminPass);
        if (!userObj) {
          return ContentService.createTextOutput(JSON.stringify({
            'result': 'error',
            'message': 'Session หมดอายุ หรือสิทธิ์ไม่ถูกต้อง กรุณาล็อกอินใหม่'
          })).setMimeType(ContentService.MimeType.JSON);
        }

        try {
          var fileUrl = data.data.url;
          var resultMessage = restoreImageFromRecycleBin(fileUrl);
          writeAdminLog(userObj.username, userObj.role, "IMAGE", "RESTORE", fileUrl, "Restored image from Recycle Bin", "TRASHED", "RESTORED", "");

          return ContentService.createTextOutput(JSON.stringify({
            'result': 'success',
            'url': fileUrl,
            'message': resultMessage
          })).setMimeType(ContentService.MimeType.JSON);

        } catch (err) {
          return ContentService.createTextOutput(JSON.stringify({
            'result': 'error',
            'message': 'ไม่สามารถกู้คืนรูปภาพได้ (อาจจะไม่มีอยู่ในถังขยะแล้ว): ' + err.message
          })).setMimeType(ContentService.MimeType.JSON);
        }
      }

      var adminActions = ['editQuestion', 'deleteQuestion', 'addCategory', 'adminImport', 'updateReportStatus', 'deleteCategory', 'updateCategory', 'deleteGroup', 'updateAccordionGroup', 'addSubject', 'updateSubject', 'deleteSubject', 'addAnnouncement', 'editAnnouncement', 'deleteAnnouncement'];
      if (adminActions.indexOf(action) > -1) {
        var userObj = null;
        if (data.sessionToken) {
          userObj = verifySessionToken(data.sessionToken);
        } else if (data.googleIdToken) {
          var payload = verifyGoogleToken(data.googleIdToken);
          if (payload) userObj = findAdminByEmail(payload.email);
        } else {
          userObj = verifyAdmin(data.username, data.adminPass);
        }

        if (!userObj) {
          return ContentService.createTextOutput(JSON.stringify({
            'result': 'error',
            'message': 'token_expired'
          })).setMimeType(ContentService.MimeType.JSON);
        }

        var user = userObj.displayName || "Unknown Admin";
        var userRole = userObj.role || "Admin";
        var metadata = data.metadata || "";

        var sheet;

        if (action === 'addAnnouncement') {
          sheet = doc.getSheetByName("Announcements");
          sheet.appendRow([
            data.data.Id,
            data.data.Text,
            data.data.Type,
            data.data.Active,
            data.data.Order
          ]);
          updateVersion();
          writeAdminLog(user, userRole, "ANNOUNCEMENT", "ADD", data.data.Id, "Added Announcement", "", data.data.Text, metadata);
          return ContentService.createTextOutput(JSON.stringify({ 'result': 'success' })).setMimeType(ContentService.MimeType.JSON);
        }

        if (action === 'editAnnouncement') {
          sheet = doc.getSheetByName("Announcements");
          var rows = sheet.getDataRange().getValues();
          for (var i = 1; i < rows.length; i++) {
            if (rows[i][0] == data.data.Id) {
              var oldText = rows[i][1];
              sheet.getRange(i + 1, 2, 1, 4).setValues([[
                data.data.Text,
                data.data.Type,
                data.data.Active,
                data.data.Order
              ]]);
              updateVersion();
              writeAdminLog(user, userRole, "ANNOUNCEMENT", "EDIT", data.data.Id, "Updated Announcement", oldText, data.data.Text, metadata);
              return ContentService.createTextOutput(JSON.stringify({ 'result': 'success' })).setMimeType(ContentService.MimeType.JSON);
            }
          }
        }

        if (action === 'deleteAnnouncement') {
          sheet = doc.getSheetByName("Announcements");
          var rows = sheet.getDataRange().getValues();
          for (var i = 1; i < rows.length; i++) {
            if (rows[i][0] == data.data.Id) {
              var oldText = rows[i][1];
              sheet.deleteRow(i + 1);
              updateVersion();
              writeAdminLog(user, userRole, "ANNOUNCEMENT", "DELETE", data.data.Id, "Deleted Announcement", oldText, "DELETED", metadata);
              return ContentService.createTextOutput(JSON.stringify({ 'result': 'success' })).setMimeType(ContentService.MimeType.JSON);
            }
          }
        }

        if (action === 'editQuestion') {
          sheet = doc.getSheetByName("Questions");
          var rows = sheet.getDataRange().getValues();
          var headers = rows[0];

          for (var i = 1; i < rows.length; i++) {
            if (rows[i][0] == data.data.id) {
              var oldRowData = {};
              for (var k = 0; k < headers.length; k++) {
                oldRowData[headers[k]] = rows[i][k];
              }

              var catToSave = Array.isArray(data.data.category) ? JSON.stringify(data.data.category) : data.data.category;
              sheet.getRange(i + 1, 2, 1, 6).setValues([
                [data.data.problem, data.data.img, data.data.choices, data.data.answer, data.data.explain, catToSave]
              ]);

              try {
                const catsForSplit = Array.isArray(data.data.category) ? data.data.category : JSON.parse(catToSave);
                autoCreateSplitCategories(data.data.id, catsForSplit);
              } catch (e) { console.log("Split error in editQuestion: " + e); }

              updateVersion();
              writeAdminLog(user, userRole, "QUESTION", "EDIT", data.data.id, "Question Updated", oldRowData, data.data, metadata);

              return ContentService.createTextOutput(JSON.stringify({
                'result': 'success'
              })).setMimeType(ContentService.MimeType.JSON);
            }
          }
        }

        if (action === 'deleteQuestion') {
          sheet = doc.getSheetByName("Questions");
          var rows = sheet.getDataRange().getValues();
          var headers = rows[0];

          for (var i = 1; i < rows.length; i++) {
            if (rows[i][0] == data.data.id) {
              var oldRowData = {};
              for (var k = 0; k < headers.length; k++) {
                oldRowData[headers[k]] = rows[i][k];
              }

              sheet.deleteRow(i + 1);
              updateVersion();
              writeAdminLog(user, userRole, "QUESTION", "DELETE", data.data.id, "Question Deleted", oldRowData, "DELETED", metadata);

              return ContentService.createTextOutput(JSON.stringify({
                'result': 'success'
              })).setMimeType(ContentService.MimeType.JSON);
            }
          }
        }

        if (action === 'adminImport') {
          var sheetMap = {
            'struct': 'Structure',
            'category': 'Category',
            'ques': 'Questions'
          };

          var payload = data.data;
          var realSheetName = sheetMap[payload.sheetName] || payload.sheetName;
          var targetSheet = doc.getSheetByName(realSheetName);

          if (!targetSheet) {
            return ContentService.createTextOutput(JSON.stringify({
              'result': 'error',
              'message': 'ไม่พบแผ่นงาน: ' + realSheetName
            })).setMimeType(ContentService.MimeType.JSON);
          }

          var importData = payload.data;
          if (!importData || importData.length === 0) {
            return ContentService.createTextOutput(JSON.stringify({
              'result': 'error',
              'message': 'ข้อมูลว่างเปล่า'
            })).setMimeType(ContentService.MimeType.JSON);
          }

          try {
            var lastRow = targetSheet.getLastRow();

            if (realSheetName === 'Questions') {
              var appended = 0, updated = 0;
              var toAppend = [];

              // T2.4: หลีกเลี่ยง setValues ต่อแถวสำหรับข้อที่อัปเดต — อ่าน block เดียว แก้ใน memory แล้วเขียนกลับครั้งเดียว
              // เงื่อนไข: ทุกแถวนำเข้าต้องกว้างเท่ากัน (uniform) จึงเขียนเป็นบล็อกสี่เหลี่ยมได้อย่างปลอดภัย
              var width = importData[0].length;
              var uniform = importData.every(function (r) { return r.length === width; });

              if (uniform) {
                // อ่านบล็อกที่มีอยู่ครั้งเดียว (กว้าง = width) เพื่อคงคอลัมน์ส่วนเกิน (ถ้ามี) ไว้ไม่ถูกแตะ
                var block = lastRow > 1 ? targetSheet.getRange(2, 1, lastRow - 1, width).getValues() : [];
                var idToIdx = {};
                for (var bi = 0; bi < block.length; bi++) idToIdx[String(block[bi][0]).trim()] = bi;

                importData.forEach(function (row) {
                  var qId = String(row[0]).trim();
                  if (idToIdx.hasOwnProperty(qId)) {
                    block[idToIdx[qId]] = row; // อัปเดตใน memory
                    updated++;
                  } else {
                    idToIdx[qId] = block.length; // กันซ้ำภายใน payload เดียวกัน
                    block.push(row);
                    toAppend.push(row);
                  }
                });

                if (block.length > 0) {
                  // เขียนทั้งบล็อก (updates + appends) ในครั้งเดียว
                  targetSheet.getRange(2, 1, block.length, width).setValues(block);
                }
                appended = toAppend.length;
              } else {
                // Fallback (แถวกว้างไม่เท่ากัน): upsert ต่อแถวแบบเดิม เพื่อความปลอดภัยของข้อมูล
                var existing = lastRow > 1 ? targetSheet.getRange(2, 1, lastRow - 1, 1).getValues() : [];
                var existingIds = existing.map(function (r) { return String(r[0]).trim(); });
                importData.forEach(function (row) {
                  var qId = String(row[0]).trim();
                  var idx = existingIds.indexOf(qId);
                  if (idx >= 0) {
                    targetSheet.getRange(idx + 2, 1, 1, row.length).setValues([row]);
                    updated++;
                  } else {
                    toAppend.push(row);
                    existingIds.push(qId);
                  }
                });
                if (toAppend.length > 0) {
                  var newLastRow = targetSheet.getLastRow();
                  targetSheet.getRange(newLastRow + 1, 1, toAppend.length, toAppend[0].length).setValues(toAppend);
                }
                appended = toAppend.length;
              }

              updateVersion();
              writeAdminLog(user, userRole, "DATA", "IMPORT", realSheetName,
                "Upserted Questions: " + appended + " added, " + updated + " updated", "",
                "Added " + appended + ", Updated " + updated, metadata);

              return ContentService.createTextOutput(JSON.stringify({
                'result': 'success',
                'count': appended + updated,
                'added': appended,
                'updated': updated,
                'message': 'นำเข้าสำเร็จ: เพิ่ม ' + appended + ' แถว, อัปเดต ' + updated + ' แถว'
              })).setMimeType(ContentService.MimeType.JSON);

            } else {
              var existingKeys = new Set();
              if (lastRow > 0) {
                var fullData = targetSheet.getDataRange().getValues();
                for (var i = 0; i < fullData.length; i++) {
                  if (realSheetName === 'Structure') {
                    existingKeys.add(fullData[i][1] + "|" + fullData[i][3]);
                  } else {
                    existingKeys.add(String(fullData[i][0]));
                  }
                }
              }

              var finalData = importData.filter(function (row) {
                var key = realSheetName === 'Structure' ? row[1] + "|" + row[3] : String(row[0]);
                return !existingKeys.has(key);
              });

              if (finalData.length > 0) {
                targetSheet.getRange(lastRow + 1, 1, finalData.length, finalData[0].length).setValues(finalData);
                updateVersion();
                writeAdminLog(user, userRole, "DATA", "IMPORT", realSheetName, "Imported " + finalData.length + " new rows (Skipped " + (importData.length - finalData.length) + " duplicates)", "", "Added " + finalData.length + " rows", metadata);

                return ContentService.createTextOutput(JSON.stringify({
                  'result': 'success',
                  'count': finalData.length,
                  'skipped': importData.length - finalData.length,
                  'message': 'นำเข้าสำเร็จ ' + finalData.length + ' แถว (ข้ามข้อมูลซ้ำ ' + (importData.length - finalData.length) + ' แถว)'
                })).setMimeType(ContentService.MimeType.JSON);
              } else {
                return ContentService.createTextOutput(JSON.stringify({
                  'result': 'success',
                  'count': 0,
                  'skipped': importData.length,
                  'message': 'ไม่มีข้อมูลใหม่ให้นำเข้า (ข้อมูลทั้งหมดมีอยู่แล้วในระบบ)'
                })).setMimeType(ContentService.MimeType.JSON);
              }
            }
          } catch (err) {
            return ContentService.createTextOutput(JSON.stringify({
              'result': 'error',
              'message': 'GS Error: ' + err.toString()
            })).setMimeType(ContentService.MimeType.JSON);
          }
        }

        if (action === 'updateReportStatus') {
          sheet = doc.getSheetByName("Report");
          var rows = sheet.getDataRange().getValues();
          for (var i = 1; i < rows.length; i++) {
            var sTime = rows[i][8] instanceof Date ? rows[i][8].toISOString() : String(rows[i][8]);

            if (sTime === String(data.data.timestamp)) {
              var oldStatus = rows[i][9];

              sheet.getRange(i + 1, 10, 1, 3).setValues([
                [data.data.status, data.data.adminNote, data.data.done]
              ]);

              updateVersion();
              writeAdminLog(user, userRole, "REPORT", "UPDATE", "Report_Row_" + (i + 1), "Updated Report Status", oldStatus, data.data.status, metadata);

              return ContentService.createTextOutput(JSON.stringify({ 'result': 'success' })).setMimeType(ContentService.MimeType.JSON);
            }
          }
        }

        if (action === 'deleteCategory') {
          var sheet = doc.getSheetByName("Category");
          var rows = sheet.getDataRange().getValues();
          for (var i = 1; i < rows.length; i++) {
            if (rows[i][0] == data.data.CategoryID) {
              var catNameOld = rows[i][3];
              sheet.deleteRow(i + 1);
              updateVersion();
              writeAdminLog(user, userRole, "CATEGORY", "DELETE", data.data.CategoryID, "Category Deleted", catNameOld, "DELETED", metadata);
              return ContentService.createTextOutput(JSON.stringify({
                'result': 'success'
              })).setMimeType(ContentService.MimeType.JSON);
            }
          }
        }

        if (action === 'updateCategory') {
          var sheet = doc.getSheetByName("Category");
          var rows = sheet.getDataRange().getValues();
          for (var i = 1; i < rows.length; i++) {
            if (rows[i][0] == data.data.CategoryID) {
              var oldName = rows[i][3];
              sheet.getRange(i + 1, 4).setValue(data.data.CategoryName);
              updateVersion();
              writeAdminLog(user, userRole, "CATEGORY", "EDIT", data.data.CategoryID, "Renamed Category", oldName, data.data.CategoryName, metadata);
              return ContentService.createTextOutput(JSON.stringify({
                'result': 'success'
              })).setMimeType(ContentService.MimeType.JSON);
            }
          }
        }

        if (action === 'deleteGroup') {
          var sheet = doc.getSheetByName("Category");
          var rows = sheet.getDataRange().getValues();
          var deletedCount = 0;
          for (var i = rows.length - 1; i >= 1; i--) {
            if (rows[i][1] == data.data.SubjectRef && rows[i][2] == data.data.AccordionGroup) {
              sheet.deleteRow(i + 1);
              deletedCount++;
            }
          }

          var structSheet = doc.getSheetByName("Structure");
          var sRows = structSheet.getDataRange().getValues();
          for (var i = sRows.length - 1; i >= 1; i--) {
            if (sRows[i][1] == data.data.SubjectRef && sRows[i][3] == data.data.AccordionGroup) {
              structSheet.deleteRow(i + 1);
            }
          }

          updateVersion();
          writeAdminLog(user, userRole, "GROUP", "DELETE", data.data.SubjectRef + "_" + data.data.AccordionGroup, "Deleted Group & " + deletedCount + " categories", "", "DELETED", metadata);
          return ContentService.createTextOutput(JSON.stringify({
            'result': 'success'
          })).setMimeType(ContentService.MimeType.JSON);
        }

        if (action === 'updateAccordionGroup') {
          var oldGroup = data.data.OldAccordionGroup;
          var newGroup = data.data.NewAccordionGroup;
          var subjectId = data.data.SubjectRef;
          var updatedCount = 0;

          var catSheet = doc.getSheetByName("Category");
          var cRows = catSheet.getDataRange().getValues();
          for (var i = 1; i < cRows.length; i++) {
            if (cRows[i][1] == subjectId && cRows[i][2] == oldGroup) {
              catSheet.getRange(i + 1, 3).setValue(newGroup);
              updatedCount++;
            }
          }

          var structSheet = doc.getSheetByName("Structure");
          var sRows = structSheet.getDataRange().getValues();
          for (var i = 1; i < sRows.length; i++) {
            if (sRows[i][1] == subjectId && sRows[i][3] == oldGroup) {
              structSheet.getRange(i + 1, 4).setValue(newGroup);
            }
          }

          updateVersion();
          writeAdminLog(user, userRole, "GROUP", "EDIT", subjectId + "_" + oldGroup, "Renamed Group", oldGroup, newGroup, metadata);
          return ContentService.createTextOutput(JSON.stringify({
            'result': 'success',
            'message': 'Updated ' + updatedCount + ' categories'
          })).setMimeType(ContentService.MimeType.JSON);
        }

        if (action === 'addSubject') {
          sheet = doc.getSheetByName("Structure");
          sheet.appendRow([data.data.Year, data.data.SubjectID, data.data.SubjectName, "GENERAL"]);
          updateVersion();
          writeAdminLog(user, userRole, "SUBJECT", "ADD", data.data.SubjectID, "Added Subject", "", data.data.SubjectName, metadata);
          return ContentService.createTextOutput(JSON.stringify({ 'result': 'success' })).setMimeType(ContentService.MimeType.JSON);
        }

        if (action === 'updateSubject') {
          sheet = doc.getSheetByName("Structure");
          var rows = sheet.getDataRange().getValues();
          for (var i = 1; i < rows.length; i++) {
            if (rows[i][1] == data.data.SubjectID) {
              var oldName = rows[i][2];
              sheet.getRange(i + 1, 1).setValue(data.data.Year);
              sheet.getRange(i + 1, 3).setValue(data.data.SubjectName);
              updateVersion();
              writeAdminLog(user, userRole, "SUBJECT", "EDIT", data.data.SubjectID, "Updated Subject Info", oldName, data.data.SubjectName, metadata);
              return ContentService.createTextOutput(JSON.stringify({ 'result': 'success' })).setMimeType(ContentService.MimeType.JSON);
            }
          }
        }

        if (action === 'deleteSubject') {
          var subjectId = data.data.SubjectID;
          var structSheet = doc.getSheetByName("Structure");
          var sRows = structSheet.getDataRange().getValues();
          for (var i = sRows.length - 1; i >= 1; i--) {
            if (sRows[i][1] == subjectId) structSheet.deleteRow(i + 1);
          }
          var catSheet = doc.getSheetByName("Category");
          var cRows = catSheet.getDataRange().getValues();
          for (var i = cRows.length - 1; i >= 1; i--) {
            if (cRows[i][1] == subjectId) catSheet.deleteRow(i + 1);
          }
          updateVersion();
          writeAdminLog(user, userRole, "SUBJECT", "DELETE", subjectId, "Deleted Subject & Related Data", "", "DELETED", metadata);
          return ContentService.createTextOutput(JSON.stringify({ 'result': 'success' })).setMimeType(ContentService.MimeType.JSON);
        }

        if (action === 'addCategory') {
          sheet = doc.getSheetByName("Category");
          sheet.appendRow([data.data.CategoryID, data.data.SubjectRef, data.data.AccordionGroup, data.data.CategoryName]);

          var structSheet = doc.getSheetByName("Structure");
          var sRows = structSheet.getDataRange().getValues();
          var groupExists = false;
          for (var i = 1; i < sRows.length; i++) {
            if (sRows[i][1] == data.data.SubjectRef && sRows[i][3] == data.data.AccordionGroup) {
              groupExists = true;
              break;
            }
          }

          if (!groupExists) {
            var year = "";
            var subjectName = "";
            for (var i = 1; i < sRows.length; i++) {
              if (sRows[i][1] == data.data.SubjectRef) {
                year = sRows[i][0];
                subjectName = sRows[i][2];
                break;
              }
            }
            if (subjectName) {
              structSheet.appendRow([year, data.data.SubjectRef, subjectName, data.data.AccordionGroup]);
            }
          }

          updateVersion();
          sortCategorySheet();
          writeAdminLog(user, userRole, "CATEGORY", "ADD", data.data.CategoryID, "Added Category", "", data.data.CategoryName, metadata);
          return ContentService.createTextOutput(JSON.stringify({ 'result': 'success' })).setMimeType(ContentService.MimeType.JSON);
        }
      }
      return ContentService.createTextOutput(JSON.stringify({
        'result': 'error',
        'message': 'Action "' + action + '" not found or logic failed'
      })).setMimeType(ContentService.MimeType.JSON);
    } finally {
      adminLock.releaseLock();
    }

  } catch (e) {
    return ContentService.createTextOutput(JSON.stringify({
      'result': 'error',
      'message': e.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

/* 
   ========================================================
   ส่วนที่ 3 IMAGE CRUD SYSTEM (Upload, Recycle Bin, Restore)
   ========================================================
*/

// 1. ฟังก์ชันช่วยหาหรือสร้าง Folder (MD > Y[ปี] > [วิชา])
function getOrCreateFolder(parentFolder, folderName) {
  var folders = parentFolder.getFoldersByName(folderName);
  return folders.hasNext() ? folders.next() : parentFolder.createFolder(folderName);
}

// 2. ฟังก์ชันแกะรหัสเพื่อหาที่อยู่โฟลเดอร์
function getQuestionRoutingInfo(questionId) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var qSheet = ss.getSheetByName('Questions');
  var qData = qSheet.getDataRange().getValues();
  var categoryId = "";
  
  for(var i=1; i<qData.length; i++) {
    if(qData[i][0] == questionId) {
      var catRaw = qData[i][6].toString();
      try { categoryId = JSON.parse(catRaw.replace(/'/g, '"'))[0]; } 
      catch(e) { categoryId = catRaw; }
      break;
    }
  }

  var cSheet = ss.getSheetByName('Category');
  var cData = cSheet.getDataRange().getValues();
  var subjectId = "";
  for(var i=1; i<cData.length; i++) {
    if(cData[i][0] == categoryId) { subjectId = cData[i][1]; break; }
  }

  var sSheet = ss.getSheetByName('Structure');
  var sData = sSheet.getDataRange().getValues();
  var year = "Unknown";
  for(var i=1; i<sData.length; i++) {
    if(sData[i][1] == subjectId) { year = sData[i][0]; break; }
  }

  return { year: year, subject: subjectId || "General" };
}

// 3. ฟังก์ชันอัปโหลดรูป
function uploadQuestionImageToDrive(base64Data, questionId, typeIdentifier) {
  var maxRetries = 3;
  var lastError;

  for (var i = 0; i < maxRetries; i++) {
    try {
      var routeInfo = getQuestionRoutingInfo(questionId);
      var rootFolder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
      
      var mdFolder = getOrCreateFolder(rootFolder, "MD");
      var yearFolder = getOrCreateFolder(mdFolder, "Y" + routeInfo.year);
      var targetFolder = getOrCreateFolder(yearFolder, routeInfo.subject);

      // Logic จัดเก็บแยกลง Sub-folder
      if (questionId && questionId.indexOf('_') > -1) {
        var parts = questionId.split('_'); 
        if (parts.length >= 2) {
          var yearType = parts[1].trim(); 
          if (/^\d{2}/.test(yearType)) {
            var examYear = yearType.substring(0, 2); 
            var examGroup = yearType.substring(2) || "General"; 
            var examYearFolder = getOrCreateFolder(targetFolder, examYear);
            targetFolder = getOrCreateFolder(examYearFolder, examGroup); 
          } else {
            targetFolder = getOrCreateFolder(targetFolder, yearType);
          }
        }
      }

      if (typeIdentifier === 'Explain') {
        targetFolder = getOrCreateFolder(targetFolder, 'Explanation');
      }

      var mimeType = "image/png";
      var fileExtension = "png";
      if (base64Data.indexOf("data:") === 0) {
        var partsMime = base64Data.split(";")[0].split(":");
        if (partsMime.length > 1) {
          mimeType = partsMime[1];
          if (mimeType === "application/pdf") {
            fileExtension = "pdf";
          } else {
            fileExtension = mimeType.split("/")[1] || "png";
          }
        }
      }

      var cleanBase64 = base64Data.split(',')[1] || base64Data;
      var uniqueID = new Date().getTime() + "_" + Math.floor(Math.random() * 10000);
      var fileName = "Q_" + questionId + "_" + typeIdentifier + "_" + uniqueID + "." + fileExtension;
      
      var blob = Utilities.newBlob(Utilities.base64Decode(cleanBase64), mimeType, fileName);
      var file = targetFolder.createFile(blob);
      
      if (mimeType === 'application/pdf') {
        return 'https://drive.google.com/file/d/' + file.getId() + '/preview';
      } else {
        return 'https://drive.google.com/uc?export=view&id=' + file.getId();
      }

    } catch (e) {
      lastError = e;
      if (e.message.indexOf("Service error") !== -1 || e.message.indexOf("ข้อผิดพลาดของบริการ") !== -1) {
        Utilities.sleep(1500);
        continue;
      }
      throw e;
    }
  }
  throw new Error("Drive ยังคงไม่ตอบสนองหลังจากพยายาม 3 ครั้ง: " + lastError.message);
}

function handleImageTrashAction(fileUrl, user) {
  try {
    var fileId = fileUrl.match(/id=([^&]+)/);
    if (!fileId) return "Invalid URL";
    
    var file = DriveApp.getFileById(fileId[1]);
    var originalParentId = file.getParents().next().getId();
    var rootFolder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
    var recycleFolder = getOrCreateFolder(rootFolder, "RecycleBin");

    recycleFolder.addFile(file);
    DriveApp.getFolderById(originalParentId).removeFile(file);

    var ss = SpreadsheetApp.openById(SHEET_ID);
    var logSheet = ss.getSheetByName("RecycleLog") || ss.insertSheet("RecycleLog");
    if(logSheet.getLastRow() == 0) logSheet.appendRow(["FileID", "Type", "OriginalParentID", "DeletedDate", "User"]);
    
    logSheet.appendRow([fileId[1], "Image", originalParentId, new Date().toISOString(), user]);
    return "Moved to Recycle Bin";
  } catch (e) {
    throw new Error("Trash Error: " + e.message);
  }
}

// 4. ระบบถังขยะ (Soft Delete)
function deleteImageToRecycleBin(fileUrl, user) {
  try {
    var match = fileUrl.match(/id=([^&]+)/) || fileUrl.match(/\/d\/([^\/]+)/);
    if (!match) throw new Error("ID ไฟล์ไม่ถูกต้อง");
    var fileId = match[1];
    
    var file = DriveApp.getFileById(fileId);
    
    // ค้นหาโฟลเดอร์หลัก
    var rootFolder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
    
    // ค้นหาหรือสร้าง RecycleBin
    var recycleFolder;
    var folders = rootFolder.getFoldersByName("RecycleBin");
    if (folders.hasNext()) {
      recycleFolder = folders.next();
    } else {
      recycleFolder = rootFolder.createFolder("RecycleBin");
    }

    // ย้ายไฟล์
    recycleFolder.addFile(file);
    
    // พยายามดึงไฟล์ออกจากโฟลเดอร์เดิมทั้งหมด
    var firstParentId = "";
    var parents = file.getParents();
    while (parents.hasNext()) {
      var parent = parents.next();
      if (firstParentId === "") firstParentId = parent.getId();
      parent.removeFile(file);
    }

    // เขียนบันทึกประวัติเพื่ออนุญาตให้กู้คืนได้ผ่านระบบ Restore
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var logSheet = ss.getSheetByName("RecycleLog") || ss.insertSheet("RecycleLog");
    if(logSheet.getLastRow() == 0) logSheet.appendRow(["FileID", "Type", "OriginalParentID", "DeletedDate", "User"]);
    logSheet.appendRow([fileId, "ExplainMedia", firstParentId, new Date().toISOString(), user]);

    return "Moved to Recycle Bin successfully";
  } catch (e) {
    throw new Error("Drive Error: " + e.message);
  }
}

// 5. ระบบกู้คืน (Restore)
function restoreImageFromRecycleBin(fileUrl) {
  var match = fileUrl.match(/id=([^&]+)/) || fileUrl.match(/\/d\/([^\/]+)/);
  if (!match) throw new Error("URL ไม่ถูกต้อง");
  var fileId = match[1];

  var ss = SpreadsheetApp.openById(SHEET_ID);
  var logSheet = ss.getSheetByName("RecycleLog");
  if (!logSheet) throw new Error("ไม่พบประวัติการลบรูปภาพ");
  
  var data = logSheet.getDataRange().getValues();
  var parentId = "";
  var rowIdx = -1;

  for (var i = 1; i < data.length; i++) {
    if (data[i][0] == fileId) { 
      parentId = data[i][2]; 
      rowIdx = i + 1; 
      break; 
    }
  }

  if (parentId) {
    var file = DriveApp.getFileById(fileId);
    var targetFolder = DriveApp.getFolderById(parentId);
    
    // ย้ายกลับที่เดิม
    targetFolder.addFile(file);
    
    // ลบออกจากโฟลเดอร์ถังขยะ
    var recycleBin = getOrCreateFolder(DriveApp.getFolderById(DRIVE_FOLDER_ID), "RecycleBin");
    try { recycleBin.removeFile(file); } catch(e) {}

    // ลบ Log ออกจาก Sheet
    logSheet.deleteRow(rowIdx);
    
    // ส่งลิงก์ UC กลับไปให้หน้าเว็บใช้งานต่อ
    return 'https://drive.google.com/uc?export=view&id=' + fileId;
  }
  throw new Error("ไม่พบประวัติไฟล์นี้ในถังขยะ (อาจเกิน 10 วันและถูกลบถาวรไปแล้ว)");
}

// 6. ลบถาวรเมื่อเกิน 10 วัน (ตั้ง Trigger รันทุกวัน)
function autoCleanupRecycleBin() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var logSheet = ss.getSheetByName("RecycleLog");
  if(!logSheet) return;
  var data = logSheet.getDataRange().getValues();
  var now = new Date();
  for(var i = data.length - 1; i >= 1; i--) {
    var diff = (now - new Date(data[i][3])) / (1000*60*60*24);
    if(diff > 10) {
      try { DriveApp.getFileById(data[i][0]).setTrashed(true); } catch(e) {}
      logSheet.deleteRow(i + 1);
    }
  }
}

function setupPermissions() {
  DriveApp.getRootFolder();
  Logger.log("ได้รับสิทธิ์เข้าถึง Google Drive เรียบร้อยแล้ว");
}

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
   ส่วนที่ 5: ระบบช่วยเหลือ (Utils) & Logging System
   =========================================
*/

// ฟังก์ชัน Log เก่า (เก็บไว้เพื่อความเข้ากันได้)
function writelog(user, action, targetId, details) {
    // แปลงให้ไปเรียก writeAdminLog แบบง่ายๆ
    writeAdminLog(user, "LEGACY", "SYSTEM", action, targetId, details, "", "", "");
}

/**
 * ฟังก์ชันใหม่: บันทึก Log ของ Admin โดยละเอียด
 */
function writeAdminLog(user, role, group, type, targetId, details, oldVal, newVal, meta) {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName("Logs") || ss.insertSheet("Logs");
    
    // Create header if empty
    if (sheet.getLastRow() == 0) {
      sheet.appendRow(["Timestamp", "User", "Role", "ActionGroup", "ActionType", "TargetID", "Details", "OldValue", "NewValue", "Metadata"]);
      sheet.getRange(1, 1, 1, 10).setFontWeight("bold").setBackground("#f3f3f3");
    }

    var oldStr = (typeof oldVal === 'object') ? JSON.stringify(oldVal) : String(oldVal || "");
    var newStr = (typeof newVal === 'object') ? JSON.stringify(newVal) : String(newVal || "");

    sheet.appendRow([
      new Date(), 
      user, 
      role, 
      group, 
      type, 
      targetId, 
      details, 
      oldStr, 
      newStr, 
      meta || ""
    ]);
  } catch (e) { console.error("Admin Log Error: " + e.message); }
}

/**
 * ฟังก์ชันใหม่: บันทึกกิจกรรมผู้ใช้ (User Activity)
 */
function writeUserActivity(data) {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName("UserActivity") || ss.insertSheet("UserActivity");
    
    // Create header if empty
    if (sheet.getLastRow() == 0) {
      sheet.appendRow(["Timestamp", "SessionID", "Action", "TargetID", "Result", "TimeSpent", "Metadata"]);
      sheet.getRange(1, 1, 1, 7).setFontWeight("bold").setBackground("#e6f7ff");
    }

    sheet.appendRow([
      new Date(),
      data.session || "N/A",
      data.action || "",
      data.target || "",
      data.result || "",
      data.timeSpent || 0,
      data.metadata || ""
    ]);
  } catch (e) { console.error("User Log Error: " + e.message); }
}


function processVotes() {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var voteSheet = ss.getSheetByName("Votes");
    var qSheet = ss.getSheetByName("Questions");
    if (!voteSheet || !qSheet) return;

    var voteValues = voteSheet.getDataRange().getValues();
    var qValues = qSheet.getDataRange().getValues();
    var qIdMap = {};
    for (var i = 1; i < qValues.length; i++) {
        qIdMap[qValues[i][0]] = i + 1;
    }

    var hasChanged = false; 

    for (var i = 1; i < voteValues.length; i++) {
        var voteCount = parseInt(voteValues[i][3]) || 0;
        var qId = voteValues[i][0];
        var categoryToAdd = voteValues[i][2];
        var status = String(voteValues[i][5]).trim();
        var currentRow = i + 1;

        // 1. ถ้าโหวตถึงเกณฑ์ Confirm (เช่น 2 คนขึ้นไป) -> Verified (เขียวเข้ม)
        if (voteCount >= VOTE_THRESHOLD_CONFIRM && status !== "Verified") {
            updateQuestionCategory(qSheet, qIdMap, qId, categoryToAdd);
            voteSheet.getRange(currentRow, 6).setValue("Verified");
            voteSheet.getRange(currentRow, 1, 1, 6).setBackground("#6aa84f"); 
            hasChanged = true;
        }
        // 2. ปรับใหม่: ถ้ามีคะแนนตั้งแต่ 1 ขึ้นไป และยังเป็น Pending -> Approved (สถานะรอเกณฑ์)
        //    หมายเหตุ: "ไม่" apply category ที่ 1 โหวตอีกต่อไป — ต้องถึง VOTE_THRESHOLD_CONFIRM (Verified) เท่านั้น
        //    ปิดบั๊ก "1 โหวต hijack หมวดถาวร" + หยุด updateVersion() ล้าง cache ทุกโหวต
        else if (voteCount >= 1 && (status === "Pending" || status === "")) {
            voteSheet.getRange(currentRow, 6).setValue("Approved");
            voteSheet.getRange(currentRow, 1, 1, 6).setBackground(null); // ล้างสี (สีขาว)
        }
    }

    if (hasChanged) {
        updateVersion();
        sortCategorySheet();
    }
}

function processReports(doc) {
    var reportSheet = doc.getSheetByName("Report");
    var qSheet = doc.getSheetByName("Questions");
    if (!reportSheet || !qSheet) return;
    var rv = reportSheet.getDataRange().getValues();
    var qv = qSheet.getDataRange().getValues();
    var qIdMap = {};
    for (var i = 1; i < qv.length; i++) qIdMap[qv[i][0]] = i + 1;

    var changed = false;
    for (var i = 1; i < rv.length; i++) {
        if (String(rv[i][9]).trim() !== "Pending") continue;
        var voteCount = parseInt(rv[i][13]) || 0;
        if (voteCount < REPORT_VOTE_THRESHOLD) continue;

        var qId = rv[i][2];
        var suggestedAns = String(rv[i][6] || "").trim();
        var suggestedExplain = String(rv[i][12] || "").trim();
        var qRowIndex = qIdMap[qId];
        if (!qRowIndex) continue;

        // Safety: suggestedChoice must exactly match an existing choice (no free-text)
        var choicesArray = String(qv[qRowIndex-1][3] || "").split("///").map(function(s){return s.trim();}).filter(Boolean);
        if (choicesArray.indexOf(suggestedAns) === -1) continue;

        var questionText = String(qv[qRowIndex-1][1] || "");
        applyReportCorrection(qSheet, qRowIndex, suggestedAns, suggestedExplain, questionText, choicesArray);

        reportSheet.getRange(i+1, 10).setValue("AutoResolved");
        reportSheet.getRange(i+1, 11).setValue("Auto-applied by community vote (" + voteCount + "/" + REPORT_VOTE_THRESHOLD + ")");
        changed = true;
    }
    if (changed) updateVersion();
}

function applyReportCorrection(qSheet, qRowIndex, newAnswer, suggestedExplain, questionText, choicesArray) {
    qSheet.getRange(qRowIndex, 5).setValue(newAnswer);
    var newExplain = suggestedExplain || "";
    try {
        var apiKeyInfo = getAvailableAIKey("Gemini");
        if (apiKeyInfo) {
            var prompt = buildExplainPrompt(questionText, choicesArray, newAnswer);
            var aiText = callGeminiAI(prompt, apiKeyInfo, null);
            newExplain = aiText.replace(/\r?\n/g, " ").trim();
        }
    } catch(e) {
        console.warn("Gemini explain failed: " + e.message);
    }
    qSheet.getRange(qRowIndex, 6).setValue(newExplain);
}

function buildExplainPrompt(questionText, choicesArray, correctAnswer) {
    var choicesText = choicesArray.map(function(c, i) {
        var display = (c.startsWith('http') || c.startsWith('<svg')) ? '[รูปภาพ]' : c;
        return String.fromCharCode(65 + i) + ". " + display;
    }).join("\n");
    return "คุณเป็นอาจารย์แพทย์ผู้เชี่ยวชาญ กรุณาเขียนคำอธิบายเฉลยข้อสอบแพทย์ต่อไปนี้เป็น paragraph เดียวต่อเนื่อง " +
        "(ห้ามใช้ bullet points หรือขึ้นบรรทัดใหม่) โดยใช้ภาษาไทยผสมคำศัพท์ทางการแพทย์ภาษาอังกฤษ ห้ามใช้ภาษาอังกฤษล้วน\n\n" +
        "โจทย์: " + questionText + "\n\n" +
        "ตัวเลือก:\n" + choicesText + "\n\n" +
        "เฉลยที่ถูกต้อง: " + correctAnswer + "\n\n" +
        "คำอธิบายต้องครอบคลุม: 1) Key concept/การวินิจฉัย 2) เหตุผลที่เฉลยถูก พร้อมชี้ clues จากโจทย์ " +
        "3) อธิบายว่าทำไมตัวเลือกที่ผิดแต่ละข้อถึงผิด 4) Clinical pearl ถ้ามี\n\n" +
        "เขียนเป็น paragraph เดียว ห้ามมี newline ในคำตอบ:";
}

function updateQuestionCategory(qSheet, qIdMap, qId, categoryToAdd) {
    var qRowIndex = qIdMap[qId];
    if (qRowIndex) {
        var catCell = qSheet.getRange(qRowIndex, 7);
        var currentCats = [];
        try {
            var val = catCell.getValue().toString();
            currentCats = (val && val !== "") ? JSON.parse(val.replace(/'/g, '"')) : [];
        } catch (e) {
            currentCats = [catCell.getValue().toString()];
        }

        if (currentCats.indexOf(categoryToAdd) === -1) {
            currentCats.push(categoryToAdd);
            catCell.setValue(JSON.stringify(currentCats));
            autoCreateSplitCategories(qId, currentCats);
        }
    }

}

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

function getPendingReportCount(filterSubject) {
  var v = getVotesVersionCached();
  var cleanFilter = filterSubject ? String(filterSubject).trim().toUpperCase() : "all";
  var cacheKey = "pending_report_count_" + v + "_" + cleanFilter;
  var cached = getLargeCache(cacheKey);
  if (cached != null) {
    return ContentService.createTextOutput(cached).setMimeType(ContentService.MimeType.JSON);
  }

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
  putLargeCache(cacheKey, responseStr, 300); // 5 minutes cache
  return ContentService.createTextOutput(responseStr).setMimeType(ContentService.MimeType.JSON);
}

function autoCreateSplitCategories(questionId, categories, skipSort) {
  if (!categories || categories.length < 2) return;

  // 1. กรอง "by AI" ออก และตรวจสอบว่ามาจาก Subject เดียวกันหรือไม่
  const validCats = categories.filter(c => !c.toLowerCase().includes("by ai"));
  if (validCats.length < 2) return;

  const firstCatId = validCats[0];
  const subjectId = firstCatId.split('_')[0]; // เช่น GI

  // ตรวจสอบว่าทุกอันขึ้นต้นด้วย Subject เดียวกัน
  const sameSubject = validCats.every(c => c.startsWith(subjectId));
  if (!sameSubject) return;

  // 2. ระบุ Source (ชุดข้อสอบเก่า) และ Lecture (หัวข้อเรียน)
  // สมมติ: อันแรกคือชุดข้อสอบ (GI_51MCQ1...), อันที่สองคือหัวข้อ (GI_ANA_...)
  const sourceCatId = validCats[0];
  const lectureCatId = validCats[1];

  // 3. Mapping หมวดหมู่ 6 กลุ่ม
  let splitSuffix = "";
  let groupKey = "";

  const upperLect = lectureCatId.toUpperCase();

  if (upperLect.includes("_ANA_")) { groupKey = "ANA"; splitSuffix = "ANATOMY (Extracted)"; }
  else if (upperLect.includes("_PHY_") || upperLect.includes("_PHYSIO_") || upperLect.includes("_BIOCHEM_")) { groupKey = "PHYSIO and BIOCHEM"; splitSuffix = "PHYSIO and BIOCHEM (Extracted)"; }
  else if (upperLect.includes("_PARASITO_") || upperLect.includes("_MICRO_")) { groupKey = "PARASITO and MICRO"; splitSuffix = "PARASITO and MICRO (Extracted)"; }
  else if (upperLect.includes("_PATHO_")) { groupKey = "PATHO"; splitSuffix = "PATHO (Extracted)"; }
  else if (upperLect.includes("_PHARM_") || upperLect.includes("_PHARMACO_")) { groupKey = "PHARM"; splitSuffix = "PHARM (Extracted)"; }
  else if (upperLect.includes("_IMAGE_") || upperLect.includes("_RADIO_") || upperLect.includes("_CLINICAL_")) { groupKey = "RADIO and CLINICAL"; splitSuffix = "RADIO and CLINICAL (Extracted)"; }

  if (!groupKey) return; // ถ้าไม่ตรงกับ 6 กลุ่มที่กำหนด ไม่ต้องทำต่อ

  // 4. สร้าง ID และชื่อใหม่
  // ตัวอย่าง: GI_51MCQ1_ANA_Extracted
  const newSplitCatId = `${sourceCatId}_${groupKey.replace(/\s+/g, '')}_Extracted`;
  const newSplitCatName = `${sourceCatId} (${splitSuffix})`;
  const newAccordionGroup = `${subjectId} (Extracted)`;

  const ss = SpreadsheetApp.openById(SHEET_ID);
  const catSheet = ss.getSheetByName("Category");
  const structSheet = ss.getSheetByName("Structure");

  // 5. ตรวจสอบและเพิ่มลงในแผ่นงาน Category
  const catValues = catSheet.getDataRange().getValues();
  let catExists = false;
  for (let i = 1; i < catValues.length; i++) {
    if (catValues[i][0] === newSplitCatId) { catExists = true; break; }
  }

  if (!catExists) {
    catSheet.appendRow([newSplitCatId, subjectId, newAccordionGroup, newSplitCatName]);
  }

  // 6. ตรวจสอบและเพิ่มลงในแผ่นงาน Structure
  const structValues = structSheet.getDataRange().getValues();
  let structExists = false;
  for (let i = 1; i < structValues.length; i++) {
    if (structValues[i][1] === subjectId && structValues[i][3] === newAccordionGroup) {
      structExists = true;
      break;
    }
  }

  if (!structExists) {
    // ดึง Year จากอันเดิมมาใส่ (ถ้าหาเจอ)
    let year = "0";
    for (let i = 1; i < structValues.length; i++) {
      if (structValues[i][1] === subjectId) { year = structValues[i][0]; break; }
    }
    structSheet.appendRow([year, subjectId, subjectId, newAccordionGroup]);
  }

  // 7. เพิ่ม NewSplitCatId เข้าไปในคำถามนั้น (ถ้ายังไม่มี)
  const qSheet = ss.getSheetByName("Questions");
  const qData = qSheet.getDataRange().getValues();
  for (let i = 1; i < qData.length; i++) {
    if (qData[i][0] === questionId) {
      let currentCats = [];
      try {
        currentCats = JSON.parse(qData[i][6].replace(/'/g, '"'));
      } catch (e) { currentCats = [qData[i][6]]; }

      if (!currentCats.includes(newSplitCatId)) {
        currentCats.push(newSplitCatId);
        qSheet.getRange(i + 1, 7).setValue(JSON.stringify(currentCats));
      }
      break;
    }
  }

  // บังคับข้ามการจัดเรียงหากทำงานอยู่ภายใต้คำสั่งประมวลผลเป็นกลุ่ม (Deferred Sorting)
  if (!skipSort) {
    sortCategorySheet();
  }
}

/**
 * ฟังก์ชันรันตรวจสอบคัดแยกวิชากลุ่มย่อย (Extracted) สำหรับคำถามทั้งหมดในแผ่นงาน Questions แบบ Manual
 */
function runManualSplitExtraction() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var ui = SpreadsheetApp.getUi();

  // ============================================================
  // STEP 1: โหลดข้อมูลทั้งหมดเข้า Memory ครั้งเดียว
  // ============================================================
  var qSheet   = ss.getSheetByName("Questions");
  var catSheet  = ss.getSheetByName("Category");
  var structSheet = ss.getSheetByName("Structure");

  if (!qSheet || !catSheet || !structSheet) {
    ui.alert("❌ ไม่พบ Sheet ที่จำเป็น");
    return;
  }

  var qData     = qSheet.getDataRange().getValues();
  var catData   = catSheet.getDataRange().getValues();
  var structData = structSheet.getDataRange().getValues();

  if (qData.length < 2) {
    ui.alert("ℹ️ ไม่พบข้อมูลคำถาม");
    return;
  }

  // ============================================================
  // STEP 2: สร้าง Lookup Maps จาก Memory (ไม่ต้องเปิด Sheet อีก)
  // ============================================================

  // Map: categoryId -> subjectId
  var catToSubject = {};
  // Set: subjectId|accordionGroup ที่มีอยู่ใน Structure แล้ว
  var existingStructKeys = {};
  // Set: categoryId ที่มีอยู่ใน Category แล้ว
  var existingCatIds = {};
  // Map: subjectId -> year (สำหรับ Structure row ใหม่)
  var subjectToYear = {};

  for (var i = 1; i < catData.length; i++) {
    var cid = String(catData[i][0]).trim();
    existingCatIds[cid] = true;
    catToSubject[cid] = String(catData[i][1]).trim();
  }
  for (var i = 1; i < structData.length; i++) {
    var sid = String(structData[i][1]).trim();
    var grp = String(structData[i][3]).trim();
    existingStructKeys[sid + "|" + grp] = true;
    if (!subjectToYear[sid]) subjectToYear[sid] = structData[i][0];
  }

  // ============================================================
  // STEP 3: คำนวณทุกอย่างใน Memory — ไม่แตะ Sheet เลยใน loop
  // ============================================================

  // ผลลัพธ์ที่จะเขียนทีเดียวตอนท้าย
  var newCatRows    = [];   // แถวใหม่สำหรับ Category Sheet
  var newStructRows = [];   // แถวใหม่สำหรับ Structure Sheet
  // Map: questionRowIndex -> categories array ที่อัปเดตแล้ว
  var qUpdates = {};

  var processCount  = 0;
  var skippedCount  = 0;

  var groupKeyMap = {
    "_ANA_":       { key: "ANA",             suffix: "ANATOMY (Extracted)" },
    "_PHY_":       { key: "PHYSIOandBIOCHEM", suffix: "PHYSIO and BIOCHEM (Extracted)" },
    "_PHYSIO_":    { key: "PHYSIOandBIOCHEM", suffix: "PHYSIO and BIOCHEM (Extracted)" },
    "_BIOCHEM_":   { key: "PHYSIOandBIOCHEM", suffix: "PHYSIO and BIOCHEM (Extracted)" },
    "_PARASITO_":  { key: "PARASITOandMICRO", suffix: "PARASITO and MICRO (Extracted)" },
    "_MICRO_":     { key: "PARASITOandMICRO", suffix: "PARASITO and MICRO (Extracted)" },
    "_PATHO_":     { key: "PATHO",            suffix: "PATHO (Extracted)" },
    "_PHARM_":     { key: "PHARM",            suffix: "PHARM (Extracted)" },
    "_PHARMACO_":  { key: "PHARM",            suffix: "PHARM (Extracted)" },
    "_IMAGE_":     { key: "RADIOandCLINICAL", suffix: "RADIO and CLINICAL (Extracted)" },
    "_RADIO_":     { key: "RADIOandCLINICAL", suffix: "RADIO and CLINICAL (Extracted)" },
    "_CLINICAL_":  { key: "RADIOandCLINICAL", suffix: "RADIO and CLINICAL (Extracted)" },
  };

  for (var i = 1; i < qData.length; i++) {
    var qId    = String(qData[i][0]).trim();
    var catRaw = String(qData[i][6]).trim();
    if (!qId || !catRaw) continue;

    // Parse categories
    var categories = [];
    try {
      categories = (catRaw.indexOf("[") > -1)
        ? JSON.parse(catRaw.replace(/'/g, '"'))
        : [catRaw];
    } catch (e) { continue; }

    // Guard: ต้องมี >= 2 categories
    if (categories.length < 2) { skippedCount++; continue; }

    // กรอง "by ai" และ "_Extracted" ออก
    var validCats = categories.filter(function(c) {
      var cu = c.toLowerCase();
      return !cu.includes("by ai") && !c.endsWith("_Extracted");
    });

    if (validCats.length < 2) { skippedCount++; continue; }

    var sourceCatId  = validCats[0];
    var lectureCatId = validCats[1];
    var subjectId    = sourceCatId.split('_')[0];

    // ตรวจว่าทุกอัน startsWith subjectId เดียวกัน
    var sameSubject = validCats.every(function(c) { return c.startsWith(subjectId); });
    if (!sameSubject) { skippedCount++; continue; }

    // หา groupKey จาก lectureCatId
    var upperLect = lectureCatId.toUpperCase();
    var matched = null;
    var keys = Object.keys(groupKeyMap);
    for (var k = 0; k < keys.length; k++) {
      if (upperLect.indexOf(keys[k]) > -1) { matched = groupKeyMap[keys[k]]; break; }
    }
    if (!matched) { skippedCount++; continue; }

    var newCatId       = sourceCatId + "_" + matched.key + "_Extracted";
    var newCatName     = sourceCatId + " (" + matched.suffix + ")";
    var newAccordion   = subjectId + " (Extracted)";

    // Guard: ข้ามถ้า extract แล้วและ question มี newCatId อยู่แล้ว
    if (existingCatIds[newCatId] && categories.indexOf(newCatId) > -1) {
      skippedCount++;
      continue;
    }

    // --- เพิ่ม Category ถ้ายังไม่มี (บันทึกใน memory) ---
    if (!existingCatIds[newCatId]) {
      newCatRows.push([newCatId, subjectId, newAccordion, newCatName]);
      existingCatIds[newCatId] = true;  // อัปเดต Map ใน memory ด้วย
      catToSubject[newCatId] = subjectId;
    }

    // --- เพิ่ม Structure ถ้ายังไม่มี (บันทึกใน memory) ---
    var structKey = subjectId + "|" + newAccordion;
    if (!existingStructKeys[structKey]) {
      var year = subjectToYear[subjectId] || "0";
      newStructRows.push([year, subjectId, subjectId, newAccordion]);
      existingStructKeys[structKey] = true;
    }

    // --- อัปเดต categories ของคำถาม (บันทึกใน memory) ---
    if (categories.indexOf(newCatId) === -1) {
      var updatedCats = categories.concat([newCatId]);
      qUpdates[i] = updatedCats;  // i = row index ใน qData array
    }

    processCount++;
  }

  // ============================================================
  // STEP 4: เขียนทุกอย่างลง Sheet ครั้งเดียว (Batch Write)
  // ============================================================

  if (newCatRows.length > 0) {
    var catLastRow = catSheet.getLastRow();
    catSheet.getRange(catLastRow + 1, 1, newCatRows.length, 4).setValues(newCatRows);
  }

  if (newStructRows.length > 0) {
    var structLastRow = structSheet.getLastRow();
    structSheet.getRange(structLastRow + 1, 1, newStructRows.length, 4).setValues(newStructRows);
  }

  // อัปเดต Questions Sheet: เขียนเฉพาะ row ที่เปลี่ยนแปลง
  var qRowIndices = Object.keys(qUpdates);
  for (var j = 0; j < qRowIndices.length; j++) {
    var rowIdx = parseInt(qRowIndices[j]);
    var sheetRow = rowIdx + 1; // +1 เพราะ getValues() เริ่มที่ index 0 = header
    qSheet.getRange(sheetRow, 7).setValue(JSON.stringify(qUpdates[rowIdx]));
  }

  // เรียก sort และ version update แค่ครั้งเดียวตอนท้าย
  if (processCount > 0 || newCatRows.length > 0) {
    updateVersion();
    sortCategorySheet();
  }

  ui.alert(
    "✅ เสร็จสิ้น\n" +
    "• ประมวลผล: " + processCount + " ข้อ\n" +
    "• Category ใหม่: " + newCatRows.length + " รายการ\n" +
    "• Structure ใหม่: " + newStructRows.length + " รายการ\n" +
    "• ข้าม: " + skippedCount + " ข้อ"
  );
}

function sortCategorySheet(ss) {
  if (!ss) ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName("Category");
  if (!sheet) return;

  // บังคับล้างลบแคชของ Category และตารางจัดกลุ่มความสัมพันธ์ในทันทีก่อนเรียงลำดับใหม่
  var cache = CacheService.getScriptCache();
  var v = getVersionCached();
  cache.remove("category_sheet_raw_" + v);
  cache.remove("cat_to_subj_map_" + v);

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  var range = sheet.getRange(2, 1, lastRow - 1, 4);
  var data = range.getValues();

  var extractYear = function (id) {
    var match = String(id).match(/\d+/);
    return match ? parseInt(match[0]) : 0;
  };

  var getSubSubjectPriority = function (group, id, name) {
    var text = (String(group) + " " + String(id) + " " + String(name)).toUpperCase();
    if (text.includes("_ANA_")) return 1;
    if (text.includes("_PHYSIO") || text.includes("BIOCHEM_")) return 2;
    if (text.includes("_MICRO") || text.includes("PARASITO_")) return 3;
    if (text.includes("_PATHO_")) return 4;
    if (text.includes("_PHARM_")) return 5;
    if (text.includes("_RADIO_") || text.includes("_CLINIC_")) return 6;
    return 7;
  };

  var getNumberSuffix = function (group, id, keyword) {
    var text = (String(group) + " " + String(id)).toUpperCase();
    var regex = new RegExp(keyword.toUpperCase() + "(\\d+)");
    var match = text.match(regex);
    if (match) return parseInt(match[1]);
    return 0;
  };

  var getGroupPriority = function (group, id) {
    var g = String(group).toUpperCase();
    var i = String(id).toUpperCase();

    if (g.includes("FMT")) return 10;
    if (g.includes("EXTRACTED") || i.includes("EXTRACTED")) return 30;
    if (g.includes("MCQ") || i.includes("MCQ")) return 20;
    if (g.includes("BY AI")) return 50;
    if (g.includes("LEC")) return 40;

    return 99;
  };

  data.sort(function (a, b) {
    var subA = String(a[1]);
    var subB = String(b[1]);
    if (subA !== subB) return subA.localeCompare(subB);

    var prioA = getGroupPriority(a[2], a[0]);
    var prioB = getGroupPriority(b[2], b[0]);

    if (prioA !== prioB) return prioA - prioB;

    if (prioA === 30 || prioA === 40 || prioA === 50) {
      var sRankA = getSubSubjectPriority(a[2], a[0], a[3]);
      var sRankB = getSubSubjectPriority(b[2], b[0], b[3]);
      if (sRankA !== sRankB) return sRankA - sRankB;
    }

    var yearA = extractYear(a[0]);
    var yearB = extractYear(b[0]);
    if (yearA !== yearB) return yearB - yearA;

    if (prioA === 10 || prioA === 20) {
      var keyword = (prioA === 10) ? "FMT" : "MCQ";
      var nA = getNumberSuffix(a[2], a[0], keyword);
      var nB = getNumberSuffix(b[2], b[0], keyword);
      if (nA !== nB) return nA - nB;
    }

    return 0;
  });

  // บังคับ Flush ข้อมูลลงชีตหลักให้เรียบร้อยก่อนเขียนทับ เพื่อความปลอดภัยของข้อมูล
  SpreadsheetApp.flush();
  range.setValues(data);
  updateVersion();
}
function verifyAllSingleCategoryVotes() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var voteSheet = ss.getSheetByName("Votes");
  var qSheet = ss.getSheetByName("Questions");
  if (!voteSheet || !qSheet) {
    SpreadsheetApp.getUi().alert("❌ ไม่พบแผ่นงาน Votes หรือ Questions");
    return;
  }

  var voteValues = voteSheet.getDataRange().getValues();
  var qValues = qSheet.getDataRange().getValues();
  
  // 1. สร้าง Map ของคำถามเพื่อความรวดเร็วในการตรวจสอบ
  // qMap[questionId] = { row: rowIndex, categories: [cat1, cat2] }
  var qMap = {};
  for (var i = 1; i < qValues.length; i++) {
    var qId = qValues[i][0];
    var catRaw = String(qValues[i][6] || "");
    var categories = [];
    
    try {
      if (catRaw !== "") {
        categories = (catRaw.indexOf("[") > -1) ? JSON.parse(catRaw.replace(/'/g, '"')) : [catRaw];
      }
    } catch (e) {
      categories = [catRaw];
    }
    qMap[qId] = { 
      row: i + 1, 
      categories: categories 
    };
  }

  var verifiedCount = 0;
  var qIdMapForUpdate = {}; // ใช้สำหรับฟังก์ชัน updateQuestionCategory เดิม
  for (var i = 1; i < qValues.length; i++) {
    qIdMapForUpdate[qValues[i][0]] = i + 1;
  }

  // 2. ไล่ดูรายการโหวต
  for (var j = 1; j < voteValues.length; j++) {
    var qId = voteValues[j][0];
    var categoryIdInVote = voteValues[j][2];
    var status = String(voteValues[j][5]).trim();
    
    // ข้ามถ้าเป็น Verified ไปแล้ว
    if (status === "Verified") continue;

    var qInfo = qMap[qId];
    if (qInfo) {
      // เงื่อนไข: มีแค่ 1 category
      if (qInfo.categories.length === 1) {
        var currentCat = qInfo.categories[0].toLowerCase();
        
        // เงื่อนไข: ต้องไม่ใช่ AI
        if (currentCat.indexOf("by ai") === -1) {
          
          // ทำการอัปเดตข้อมูล (เรียกใช้ฟังก์ชันที่มีอยู่แล้ว)
          updateQuestionCategory(qSheet, qIdMapForUpdate, qId, categoryIdInVote);
          
          // อัปเดตสถานะในหน้าโหวต
          voteSheet.getRange(j + 1, 6).setValue("Verified");
          voteSheet.getRange(j + 1, 1, 1, 6).setBackground("#6aa84f"); // สีเขียวเข้ม
          
          verifiedCount++;
        }
      }
    }
  }

  // 3. สรุปผล
  if (verifiedCount > 0) {
    updateVersion();
    sortCategorySheet();
    SpreadsheetApp.getUi().alert('✅ ดำเนินการ Verified เรียบร้อยแล้ว ' + verifiedCount + ' รายการ');
  } else {
    SpreadsheetApp.getUi().alert('ℹ️ ไม่พบคำถามที่ตรงตามเงื่อนไข (1 Category & No AI)');
  }
}

/**
 * สคริปต์ตรวจสอบรูปภาพในระบบเปรียบเทียบกับไฟล์ใน Google Drive
 * สร้างโดยอ้างอิง Root Folder ID: 1nzLH2ia2lL2TMxfrr6Kv-5fhsWwOWSCm
 */

function generateImageVerificationReport() {
  var TARGET_DRIVE_ID = '1nzLH2ia2lL2TMxfrr6Kv-5fhsWwOWSCm'; // ลิงก์ที่คุณให้มา
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var qSheet = ss.getSheetByName("Questions");
  
  // 1. สร้าง/เตรียมแผ่นงาน Report
  var reportSheet = ss.getSheetByName("Image_Migration_Report");
  if (reportSheet) {
    reportSheet.clear();
  } else {
    reportSheet = ss.insertSheet("Image_Migration_Report");
  }
  
  // เขียน Header ของ Report
  reportSheet.appendRow(["QuestionID", "Source Column", "Original URL", "Extracted ID", "Status", "Filename in Drive"]);
  reportSheet.getRange(1, 1, 1, 6).setFontWeight("bold").setBackground("#d9ead3");

  // 2. Brute Force สแกนหาไฟล์ทั้งหมดใน Drive (รวม Folder ย่อย)
  console.log("Starting Drive Scan... might take a while.");
  var driveMap = {}; // เก็บ {id: filename}
  var rootFolder = DriveApp.getFolderById(TARGET_DRIVE_ID);
  recursiveMapFolder(rootFolder, driveMap);
  console.log("Drive Scan Complete. Found " + Object.keys(driveMap).length + " files.");

  // 3. ดึงข้อมูลจาก Questions Sheet
  var data = qSheet.getDataRange().getValues();
  var reportRows = [];

  // เริ่มวนลูปข้อมูล (ข้าม Header แถวที่ 0)
  for (var i = 1; i < data.length; i++) {
    var qid = data[i][0];
    var imgCol = data[i][2];     // คอลัมน์ Image (โจทย์)
    var choicesCol = data[i][3]; // คอลัมน์ Choices (เผื่อมีรูปในตัวเลือก)

    // ตรวจสอบคอลัมน์รูปโจทย์
    processVerification(qid, "Image (Main)", imgCol, driveMap, reportRows);
    
    // ตรวจสอบคอลัมน์ตัวเลือก
    processVerification(qid, "Choices", choicesCol, driveMap, reportRows);
  }

  // 4. บันทึกผลลัพธ์ลงใน Sheet Report
  if (reportRows.length > 0) {
    reportSheet.getRange(2, 1, reportRows.length, 6).setValues(reportRows);
        console.log("ตรวจสอบเสร็จสิ้น! พบลิงก์รูปภาพทั้งหมด " + reportRows.length + " รายการ ใน Sheet 'Image_Migration_Report'");
return "Success: Found " + reportRows.length + " images.";
  } else {
    SpreadsheetApp.getUi().alert("ไม่พบลิงก์รูปภาพในฐานข้อมูล");
    return "No images found.";
  }
}

/**
 * ฟังก์ชันช่วยวนลูปหาไฟล์ในโฟลเดอร์ย่อยทั้งหมด (Recursive)
 */
function recursiveMapFolder(folder, driveMap) {
  var files = folder.getFiles();
  while (files.hasNext()) {
    var file = files.next();
    driveMap[file.getId()] = file.getName();
  }
  
  var subFolders = folder.getFolders();
  while (subFolders.hasNext()) {
    recursiveMapFolder(subFolders.next(), driveMap);
  }
}

/**
 * ฟังก์ชันสกัด ID และตรวจสอบสถานะ
 */
function processVerification(qid, colName, cellValue, driveMap, reportRows) {
  if (!cellValue || cellValue == "") return;
  
  // แยกส่วนด้วย /// กรณีมีหลายรูป
  var parts = String(cellValue).split("///");
  
  parts.forEach(function(part) {
    var trimmedPart = part.trim();
    if (trimmedPart.includes("drive.google.com") || trimmedPart.includes("id=")) {
      var fileId = extractId(trimmedPart);
      var status = "Not Found";
      var fileName = "-";
      
      if (driveMap[fileId]) {
        status = "Found";
        fileName = driveMap[fileId];
      }
      
      reportRows.push([qid, colName, trimmedPart, fileId, status, fileName]);
    }
  });
}

/**
 * ฟังก์ชัน Regex สกัด ID จาก URL หลากหลายรูปแบบ
 */
function extractId(url) {
  var match = url.match(/\/d\/(.*?)\//) || 
              url.match(/id=([^&]+)/) || 
              url.match(/\/d\/([^\/\?]+)/);
  return (match && match[1]) ? match[1] : url;
}

/**
 * สคริปต์จัดระเบียบรูปภาพวิชา GEN5
 * หากไม่ได้เป็นเจ้าของไฟล์ จะย้ายไปที่โฟลเดอร์ MD > Unknown
 */
function migrateGEN5WithFullStructure() {
  var ROOT_FOLDER_ID = '1nzLH2ia2lL2TMxfrr6Kv-5fhsWwOWSCm'; 
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var qSheet = ss.getSheetByName("Questions");
  var data = qSheet.getDataRange().getValues();
  var myEmail = Session.getEffectiveUser().getEmail(); // อีเมลผู้รันสคริปต์
  
  try {
    var rootFolder = DriveApp.getFolderById(ROOT_FOLDER_ID);
    console.log("--- Starting GEN5 Migration (Active User: " + myEmail + ") ---");

    for (var i = 1; i < data.length; i++) {
      var qid = String(data[i][0]); 
      var imgCell = String(data[i][2]); 
      var choicesCell = String(data[i][3]); 

      // 1. กรองเฉพาะวิชา GEN5
      if (!qid || !qid.startsWith("MS")) continue;
      
      var hasMainImg = (imgCell && imgCell != "" && !imgCell.toLowerCase().includes("require_img"));
      var hasChoiceImg = (choicesCell && choicesCell.includes("drive.google.com"));
      
      if (!hasMainImg && !hasChoiceImg) continue; 

      // 2. วิเคราะห์ Path
      var parts = qid.split('_');
      if (parts.length < 2) continue;
      
      var yearType = parts[1].trim();
      var pathNames = ["MD", "Y2", "MS"]; 
      
      if (/^\d{2}/.test(yearType)) {
        pathNames.push(yearType.substring(0, 2)); // Year
        pathNames.push(yearType.substring(2) || "General"); // Type
      } else {
        pathNames.push(yearType); 
      }

      // 3. เตรียมโฟลเดอร์หลักและโฟลเดอร์ Unknown
      var mdFolder = getOrCreateSubFolder(rootFolder, "MD");
      var unknownFolder = getOrCreateSubFolder(mdFolder, "Unknown");

      // เตรียมโฟลเดอร์ปลายทางตามโครงสร้างปกติ
      var targetFolder = rootFolder;
      pathNames.forEach(function(name) {
        targetFolder = getOrCreateSubFolder(targetFolder, name);
      });

      // 4. จัดการรูปโจทย์
      if (hasMainImg) {
        var imgUrls = imgCell.split("///");
        imgUrls.forEach(function(url, idx) {
          var newName = "Q_" + qid + "_Main_" + (idx + 1);
          moveAndRenameFile(url, newName, targetFolder, unknownFolder, myEmail);
        });
      }

      // 5. จัดการรูปในตัวเลือก
      if (hasChoiceImg) {
        var choiceParts = choicesCell.split("///");
        choiceParts.forEach(function(content, idx) {
          if (content.includes("drive.google.com") || content.includes("id=")) {
            var letter = String.fromCharCode(65 + idx); 
            var newName = "Q_" + qid + "_Choice_" + letter;
            moveAndRenameFile(content, newName, targetFolder, unknownFolder, myEmail);
          }
        });
      }
    }
    console.log("--- Finished GEN5 Migration ---");
  } catch (e) {
    console.error("Critical Error: " + e.message);
  }
}

/**
 * ฟังก์ชันย้ายไฟล์และเปลี่ยนชื่อ พร้อมตรวจสอบความเป็นเจ้าของ
 */
function moveAndRenameFile(url, newName, targetFolder, unknownFolder, myEmail) {
  try {
    var fileId = extractIdFromUrl(url);
    if (!fileId) return;

    var file = DriveApp.getFileById(fileId);
    var owner = file.getOwner() ? file.getOwner().getEmail() : "Unknown";
    
    // ตรวจสอบความเป็นเจ้าของ
    var finalDestination = targetFolder;
    if (owner !== myEmail) {
      finalDestination = unknownFolder;
      console.log("Not Owner (" + owner + "): Moving " + newName + " to Unknown folder");
    }

    var currentParents = file.getParents();
    var currentParentId = currentParents.hasNext() ? currentParents.next().getId() : "";
    
    // ถ้าชื่อตรงและที่อยู่ตรงแล้ว ให้ข้าม
    if (file.getName() === newName && currentParentId === finalDestination.getId()) {
      return; 
    }

    // พยายามเปลี่ยนชื่อ (ถ้าสิทธิ์ไม่พอจะติด Catch)
    try {
      file.setName(newName);
    } catch(e) {
      console.warn("Cannot rename (Permission): " + newName);
    }
    
    // ย้ายไฟล์ (ถ้าไม่ใช่เจ้าของแต่อยู่ในโฟลเดอร์ที่แชร์ Editor ไว้ก็อาจย้ายได้)
    if (currentParentId !== finalDestination.getId()) {
      try {
        finalDestination.addFile(file);
        DriveApp.getFolderById(currentParentId).removeFile(file);
      } catch(e) {
        // ถ้า remove ไม่ได้ (เพราะไม่ใช่เจ้าของ) แต่อย่างน้อยก็เพิ่มไฟล์เข้าไปในที่ใหม่ได้
        console.warn("Move limited: " + newName + " added to " + finalDestination.getName());
      }
    }
    
    console.log("Processed: " + newName + " (at " + finalDestination.getName() + ")");

  } catch (e) {
    console.warn("Error processing " + newName + ": " + e.message);
  }
}

function getOrCreateSubFolder(parentFolder, name) {
  var folders = parentFolder.getFoldersByName(name);
  if (folders.hasNext()) {
    return folders.next();
  } else {
    return parentFolder.createFolder(name);
  }
}

// Helper: ดึง File ID จาก URL
function extractIdFromUrl(url) {
  var match = url.match(/\/d\/(.*?)\//) || url.match(/id=([^&]+)/) || url.match(/\/d\/([^\/\?]+)/);
  return (match && match[1]) ? match[1] : null;
}

// ลบแถวใน Questions ที่มี img = "require_img" (placeholder ที่ไม่ถูก patch ก่อน import)
// ตั้ง Time Trigger รันทุกวัน 3-4 AM จาก Apps Script console
function cleanupStaging() {
  var doc = SpreadsheetApp.openById(SHEET_ID);
  var sheet = doc.getSheetByName('Questions');
  if (!sheet) return;
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  var deleted = 0;
  // วนลูปจากล่างขึ้นบนเพื่อไม่ให้ index เลื่อน
  for (var i = lastRow; i >= 2; i--) {
    var imgVal = String(sheet.getRange(i, 3).getValue()).trim();
    if (imgVal === 'require_img') {
      sheet.deleteRow(i);
      deleted++;
    }
  }
  console.log('cleanupStaging: deleted ' + deleted + ' require_img rows');
}

/*
   =========================================
   ส่วนที่ 3: KKU IntelSphere Shared Key Pool
   (Idea/interested-using-kkuintel.md — v7)
   =========================================
*/

var INTELSPHERE_SHEET_NAME = "IntelSphere_Keys";
var INTELSPHERE_ENDPOINT = "https://gen.ai.kku.ac.th/api/v1/chat/completions";
var INTELSPHERE_QUOTA_FLOOR = 0.05; // skip a provider whose remaining < 5% of its daily limit

var INTELSPHERE_LIMITS = {
  "Deepseek": 1000000, "Gemini": 350000, "Meta": 200000, "Nova": 200000, "xAI": 200000,
  "Qwen": 200000, "OpenAI": 150000, "Claude": 150000, "Mistral": 150000, "MiniMax": 100000
  // Perplexity intentionally excluded — no published model ID
};

var INTELSPHERE_PROVIDER_PRIORITY = [
  "Deepseek", "Gemini", "Meta", "Nova", "xAI", "Qwen", "OpenAI", "Claude", "Mistral", "MiniMax"
];

// One flagship model per provider — used ONLY when rotation moves to a provider
// other than the one the student explicitly requested.
var PROVIDER_MODEL_MAP = {
  "Deepseek": "deepseek-v4-pro",  "Gemini": "gemini-2.5-flash",   "Meta": "llama-4-maverick",
  "Nova":     "nova-pro-v1",       "xAI":    "grok-4",             "Qwen": "qwen3.7-plus",
  "OpenAI":   "gpt-5-mini",        "Claude": "claude-sonnet-4.5",  "Mistral": "mistral-medium-3",
  "MiniMax":  "minimax-m3"
};

// Hardcoded fallback catalog — used ONLY when the live GET /models fetch fails.
var PROVIDER_MODELS_FALLBACK = {
  "Claude":   ["claude-sonnet-5","claude-sonnet-4.6","claude-sonnet-4.5","claude-haiku-4.5","claude-sonnet-4","claude-3.7-sonnet"],
  "Deepseek": ["deepseek-v4-pro","deepseek-v4-flash","deepseek-v3.2","deepseek-v3.2-exp","deepseek-chat-v3.1"],
  "Gemini":   ["gemini-3.5-flash","gemini-3.1-pro-preview","gemini-3.1-flash-lite","gemini-3.1-flash-lite-preview","gemini-3-flash-preview","gemini-2.5-pro","gemini-2.5-flash","gemini-2.5-flash-lite","gemini-3-pro-preview"],
  "Meta":     ["llama-4-maverick","llama-4-scout"],
  "MiniMax":  ["minimax-m3"],
  "Mistral":  ["mistral-small-2603","mistral-large-2512","mistral-medium-3","codestral-2508","devstral-medium","codestral-2501"],
  "Nova":     ["nova-2-lite-v1","nova-pro-v1"],
  "OpenAI":   ["gpt-5.4","gpt-5.4-mini","gpt-5.4-nano","gpt-5.2","gpt-5.1","gpt-5.1-codex","gpt-5","gpt-5-mini","gpt-5-nano","gpt-5.5"],
  "Qwen":     ["qwen3.7-plus","qwen3.7-max","qwen3.6-flash","qwen3.5-9b","qwen3-235b-a22b-2507","qwen3-next-80b-a3b-instruct","qwen3-coder-flash","qwen3-coder","qwen3-vl-32b-instruct"],
  "xAI":      ["grok-4.3","grok-4.1-fast","grok-4","grok-3"]
};

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

// เลือก (key, provider) ที่ยังมีโควต้า — weighted-random ตาม remaining tokens
// (กับ key เดียวในระบบ พฤติกรรมเทียบเท่า top-down scan; รองรับหลาย key อัตโนมัติเมื่อมีผู้บริจาคเพิ่ม)
function getActiveIntelSphereKey(requestedProvider) {
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
function executeChatbotQuery(prompt, requestedModel, attempt) {
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

  var payload = { model: actualModel, messages: [{ role: "user", content: prompt }], max_tokens: 800, temperature: 0.3 };
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
      return executeChatbotQuery(prompt, requestedModel, attempt + 1);
    }

    if (errText.indexOf("Invalid model") >= 0) {
      // Catalog drift — bust cache แล้ว retry ด้วย flagship ของ provider เดิม (ไม่แตะ donor key)
      CacheService.getScriptCache().remove("intelsphere_catalog");
      console.warn("[IntelSphere] Invalid model at request time: " + actualModel + " — catalog cache cleared");
      if (actualModel !== PROVIDER_MODEL_MAP[keyObj.provider]) {
        return executeChatbotQuery(prompt, PROVIDER_MODEL_MAP[keyObj.provider], attempt + 1);
      }
      throw new Error("เกิดข้อผิดพลาดในการตั้งค่าโมเดล AI กรุณาแจ้งทีม IT");
    }

    // Key เสีย/ถูกเพิกถอนจริงๆ
    sheet2.getRange(keyObj.rowIndex, headers2.indexOf("Status") + 1).setValue("Invalid");
    SpreadsheetApp.flush();
    return executeChatbotQuery(prompt, requestedModel, attempt + 1);
  }

  if (code === 400) throw new Error("เกิดข้อผิดพลาดในการส่งคำขอ กรุณาลองใหม่อีกครั้ง");

  // 500/503/gateway — ไม่พยายาม JSON.parse body ที่อาจไม่ใช่ JSON
  throw new Error("เกิดข้อผิดพลาดจาก AI API (HTTP " + code + ") กรุณาลองใหม่อีกครั้ง");
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

  return ContentService.createTextOutput(JSON.stringify({
    result: 'success', appended: appended, updatedExisting: (rowIndex > 0)
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