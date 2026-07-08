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
  if (action == 'getRelatedQuestions') return getRelatedQuestionsData(e.parameter.subject); // Feature 4: relations map ต่อวิชา (อ่านอย่างเดียว, chunked cache)
  if (action == 'getKB') return getKBData(e.parameter.subject); // §1.8 KB corpus: chunks ต่อวิชา (public read, chunked cache)
  if (action == 'getGlossary') return getGlossaryData(e.parameter.subject); // Feature 2: glossary ต่อวิชา (public read, chunked cache)
  if (action == 'getHighYield') return getHighYieldData(e.parameter.category); // Feature 3: ชีทสรุป high-yield ต่อหมวด (public read, chunked cache)
  if (action == 'getKeywordIndex') return getKeywordIndexData(e.parameter.category); // Feature 6: คำสำคัญที่ออกบ่อย ต่อหมวด (public read, chunked cache — list)
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

    // บริจาค/อัปเดต IntelSphere key หนึ่งใบ (idempotent by API_Key) — public donation form, lock-free
    // Rate limit ก่อน validate — กันคนใช้ endpoint นี้เป็นเครื่องเดา key (key-testing oracle)
    if (action === 'seedIntelSphereKey') {
      if (!checkActionRateLimit('rl_seed_', data.sessionToken || 'anon', 3)) {
        return ContentService.createTextOutput(JSON.stringify({
          result: 'error', message: 'คุณส่งคำขอบริจาคบ่อยเกินไป (สูงสุด 3 ครั้ง/ชั่วโมง) กรุณาลองใหม่ภายหลัง'
        })).setMimeType(ContentService.MimeType.JSON);
      }
      return seedIntelSphereKey(data.apiKey, data.donorName, data.notes);
    }

    // อ่าน catalog โมเดล IntelSphere + รายชื่อผู้บริจาค (read-only, ไม่ต้อง auth) — lock-free
    if (action === 'listModels') {
      return ContentService.createTextOutput(JSON.stringify({
        result: 'success',
        catalog: getIntelSphereModelCatalog(),
        donors: getDonorCredits()
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // agentQuery — owner-only agentic traffic จาก claude-kkuintelsphere-router proxy (lock-free: อ่าน + quota-column write แบบเดียวกับ askAIExpert)
    // แยกจาก askAIExpert เพราะ (1) checkRateLimit 15/hr ต่ำเกินไปมากสำหรับ agentic coding
    // (2) ต้องบังคับ session token (branch IntelSphere ของ askAIExpert เป็น public)
    // (3) priority เป็น Claude-first เพื่อคุณภาพ ไม่ใช่ Deepseek-first เพื่อ cost แบบ INTELSPHERE_PROVIDER_PRIORITY
    if (action === 'agentQuery') {
      // Auth: non-expiring owner secret (preferred) OR legacy 30-day session token.
      // The secret's hash is kept in Script Properties — see generateAgentQueryOwnerSecret().
      var agentAuthed = verifyAgentQueryOwnerSecret(data.ownerSecret) || !!verifySessionToken(data.sessionToken);
      if (!agentAuthed) {
        return ContentService.createTextOutput(JSON.stringify({
          result: 'error', message: 'session_expired'
        })).setMimeType(ContentService.MimeType.JSON);
      }
      try {
        var agentResult = executeAgentQuery(data.request || {});
        return ContentService.createTextOutput(JSON.stringify({
          result: 'success', provider: agentResult.provider, completion: agentResult.completion
        })).setMimeType(ContentService.MimeType.JSON);
      } catch (agentErr) {
        return ContentService.createTextOutput(JSON.stringify({
          result: 'error', message: agentErr.message
        })).setMimeType(ContentService.MimeType.JSON);
      }
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
    // submitAiFeedback — append-only เดี่ยวแบบเดียวกับ batchLog จึงอยู่ lock-free tier
    // rate limit 30/ชม. ต่อ token; เกิน limit ตอบ success (dropped) เพราะ frontend เป็น fire-and-forget
    // ----------------------------------------------------
    if (action === 'submitAiFeedback') {
      if (!checkActionRateLimit('rl_fb_', data.sessionToken || 'anon', 30)) {
        return ContentService.createTextOutput(JSON.stringify({ result: 'success', dropped: true }))
          .setMimeType(ContentService.MimeType.JSON);
      }
      return submitAiFeedbackRow(data);
    }

    // ----------------------------------------------------
    // §1.8 ingestKB — เขียน KB_Chunks (logged-in-only). auth + rate-limit ทำ "นอก lock" (อ่านล้วน)
    // เพื่อไม่ให้ garbage-token flood ไปแย่ง shared localized lock ของ vote/report และไม่ให้บังคับ
    // reject-log ไม่จำกัด; ล็อกเฉพาะช่วงเขียนจริง → การเขียนยังอยู่ localized-15s tier ตามแผน (ไม่มี UrlFetchApp ใต้ lock)
    // ----------------------------------------------------
    if (action === 'ingestKB') {
      // rate-limit ก่อน (keyed by token/'anon') — กัน flood ทั้ง lock และ reject-log
      if (!checkActionRateLimit('rl_kb_', data.sessionToken || 'anon', 10)) {
        return ContentService.createTextOutput(JSON.stringify({
          result: 'error', message: 'อัปโหลดบ่อยเกินไป (สูงสุด 10 ครั้ง/ชั่วโมง) กรุณาลองใหม่ภายหลัง'
        })).setMimeType(ContentService.MimeType.JSON);
      }
      // AUTH GATE (server-side, ก่อนเขียน KB_Chunks ทุกกรณี) — ต้องมี session token ที่ยัง valid
      var kbUser = verifySessionToken(data.sessionToken);
      if (!kbUser) {
        try {
          writeAdminLog(data.sessionToken ? 'INVALID_TOKEN' : 'NO_TOKEN', 'GUEST', 'KB', 'INGEST_REJECT',
            'KB_Chunks', 'Unauthenticated ingestKB attempt (subject=' + (data.subject || '') +
            ', source=' + (data.source || '') + ')', '', '', '');
        } catch (kbLogErr) {}
        return ContentService.createTextOutput(JSON.stringify({
          result: 'error', message: 'ต้องเข้าสู่ระบบก่อนอัปโหลด'
        })).setMimeType(ContentService.MimeType.JSON);
      }
      // ผ่าน auth แล้ว → ล็อก localized-15s เฉพาะช่วงเขียน
      var kbLock = LockService.getScriptLock();
      if (!kbLock.tryLock(15000)) {
        return ContentService.createTextOutput(JSON.stringify({
          result: 'error', message: 'เซิร์ฟเวอร์ไม่ตอบสนองเนื่องจากโหลดสูง (Lock Timeout)'
        })).setMimeType(ContentService.MimeType.JSON);
      }
      try {
        var kbRes = ingestKBChunks(data.subject, data.source, data.markdown, data.categoryId); // §1.9: categoryId optional
        return ContentService.createTextOutput(JSON.stringify(kbRes)).setMimeType(ContentService.MimeType.JSON);
      } finally {
        kbLock.releaseLock();
      }
    }

    // ----------------------------------------------------
    // §2.5 askGlossaryTerm — tap/select miss-path (public, self-populating). โครงเดียวกับ ingestKB:
    // rate-limit + LLM ทำ "นอก lock"; ล็อกเฉพาะช่วงเขียน 1 แถว. ***ห้ามยิง LLM ใต้ lock (advisor rule)***
    // ไม่อยู่ใน admin tier — guest ต้อง tap แปลได้ (นี่คือ UX หลักของ feature). dedup ก่อนยิง LLM = ประหยัดโทเคน
    // NOTE (shared-bucket tradeoff): guest ทุกคน key เป็น 'anon' → rl_glossary_anon เป็นถังเดียว "รวม" 20/ชม.
    //   ทั้ง guest ทุกคน; ผู้ล็อกอินได้ถังของตัวเอง (key=token). ยอมรับได้เพราะ hit path เป็น client-side ไม่จำกัด
    //   → feature degrade เป็น instant-lookup เมื่อ glossary แน่นขึ้น และผู้ใช้ KKU ส่วนใหญ่ล็อกอิน
    // ----------------------------------------------------
    if (action === 'askGlossaryTerm') {
      var glToken = data.sessionToken || 'anon';
      if (!checkActionRateLimit('rl_glossary_', glToken, GLOSSARY_ASK_RATE_LIMIT)) {
        return ContentService.createTextOutput(JSON.stringify({
          result: 'error', message: 'ค้นหาคำศัพท์บ่อยเกินไป (สูงสุด ' + GLOSSARY_ASK_RATE_LIMIT + ' ครั้ง/ชั่วโมง) กรุณาลองใหม่ภายหลัง'
        })).setMimeType(ContentService.MimeType.JSON);
      }
      var glWordKey = normalizeGlossaryTerm(data.word);
      if (!glWordKey) {
        return ContentService.createTextOutput(JSON.stringify({
          result: 'error', message: 'ไม่พบคำที่เลือก'
        })).setMimeType(ContentService.MimeType.JSON);
      }
      // (1) dedup ก่อนยิง LLM (อ่านผ่าน getGlossary cache — ไม่ getDataRange ทุก tap) → ไม่เปลืองโทเคนถ้ามีแล้ว
      var glHit = glossaryLookupCached(data.subject, glWordKey);
      if (glHit) {
        return ContentService.createTextOutput(JSON.stringify({
          result: 'success', term: glHit, cached: true
        })).setMimeType(ContentService.MimeType.JSON);
      }
      // (2) LLM lock-free (cheap flash tier) — ผ่าน IntelSphere path เดียวกับ askAIExpert
      var glParsed;
      try {
        var glRaw = executeChatbotQuery(buildGlossaryTermPrompt(data.word, data.sentence), GLOSSARY_MODEL, 1);
        glParsed = parseGlossaryJson(glRaw.content);
      } catch (glErr) {
        return ContentService.createTextOutput(JSON.stringify({
          result: 'error', message: (glErr && glErr.message) || 'AI ไม่พร้อมใช้งาน กรุณาลองใหม่'
        })).setMimeType(ContentService.MimeType.JSON);
      }
      if (!glParsed || !glParsed.en) {
        return ContentService.createTextOutput(JSON.stringify({
          result: 'error', message: 'ไม่สามารถแปลคำนี้ได้ กรุณาลองใหม่'
        })).setMimeType(ContentService.MimeType.JSON);
      }
      // (2b) dedup post-LLM บน canonical en (เผื่อโมเดล normalize ไปตรงแถวที่มีอยู่แล้ว) — ยังไม่ต้อง lock
      var glCanonHit = glossaryLookupCached(data.subject, normalizeGlossaryTerm(glParsed.en));
      if (glCanonHit) {
        return ContentService.createTextOutput(JSON.stringify({
          result: 'success', term: glCanonHit, cached: true
        })).setMimeType(ContentService.MimeType.JSON);
      }
      // (3) เขียน 1 แถวใต้ localized-15s lock (ล็อกเฉพาะการเขียน; re-check dedup race-safe ข้างใน)
      var glLock = LockService.getScriptLock();
      if (!glLock.tryLock(15000)) {
        return ContentService.createTextOutput(JSON.stringify({
          result: 'error', message: 'เซิร์ฟเวอร์ไม่ตอบสนองเนื่องจากโหลดสูง (Lock Timeout)'
        })).setMimeType(ContentService.MimeType.JSON);
      }
      try {
        var glWrite = writeGlossaryRowLocked(data.subject, glParsed, data.questionId);
        return ContentService.createTextOutput(JSON.stringify({
          result: 'success', term: glWrite.term, cached: glWrite.cached
        })).setMimeType(ContentService.MimeType.JSON);
      } finally {
        glLock.releaseLock();
      }
    }

    // ----------------------------------------------------
    // §3.6 generateHighYield — lazy-generate-then-cache miss-path (public, self-populating). โครงเดียวกับ askGlossaryTerm:
    // rate-limit → dedup(cache) ก่อนยิง LLM → LLM ทำ "นอก lock" → เขียน 1 แถวใต้ localized-15s lock. ***ห้ามยิง LLM ใต้ lock***
    // ไม่อยู่ใน admin tier — guest กดสร้างชีทสรุปได้ (เป็น UX หลักของ feature). subject resolve จาก categoryId ฝั่ง server
    // ----------------------------------------------------
    if (action === 'generateHighYield') {
      var hyToken = data.sessionToken || 'anon';
      if (!checkActionRateLimit('rl_highyield_', hyToken, HIGHYIELD_GEN_RATE_LIMIT)) {
        return ContentService.createTextOutput(JSON.stringify({
          result: 'error', message: 'สร้างชีทสรุปบ่อยเกินไป (สูงสุด ' + HIGHYIELD_GEN_RATE_LIMIT + ' ครั้ง/ชั่วโมง) กรุณาลองใหม่ภายหลัง'
        })).setMimeType(ContentService.MimeType.JSON);
      }
      var hyCat = String(data.category || '').trim();
      if (!hyCat) {
        return ContentService.createTextOutput(JSON.stringify({
          result: 'error', message: 'ไม่พบหัวข้อ (category)'
        })).setMimeType(ContentService.MimeType.JSON);
      }
      // (1) dedup ก่อนยิง LLM (อ่านผ่าน getHighYield cache) → ไม่เปลืองโทเคนถ้ามีแล้ว
      var hyHit = highYieldLookupCached(hyCat);
      if (hyHit) {
        return ContentService.createTextOutput(JSON.stringify({
          result: 'success', highyield: hyHit, cached: true
        })).setMimeType(ContentService.MimeType.JSON);
      }
      // (2) รวมข้อสอบของหมวดนี้ (อ่านอย่างเดียว, lock-free) — subject resolve จาก categoryId ฝั่ง server (fallback = data.subject)
      var hyText = aggregateCategoryText(hyCat, data.subject);
      if (!hyText) {
        return ContentService.createTextOutput(JSON.stringify({
          result: 'error', message: 'ไม่พบข้อสอบในหัวข้อนี้ จึงยังสร้างชีทสรุปไม่ได้'
        })).setMimeType(ContentService.MimeType.JSON);
      }
      // (3) LLM lock-free (cheap flash tier, max_tokens สูงกว่า default กัน JSON ขาด) — ผ่าน IntelSphere path เดียวกับ askAIExpert
      var hyParsed;
      try {
        var hyRaw = executeChatbotQuery(buildHighYieldPrompt(hyText), HIGHYIELD_MODEL, 1, HIGHYIELD_MAX_TOKENS);
        hyParsed = parseHighYieldJson(hyRaw.content);
      } catch (hyErr) {
        return ContentService.createTextOutput(JSON.stringify({
          result: 'error', message: (hyErr && hyErr.message) || 'AI ไม่พร้อมใช้งาน กรุณาลองใหม่'
        })).setMimeType(ContentService.MimeType.JSON);
      }
      if (!hyParsed || !hyParsed.summary_md) {
        return ContentService.createTextOutput(JSON.stringify({
          result: 'error', message: 'สร้างชีทสรุปไม่สำเร็จ (คำตอบ AI ไม่สมบูรณ์) กรุณาลองใหม่'
        })).setMimeType(ContentService.MimeType.JSON);
      }
      // (4) เขียน 1 แถวใต้ localized-15s lock (ล็อกเฉพาะการเขียน; re-check dedup race-safe ข้างใน)
      var hyLock = LockService.getScriptLock();
      if (!hyLock.tryLock(15000)) {
        return ContentService.createTextOutput(JSON.stringify({
          result: 'error', message: 'เซิร์ฟเวอร์ไม่ตอบสนองเนื่องจากโหลดสูง (Lock Timeout)'
        })).setMimeType(ContentService.MimeType.JSON);
      }
      try {
        var hyWrite = writeHighYieldRowLocked(hyCat, hyParsed, getVersionCached());
        return ContentService.createTextOutput(JSON.stringify({
          result: 'success', highyield: hyWrite.row, cached: hyWrite.cached
        })).setMimeType(ContentService.MimeType.JSON);
      } finally {
        hyLock.releaseLock();
      }
    }

    // ----------------------------------------------------
    // §3.4 voteHighYieldMnemonic — 👍/🚩 mnemonic (public). localized-15s tier (standalone block, mirror askGlossaryTerm write-lock).
    // อ่านคอลัมน์ H (Mnemonic_Votes JSON) "สดจากชีตใต้ lock" → mutate {idx:net+delta} → เขียนกลับ → invalidate cache.
    // ***ห้ามอ่านจาก cache*** (สอง user โหวต mnemonic คนละตัวของหมวดเดียวกันพร้อมกันจะทับกัน). re-vote guard เป็น localStorage ฝั่ง client
    // ----------------------------------------------------
    if (action === 'voteHighYieldMnemonic') {
      var mvToken = data.sessionToken || 'anon';
      if (!checkActionRateLimit('rl_hymvote_', mvToken, HIGHYIELD_VOTE_RATE_LIMIT)) {
        return ContentService.createTextOutput(JSON.stringify({
          result: 'error', message: 'โหวตบ่อยเกินไป กรุณาลองใหม่ภายหลัง'
        })).setMimeType(ContentService.MimeType.JSON);
      }
      var mvCat = String(data.category || '').trim();
      var mvIdx = parseInt(data.mnemonicIdx, 10);
      var mvDelta = (parseInt(data.delta, 10) < 0) ? -1 : 1; // clamp เป็น ±1
      if (!mvCat || isNaN(mvIdx) || mvIdx < 0) {
        return ContentService.createTextOutput(JSON.stringify({
          result: 'error', message: 'ข้อมูลโหวตไม่ถูกต้อง'
        })).setMimeType(ContentService.MimeType.JSON);
      }
      var mvLock = LockService.getScriptLock();
      if (!mvLock.tryLock(15000)) {
        return ContentService.createTextOutput(JSON.stringify({
          result: 'error', message: 'เซิร์ฟเวอร์ไม่ตอบสนองเนื่องจากโหลดสูง (Lock Timeout)'
        })).setMimeType(ContentService.MimeType.JSON);
      }
      try {
        var mvRes = voteHighYieldMnemonicLocked(mvCat, mvIdx, mvDelta); // อ่านสด+เขียนใต้ lock
        return ContentService.createTextOutput(JSON.stringify(mvRes)).setMimeType(ContentService.MimeType.JSON);
      } finally {
        mvLock.releaseLock();
      }
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

      var adminActions = ['editQuestion', 'deleteQuestion', 'addCategory', 'adminImport', 'updateReportStatus', 'deleteCategory', 'updateCategory', 'deleteGroup', 'updateAccordionGroup', 'addSubject', 'updateSubject', 'deleteSubject', 'addAnnouncement', 'editAnnouncement', 'deleteAnnouncement', 'runRelationsBatchManual', 'runGlossaryBatchManual', 'runHighYieldBatchManual', 'runKeywordIndexBatchManual'];
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

        // Feature 4: สั่งสร้าง relations ของวิชาเดียวแบบ manual เพื่อทดสอบ/populate โดยไม่ต้องรอ cron
        // อยู่ใน admin tier (ต้องมี session ที่ถูกต้อง); เป็น write จึงทำใต้ lock. ใช้กับวิชาเล็กเท่านั้น
        // (วิชาใหญ่ เช่น CVS 2516 ข้อ ควรใช้ nightly runQuestionRelationsBatch เพื่อเลี่ยงถือ lock นาน)
        if (action === 'runRelationsBatchManual') {
          var relSubj = data.subject;
          if (!relSubj) {
            return ContentService.createTextOutput(JSON.stringify({
              result: 'error', message: 'ต้องระบุ subject'
            })).setMimeType(ContentService.MimeType.JSON);
          }
          var relCount = generateQuestionRelationsForSubject(relSubj);
          writeAdminLog(user, userRole, "RELATIONS", "GENERATE", relSubj, "Generated question relations", "", String(relCount), metadata);
          return ContentService.createTextOutput(JSON.stringify({
            result: 'success', subject: relSubj, relationRows: relCount
          })).setMimeType(ContentService.MimeType.JSON);
        }

        // Feature 2: สั่งสกัด glossary ของวิชาเดียวแบบ manual (ทดสอบ/populate) โดยไม่ต้องรอ cron
        // ⚠️ ยิง LLM "ใต้ admin lock 25s" = admin wall ที่แผนยอมรับ — ใช้เฉพาะวิชาเล็ก/ทดสอบเท่านั้น
        // (nightly runGlossaryBatch วิ่งผ่าน time-driven trigger ไม่ผ่าน doPost จึง "ไม่ถือ lock";
        //  trigger เว้นไว้ไม่ติดตั้ง — tap/select path (askGlossaryTerm) เติม glossary เองแบบ lock-free)
        if (action === 'runGlossaryBatchManual') {
          var glSubj = data.subject;
          if (!glSubj) {
            return ContentService.createTextOutput(JSON.stringify({
              result: 'error', message: 'ต้องระบุ subject'
            })).setMimeType(ContentService.MimeType.JSON);
          }
          var glWritten = generateGlossaryForSubject(glSubj);
          writeAdminLog(user, userRole, "GLOSSARY", "GENERATE", glSubj, "Generated glossary terms", "", String(glWritten), metadata);
          return ContentService.createTextOutput(JSON.stringify({
            result: 'success', subject: glSubj, termsWritten: glWritten
          })).setMimeType(ContentService.MimeType.JSON);
        }

        // Feature 3: สั่งสร้างชีทสรุป high-yield ของ "ทุกหมวดในวิชาเดียว" แบบ manual (ทดสอบ/populate) โดยไม่ต้องรอ cron
        // ⚠️ ยิง LLM (หลาย call ต่อหมวด) "ใต้ admin lock 25s" = admin wall ที่แผนยอมรับ — ใช้เฉพาะวิชาเล็ก/ทดสอบเท่านั้น
        // (วิชาใหญ่ควรใช้ nightly runHighYieldBatch ผ่าน time-driven trigger ซึ่งไม่ผ่าน doPost จึงไม่ถือ lock; trigger เว้นไว้ไม่ติดตั้ง)
        if (action === 'runHighYieldBatchManual') {
          var hySubj = data.subject;
          if (!hySubj) {
            return ContentService.createTextOutput(JSON.stringify({
              result: 'error', message: 'ต้องระบุ subject'
            })).setMimeType(ContentService.MimeType.JSON);
          }
          var hyGen = generateHighYieldForSubject(hySubj);
          writeAdminLog(user, userRole, "HIGHYIELD", "GENERATE", hySubj, "Generated high-yield sheets", "", String(hyGen), metadata);
          return ContentService.createTextOutput(JSON.stringify({
            result: 'success', subject: hySubj, categoriesWritten: hyGen
          })).setMimeType(ContentService.MimeType.JSON);
        }

        // Feature 6: สั่งสร้าง keyword index ของ "ทุกหมวดในวิชาเดียว" แบบ manual (§6.2 idle-day/admin เท่านั้น) โดยไม่ต้องรอ cron
        // token-free 100% (ไม่ยิง LLM) → เร็ว ไม่ต้อง checkpoint. แต่วนทุกหมวด × ทุก term × ทุกข้อ "ใต้ admin lock 25s"
        // ⚠️ วิชาใหญ่ (เช่น CVS 2516 ข้อ หลายหมวด) อาจถือ lock นาน — ใช้เฉพาะวิชาเล็ก/ทดสอบ; ไม่มี public gen endpoint / ไม่ติดตั้ง trigger
        if (action === 'runKeywordIndexBatchManual') {
          var kwSubj = data.subject;
          if (!kwSubj) {
            return ContentService.createTextOutput(JSON.stringify({
              result: 'error', message: 'ต้องระบุ subject'
            })).setMimeType(ContentService.MimeType.JSON);
          }
          var kwGen = generateKeywordIndexForSubject(kwSubj);
          writeAdminLog(user, userRole, "KEYWORDINDEX", "GENERATE", kwSubj, "Generated keyword index", "", String(kwGen.rows), metadata);
          return ContentService.createTextOutput(JSON.stringify({
            result: 'success', subject: kwSubj, categoriesProcessed: kwGen.categories, rowsWritten: kwGen.rows
          })).setMimeType(ContentService.MimeType.JSON);
        }

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
var AI_FEEDBACK_SHEET_NAME = "AI_Feedback";

// ── Feature 4: Related-Questions relations (token-free v1) ─────────────
var QUESTION_RELATIONS_SHEET_NAME = "Question_Relations";
var RELATIONS_TOPK = 5;                 // เก็บสูงสุด k ความสัมพันธ์ต่อข้อ
var RELATIONS_MIN_SCORE = 2;            // เกณฑ์คะแนนขั้นต่ำ (จำนวน token ร่วม) จึงจะถือว่าเกี่ยวข้อง
var RELATIONS_MIN_SHARED_TOKENS = 2;    // prefilter: คู่ที่จะนำมาให้คะแนนต้องแชร์ token >= ค่านี้
var RELATIONS_MAX_POSTINGS = 120;       // token ที่ปรากฏในเอกสารมากกว่านี้ = stopword-like ข้ามไป (กัน N^2 ระเบิด)
var RELATIONS_CANDIDATE_CAP = 60;       // เพดานจำนวน candidate ที่ให้คะแนนต่อข้อ
var RELATIONS_BATCH_BUDGET_MS = 300000; // งบเวลาต่อรอบ ~5 นาที (ต่ำกว่าลิมิต 6 นาทีของ GAS)
var RELATIONS_CHECKPOINT_KEY = "RELATIONS_BATCH_CHECKPOINT";

// ── §1.8: Knowledge-base corpus (Markdown textbook/lecture chunks, token-free v1) ──
// เสิร์ฟผ่าน getKB (doGet, chunked cache); เขียนผ่าน ingestKB (doPost, ต้อง login ก่อน)
var KB_CHUNKS_SHEET_NAME = "KB_Chunks";
var KB_CHUNK_MAX_WORDS = 500;   // section ที่ยาวเกินนี้ถูกตัดเป็น chunk ย่อย (คุม top-k ให้ถูก); ~200-500 คำ/chunk

// ── Feature 2: Glossary (root-word + Thai↔English, §2.1–§2.6) ──
// tap/select miss-path = askGlossaryTerm (public, standalone block, LLM lock-free); เสิร์ฟผ่าน getGlossary
var GLOSSARY_SHEET_NAME = "Glossary";
var GLOSSARY_MODEL = "gemini-2.5-flash";  // cheap flash tier; executeChatbotQuery rotate ต่อถ้า Gemini หมดโควต้า
var GLOSSARY_ASK_RATE_LIMIT = 20;         // ต่อ token/'anon' ต่อชั่วโมง (public token-spending + write; กัน spam)
var GLOSSARY_BATCH_BUDGET_MS = 300000;    // งบเวลา nightly ~5 นาที (ต่ำกว่าลิมิต 6 นาทีของ GAS)
var GLOSSARY_CHECKPOINT_KEY = "GLOSSARY_BATCH_CHECKPOINT";
var GLOSSARY_BATCH_CHARS = 6000;          // ขนาดก้อนข้อความต่อ 1 LLM call ตอน batch (กันโพรมป์ยาวเกิน context)

// ── Feature 3: High-yield cram sheet (§3.1–§3.6) ──
// lazy-generate miss-path = generateHighYield (public, standalone block, LLM lock-free) — mirror askGlossaryTerm
// mnemonic vote = voteHighYieldMnemonic (standalone block, 15s tryLock = localized tier); เสิร์ฟผ่าน getHighYield (doGet, chunked cache)
// batch = generateHighYieldForSubject / runHighYieldBatch (checkpointed) + runHighYieldBatchManual (admin tier). trigger เว้นไว้ไม่ติดตั้ง
var HIGHYIELD_SHEET_NAME = "HighYield_Cache";
var HIGHYIELD_MODEL = "gemini-2.5-flash";  // cheap flash tier — เหมือน glossary; executeChatbotQuery rotate ต่อถ้าโควต้าหมด
var HIGHYIELD_MAX_TOKENS = 1800;           // output ก้อนใหญ่ (summary+mnemonics+keywords) — มากกว่า default 800 เพื่อกัน JSON ขาดกลาง
var HIGHYIELD_GEN_RATE_LIMIT = 6;          // ต่อ token/'anon' ต่อชั่วโมง (call ใหญ่/แพงกว่า glossary มาก → เข้มกว่า 20)
var HIGHYIELD_VOTE_RATE_LIMIT = 40;        // 👍/🚩 mnemonic เบามาก — กัน spam อย่างเดียว
var HIGHYIELD_MAX_QUESTIONS = 80;          // เพดานจำนวนข้อที่รวมต่อ 1 หมวด (IntelSphere context เล็ก) — เลือกข้อมีเฉลยก่อน
var HIGHYIELD_MAX_CHARS = 6000;            // เพดานตัวอักษรที่ป้อน LLM (บังคับก่อน MAX_QUESTIONS)
var HIGHYIELD_MAX_MNEMONICS = 6;           // จำกัดจำนวน mnemonics ที่เก็บ (คุม output + vote index)
var HIGHYIELD_MAX_KEYWORDS = 15;           // จำกัดจำนวน keywords ที่เก็บ
var HIGHYIELD_BATCH_BUDGET_MS = 300000;    // งบเวลา nightly ~5 นาที (ต่ำกว่าลิมิต 6 นาทีของ GAS)
var HIGHYIELD_CHECKPOINT_KEY = "HIGHYIELD_BATCH_CHECKPOINT";

// ── Feature 6: Frequently-tested keyword index (§6.1–§6.4) — token-free ทั้งหมด (ไม่ยิง LLM เลย) ──
// term set = HighYield keywords ของหมวด ∪ Glossary terms ของวิชา (สกัดจาก 2 pass เดิม, ไม่ pass ที่ 3)
// นับความถี่ lexical: EN → word-boundary regex, TH → substring (ไทยไม่มี word boundary; v1 characteristic)
// เสิร์ฟผ่าน getKeywordIndex (doGet, chunked cache); gen ผ่าน runKeywordIndexBatchManual (admin tier เท่านั้น — §6.2 idle-day/admin, ไม่มี public endpoint, ไม่มี trigger)
var KEYWORD_INDEX_SHEET_NAME = "Keyword_Index";
var KEYWORD_MIN_LEN_EN = 2;   // §6.4 min length กัน keyword สั้น/กำกวม (EN มี word-boundary ป้องกันอยู่แล้ว → 2 พอ เช่น "MI")
var KEYWORD_MIN_LEN_TH = 3;   // TH ใช้ substring (อันตรายกว่า) → ต้องยาว ≥3 กัน match มั่วทั่ว

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

  // maxTokens optional (default 800) — high-yield ต้องการ output ยาวกว่า glossary/chatbot (summary+mnemonics+keywords ก้อนเดียว)
  var payload = { model: actualModel, messages: [{ role: "user", content: prompt }], max_tokens: maxTokens || 800, temperature: 0.3 };
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

// รวมโควต้าคงเหลือรายวันต่อ provider (ทุก donor key ที่ Active) — ใช้จัดลำดับ chain แบบ load-balance
// นับเฉพาะ key ที่เกิน quota floor (เกณฑ์เดียวกับ getActiveIntelSphereKey) — ต่ำกว่า floor คือ serve ไม่ได้จริง
// ข้อจำกัดที่ยอมรับ: อ่านค่าก่อน daily reset (reset เกิดใน getActiveIntelSphereKey ทีหลัง) —
// request แรกของวันอาจเรียงด้วยค่าค้างของเมื่อวาน แล้วหายเองใน request ถัดไป
function getIntelSphereQuotaTotals() {
  var totals = {};
  var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(INTELSPHERE_SHEET_NAME);
  if (!sheet) return totals;
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var colStatus = headers.indexOf("Status");
  for (var i = 1; i < data.length; i++) {
    if (data[i][colStatus] !== "Active") continue;
    for (var j = 0; j < AGENT_QUERY_PROVIDER_PRIORITY.length; j++) {
      var p = AGENT_QUERY_PROVIDER_PRIORITY[j];
      var col = headers.indexOf(p + "_Remaining");
      if (col < 0) continue;
      var rem = Number(data[i][col]);
      var floor = INTELSPHERE_LIMITS[p] * INTELSPHERE_QUOTA_FLOOR;
      if (!isNaN(rem) && rem > floor) totals[p] = (totals[p] || 0) + rem;
    }
  }
  return totals;
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
  var order = orderAgentProviders(request, getIntelSphereQuotaTotals());
  var maxAttempts = order.length * 2; // เผื่อหลาย donor key ต่อ provider

  while (attempts < maxAttempts) {
    attempts++;

    var keyObj = null, provider = null;
    for (var i = 0; i < order.length; i++) {
      var p = order[i];
      if (skip[p]) continue;
      var candidate = getActiveIntelSphereKey(p);
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
function setupAiFeedbackSheet() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(AI_FEEDBACK_SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(AI_FEEDBACK_SHEET_NAME);

  var headers = ["Timestamp", "Rating", "Model", "Subject", "QuestionId", "Prompt_Snippet", "Answer_Snippet", "SessionToken"];
  if (!sheet.getRange(1, 1).getValue()) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#e6f7ff");
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// บันทึก feedback (ดี/เฉยๆ/แย่) ต่อคำตอบ AI หนึ่งฟอง — append-only, ไม่ต้อง lock (แบบเดียวกับ batchLog)
function submitAiFeedbackRow(data) {
  if (['good', 'neutral', 'bad'].indexOf(data.rating) < 0) {
    return ContentService.createTextOutput(JSON.stringify({
      result: 'error', message: 'ค่า rating ไม่ถูกต้อง'
    })).setMimeType(ContentService.MimeType.JSON);
  }
  var sheet = setupAiFeedbackSheet(); // lazy-create ครั้งแรก
  sheet.appendRow([
    new Date(),
    data.rating,
    String(data.model || "").slice(0, 100),
    String(data.subject || "").slice(0, 100),
    String(data.questionId || "").slice(0, 100),
    String(data.promptSnippet || "").slice(0, 300),
    String(data.answerSnippet || "").slice(0, 300),
    String(data.sessionToken || "").slice(0, 64)
  ]);
  return ContentService.createTextOutput(JSON.stringify({ result: 'success' }))
    .setMimeType(ContentService.MimeType.JSON);
}

/* =========================================================================
   FEATURE 4 — Related-Questions cross-reference (token-free v1)
   ยิงคะแนนความคล้ายแบบ lexical ต่อคู่ข้อสอบในวิชาเดียวกัน แล้วเก็บ top-k ลงชีต
   Question_Relations. เสิร์ฟผ่าน getRelatedQuestions (doGet, อ่านอย่างเดียว).
   หมายเหตุ: GAS ไม่มี Intl.Segmenter → tokenize = lowercase + split บน whitespace/
   เครื่องหมายวรรคตอน และจับคู่แบบ "token เท่ากัน" (ไม่ใช่ substring แบบ §1.6 บน
   frontend) — เป็นเงื่อนไขที่ทำให้ prefilter O(N^2)-safe ได้ แต่แลกกับ recall ของ
   ไทยล้วน (คำไทยยาวๆ ที่ไม่มี term อังกฤษ/ตัวเลขร่วมจะได้ relation น้อย); v2 vectors แก้ทีหลัง
   ========================================================================= */

// idempotent: สร้างชีต Question_Relations ถ้ายังไม่มี (mirror setupAiFeedbackSheet)
function setupQuestionRelationsSheet() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(QUESTION_RELATIONS_SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(QUESTION_RELATIONS_SHEET_NAME);

  var headers = ["Question_ID", "Related_ID", "Score", "Shared_Concept", "Src_Version", "Updated_At"];
  if (!sheet.getRange(1, 1).getValue()) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#e6f7ff");
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// tokenize สำหรับ relations (ฝั่ง server): lowercase + split บนอักขระที่ไม่ใช่ a-z0-9 หรือช่วงไทย ก-๙
// คืน array ของ token "ไม่ซ้ำ" ความยาว >= 2 (คำสั้นมากตัดทิ้งกัน noise)
function relationsTokenize(text) {
  text = (text == null ? "" : String(text)).toLowerCase();
  var raw = text.split(/[^a-z0-9฀-๿]+/);
  var seen = {};
  var out = [];
  for (var i = 0; i < raw.length; i++) {
    var t = raw[i];
    if (t.length < 2) continue;
    if (seen[t]) continue;
    seen[t] = true;
    out.push(t);
  }
  return out;
}

// ล้าง relations cache ของวิชาหนึ่ง (บังคับให้ getRelatedQuestions ดึงใหม่หลัง generate)
// เอา marker "_chunks" ออกก็พอ — getLargeCache จะถือเป็น miss ทันที (chunk ที่เหลือหมดอายุเอง)
function invalidateRelationsCache(subject) {
  var v = getVersionCached();
  var cleanFilter = subject ? String(subject).trim().toUpperCase() : "all";
  try { CacheService.getScriptCache().remove("relations_" + v + "_" + cleanFilter + "_chunks"); } catch (e) {}
}

// สร้าง relations ของวิชาเดียว: อ่านข้อสอบ (ผ่าน reader เดียวกับ getQuestions) → inverted index
// prefilter → ให้คะแนน token ร่วม → เก็บ top-k → เขียนทับเฉพาะแถวของวิชานี้ในชีต. คืนจำนวนแถวที่เขียน
function generateQuestionRelationsForSubject(subjectId) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var srcVersion = getVersionCached();

  // ใช้ data source เดียวกับ getQuestions (มี cache) — parse JSON ที่ ContentService คืนมา
  var qJson = getQuestionsDataCached(subjectId, ss).getContent();
  var questions = [];
  try { questions = JSON.parse(qJson) || []; } catch (e) { questions = []; }

  // เตรียม docs: token ไม่ซ้ำต่อข้อ จาก problem + choices + explain + answer
  var docs = [];
  for (var i = 0; i < questions.length; i++) {
    var q = questions[i];
    var blob = (q.problem || "") + " " + (q.choices || "") + " " + (q.explain || "") + " " + (q.answer || "");
    docs.push({ id: q.questionId, tokens: relationsTokenize(blob) });
  }

  var n = docs.length;
  var subjIdSet = {}; // id ของวิชานี้ (ใช้ตอนเขียนทับชีต)
  for (var s = 0; s < n; s++) subjIdSet[String(docs[s].id)] = true;

  // สร้าง inverted index token -> [docIndex...]
  var postings = {};
  for (var d = 0; d < n; d++) {
    var toks = docs[d].tokens;
    for (var t = 0; t < toks.length; t++) {
      var key = toks[t];
      (postings[key] || (postings[key] = [])).push(d);
    }
  }

  var nowIso = new Date().toISOString();
  var newRows = [];

  for (var a = 0; a < n; a++) {
    var aTokens = docs[a].tokens;
    var shared = {}; // docIndex -> จำนวน token ร่วมกับ a

    for (var at = 0; at < aTokens.length; at++) {
      var plist = postings[aTokens[at]];
      if (!plist) continue;
      // ข้าม token ที่พบบ่อยเกิน (stopword-like) — ป้องกัน candidate ระเบิดแบบ N^2
      if (plist.length > RELATIONS_MAX_POSTINGS) continue;
      for (var p = 0; p < plist.length; p++) {
        var b = plist[p];
        if (b === a) continue;
        shared[b] = (shared[b] || 0) + 1;
      }
    }

    // กรอง candidate ที่แชร์ token ถึงเกณฑ์ แล้วจัด top-k
    var cands = [];
    for (var bIdx in shared) {
      var sc = shared[bIdx];
      if (sc >= RELATIONS_MIN_SHARED_TOKENS && sc >= RELATIONS_MIN_SCORE) {
        cands.push({ b: parseInt(bIdx, 10), score: sc });
      }
    }
    if (cands.length > RELATIONS_CANDIDATE_CAP) {
      cands.sort(function (x, y) { return y.score - x.score; });
      cands = cands.slice(0, RELATIONS_CANDIDATE_CAP);
    }
    cands.sort(function (x, y) { return y.score - x.score; });
    var top = cands.slice(0, RELATIONS_TOPK);

    for (var c = 0; c < top.length; c++) {
      newRows.push([docs[a].id, docs[top[c].b].id, top[c].score, "", srcVersion, nowIso]);
    }
  }

  // เขียนทับเฉพาะแถวของวิชานี้: อ่านทั้งชีต, เก็บแถวของวิชาอื่นไว้, ต่อด้วยแถวใหม่, clear แล้ว setValues ครั้งเดียว
  var sheet = setupQuestionRelationsSheet();
  var existing = sheet.getDataRange().getValues();
  var header = existing.length ? existing[0] : ["Question_ID", "Related_ID", "Score", "Shared_Concept", "Src_Version", "Updated_At"];
  var kept = [];
  for (var r = 1; r < existing.length; r++) {
    var rowQid = String(existing[r][0]).trim();
    if (!subjIdSet[rowQid]) kept.push(existing[r]); // แถวของวิชาอื่น — คงไว้
  }

  var finalRows = kept.concat(newRows);
  // clear เนื้อหาเดิม (นับรวม header) ก่อน แล้วเขียนใหม่ — กันแถวเก่าค้างเมื่อจำนวนแถวใหม่น้อยกว่าเดิม
  if (existing.length > 0) {
    sheet.getRange(1, 1, existing.length, Math.max(existing[0].length, 6)).clearContent();
  }
  sheet.getRange(1, 1, 1, header.length).setValues([header]);
  if (finalRows.length > 0) {
    sheet.getRange(2, 1, finalRows.length, 6).setValues(finalRows);
  }

  invalidateRelationsCache(subjectId);
  return newRows.length;
}

// รายชื่อ subjectId ทั้งหมดจากชีต Structure พร้อมจำนวนข้อ (ไว้เรียงเล็ก→ใหญ่ ลดโอกาส starve วิชาใหญ่)
function getSubjectsSortedBySize(ss) {
  if (!ss) ss = SpreadsheetApp.openById(SHEET_ID);

  var structRows = getStructureSheetDataCached(ss);
  var subjSet = {};
  for (var i = 1; i < structRows.length; i++) {
    var sid = String(structRows[i][1]).trim().toUpperCase();
    if (sid) subjSet[sid] = true;
  }

  // นับจำนวนข้อต่อวิชาจาก all questions + category->subject map (อ่านจาก cache ทั้งคู่)
  var catToSubj = getCategoryToSubjectMapCached(ss);
  var qData = getAllQuestionsCached(ss);
  var counts = {};
  for (var q = 0; q < qData.length; q++) {
    var catRaw = String(qData[q][6] || "").trim();
    var cats = [];
    try { cats = (catRaw.indexOf("[") > -1) ? JSON.parse(catRaw.replace(/'/g, '"')) : (catRaw ? [catRaw] : []); }
    catch (e) { cats = catRaw ? [catRaw] : []; }
    var counted = {};
    for (var k = 0; k < cats.length; k++) {
      var subjOf = catToSubj[cats[k]] || "";
      if (subjOf && subjSet[subjOf] && !counted[subjOf]) { counts[subjOf] = (counts[subjOf] || 0) + 1; counted[subjOf] = true; }
    }
  }

  var list = Object.keys(subjSet).map(function (sid) { return { subject: sid, count: counts[sid] || 0 }; });
  // เรียงเล็กก่อน (ทำวิชาเล็กให้ครบก่อน) — tiebreak ด้วยชื่อวิชาให้ลำดับ deterministic ระหว่างรอบ
  list.sort(function (x, y) { return (x.count - y.count) || (x.subject < y.subject ? -1 : 1); });
  return list;
}

// nightly entry point: ไล่ทีละวิชา, checkpoint ใน PropertiesService, เคารพงบเวลา ~5 นาที แล้วไปต่อรอบหน้า
function runQuestionRelationsBatch() {
  var deadline = Date.now() + RELATIONS_BATCH_BUDGET_MS;
  var srcVersion = getVersionCached();
  var props = PropertiesService.getScriptProperties();

  var ckpt = {};
  try { ckpt = JSON.parse(props.getProperty(RELATIONS_CHECKPOINT_KEY) || "{}"); } catch (e) { ckpt = {}; }
  // version drift → เริ่มรอบใหม่ทั้งหมด (relations เก่าถือว่า stale)
  if (ckpt.srcVersion !== srcVersion) ckpt = { srcVersion: srcVersion, done: [] };
  var doneMap = {};
  for (var i = 0; i < (ckpt.done || []).length; i++) doneMap[ckpt.done[i]] = true;

  var subjects = getSubjectsSortedBySize();
  var processed = 0;
  for (var s = 0; s < subjects.length; s++) {
    var sid = subjects[s].subject;
    if (doneMap[sid]) continue;
    if (Date.now() > deadline) break; // หมดงบเวลา — resume รอบหน้า

    try {
      generateQuestionRelationsForSubject(sid);
    } catch (err) {
      console.error("runQuestionRelationsBatch: subject " + sid + " failed: " + err.message);
    }
    ckpt.done.push(sid);
    doneMap[sid] = true;
    processed++;
    props.setProperty(RELATIONS_CHECKPOINT_KEY, JSON.stringify(ckpt)); // checkpoint หลังจบแต่ละวิชา
  }

  return { srcVersion: srcVersion, processedThisRun: processed, totalDone: ckpt.done.length, totalSubjects: subjects.length };
}

// อ่าน relations ของวิชาหนึ่งเป็น map { questionId: [{relatedId, score}, ...] } — chunked cache, อ่านอย่างเดียว
function getRelatedQuestionsData(subject) {
  var v = getVersionCached();
  var cleanFilter = subject ? String(subject).trim().toUpperCase() : "all";
  var cacheKey = "relations_" + v + "_" + cleanFilter;

  var cachedStr = getLargeCache(cacheKey);
  if (cachedStr != null) {
    return ContentService.createTextOutput(cachedStr).setMimeType(ContentService.MimeType.JSON);
  }

  var ss = SpreadsheetApp.openById(SHEET_ID);
  var relations = {};

  var sheet = ss.getSheetByName(QUESTION_RELATIONS_SHEET_NAME);
  if (sheet && sheet.getLastRow() > 1) {
    // ชีตไม่มีคอลัมน์ subject (สคีมา A-F ตายตัว) → กรองด้วยเซ็ต questionId ของวิชานี้
    var qJson = getQuestionsDataCached(subject, ss).getContent();
    var subjQuestions = [];
    try { subjQuestions = JSON.parse(qJson) || []; } catch (e) { subjQuestions = []; }
    var idSet = {};
    for (var i = 0; i < subjQuestions.length; i++) idSet[String(subjQuestions[i].questionId)] = true;

    var rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues(); // A,B,C = Question_ID, Related_ID, Score
    for (var r = 0; r < rows.length; r++) {
      var qid = String(rows[r][0]).trim();
      if (!idSet[qid]) continue;
      (relations[qid] || (relations[qid] = [])).push({
        relatedId: rows[r][1],
        score: Number(rows[r][2]) || 0
      });
    }
  }

  var payload = JSON.stringify({ result: 'success', relations: relations });
  putLargeCache(cacheKey, payload, 1800); // 30 นาที (คีย์ผูก version อยู่แล้ว)
  return ContentService.createTextOutput(payload).setMimeType(ContentService.MimeType.JSON);
}

// ติดตั้ง time-driven trigger รันทุกวัน ~ตี 3 (idempotent). *ไม่* เรียกอัตโนมัติตอนโหลด — เรียกเองภายหลังเมื่อพร้อม
function installRelationsTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'runQuestionRelationsBatch') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  ScriptApp.newTrigger('runQuestionRelationsBatch').timeBased().everyDays(1).atHour(3).create();
  return 'installed';
}

/* =========================================================================
   §1.8 — Knowledge-base corpus (Markdown textbook/lecture) — RAG grounding source #2
   ingest = แยกข้อความล้วน (token-free): split markdown บนหัวข้อ #/## → chunk ~500 คำ →
   dedup บน (Source, Heading) → เขียนแถว Status:auto. auth ทำที่ doPost (ingestKB) แล้ว.
   เสิร์ฟผ่าน getKB (public, chunked cache แบบเดียวกับ getRelatedQuestionsData).
   Category_ID (คอลัมน์ I) เว้นว่างไว้ใน v1 — §1.9 (step 5) จะ backfill + ใส่ใน payload ภายหลัง.
   ========================================================================= */

// idempotent: สร้างชีต KB_Chunks ถ้ายังไม่มี (mirror setupQuestionRelationsSheet)
// สคีมา A-H ตรงตาม §1.8 เป๊ะ + คอลัมน์ I (Category_ID) ต่อท้ายไว้เผื่อ §1.9 (ว่างใน v1)
function setupKBChunksSheet() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(KB_CHUNKS_SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(KB_CHUNKS_SHEET_NAME);

  var headers = ["Chunk_ID", "Subject_ID", "Source", "Heading", "Chunk_MD", "Embedding", "Status", "Updated_At", "Category_ID"];
  if (!sheet.getRange(1, 1).getValue()) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#e6f7ff");
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// แตก markdown เป็น chunk: split บนบรรทัดหัวข้อ (#..######) → ต่อ body → section ที่ยาวเกิน
// KB_CHUNK_MAX_WORDS คำ ตัดเป็น chunk ย่อย (heading เดิม). คืน [{heading, chunk_md}, ...]
// หมายเหตุ: นับคำด้วย whitespace → ไทยล้วน (ไม่มีช่องว่าง) จะไม่ถูกตัดย่อย = ยอมรับใน v1 token-free
function kbChunkMarkdown(markdown) {
  var text = (markdown == null ? "" : String(markdown));
  var lines = text.split(/\r?\n/);
  var sections = [];
  var curHeading = "";
  var curLines = [];
  function flush() {
    var body = curLines.join("\n").trim();
    if (body) sections.push({ heading: curHeading, body: body });
    curLines = [];
  }
  for (var i = 0; i < lines.length; i++) {
    var m = lines[i].match(/^#{1,6}\s+(.+?)\s*#*\s*$/); // บรรทัดหัวข้อ markdown
    if (m) {
      flush();
      curHeading = m[1].trim();
    } else {
      curLines.push(lines[i]);
    }
  }
  flush();

  var chunks = [];
  for (var s = 0; s < sections.length; s++) {
    var words = sections[s].body.split(/\s+/).filter(Boolean);
    if (words.length <= KB_CHUNK_MAX_WORDS) {
      chunks.push({ heading: sections[s].heading, chunk_md: sections[s].body });
    } else {
      for (var w = 0; w < words.length; w += KB_CHUNK_MAX_WORDS) {
        chunks.push({ heading: sections[s].heading, chunk_md: words.slice(w, w + KB_CHUNK_MAX_WORDS).join(" ") });
      }
    }
  }
  return chunks;
}

// ล้าง KB cache ของวิชาหนึ่ง (บังคับให้ getKB ดึงใหม่หลัง ingest) — mirror invalidateRelationsCache
function invalidateKBCache(subject) {
  var v = getVersionCached();
  var cleanFilter = subject ? String(subject).trim().toUpperCase() : "all";
  try { CacheService.getScriptCache().remove("kb_" + v + "_" + cleanFilter + "_chunks"); } catch (e) {}
}

// เขียน KB chunks จาก markdown ลงชีต KB_Chunks — dedup บน (Source, Heading), Status:auto.
// auth ถูกตรวจที่ doPost (ingestKB) ก่อนเรียกฟังก์ชันนี้เสมอ. คืน {result, written, skipped, total}
// §1.9: categoryId เป็น optional — ถ้าให้มา เขียนลงคอลัมน์ I (source-map routing), ถ้าไม่ให้ = "" (back-compat)
function ingestKBChunks(subject, source, markdown, categoryId) {
  var cleanSubject = String(subject || "").trim().toUpperCase();
  var cleanSource = String(source || "").trim();
  var cleanCategoryId = String(categoryId || "").trim();
  if (!cleanSubject || !cleanSource) {
    return { result: 'error', message: 'ต้องระบุ subject และ source' };
  }
  var chunks = kbChunkMarkdown(markdown);
  if (!chunks.length) {
    return { result: 'error', message: 'ไม่พบเนื้อหาใน markdown' };
  }

  var sheet = setupKBChunksSheet();
  // dedup key = (Source, Heading) จากแถวที่ "มีอยู่แล้ว" เท่านั้น — ไม่ mark ภายใน batch
  // เพราะ section เดียวที่ถูกตัดเป็นหลาย sub-chunk มี heading ซ้ำได้ (ต้องเขียนครบทุก sub-chunk)
  // → re-ingest source เดิมทีหลัง = heading อยู่ใน sheet แล้ว → skip ทั้งหมด (idempotent)
  var existing = sheet.getDataRange().getValues();
  var seen = {};
  for (var r = 1; r < existing.length; r++) {
    var key = String(existing[r][2]).trim() + "///" + String(existing[r][3]).trim();
    seen[key] = true;
  }

  var nowIso = new Date().toISOString();
  var newRows = [];
  var skipped = 0;
  for (var c = 0; c < chunks.length; c++) {
    var heading = chunks[c].heading || "";
    var dedupKey = cleanSource + "///" + String(heading).trim();
    if (seen[dedupKey]) { skipped++; continue; }
    var chunkId = "KB_" + cleanSubject + "_" + Utilities.getUuid().slice(0, 8);
    // A..I: Chunk_ID, Subject_ID, Source, Heading, Chunk_MD, Embedding(ว่าง v1), Status, Updated_At, Category_ID(§1.9)
    newRows.push([chunkId, cleanSubject, cleanSource, heading, chunks[c].chunk_md, "", "auto", nowIso, cleanCategoryId]);
  }

  if (newRows.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, 9).setValues(newRows);
  }
  invalidateKBCache(cleanSubject);
  return { result: 'success', written: newRows.length, skipped: skipped, total: chunks.length };
}

// อ่าน KB chunks ของวิชาหนึ่ง (public, chunked cache) — mirror getRelatedQuestionsData
// degrade เป็น chunks:[] เมื่อยังไม่มีชีต/ไม่มีแถวของวิชานี้
function getKBData(subject) {
  var v = getVersionCached();
  var cleanFilter = subject ? String(subject).trim().toUpperCase() : "all";
  var cacheKey = "kb_" + v + "_" + cleanFilter;

  var cachedStr = getLargeCache(cacheKey);
  if (cachedStr != null) {
    return ContentService.createTextOutput(cachedStr).setMimeType(ContentService.MimeType.JSON);
  }

  var ss = SpreadsheetApp.openById(SHEET_ID);
  var chunks = [];
  var sheet = ss.getSheetByName(KB_CHUNKS_SHEET_NAME);
  if (sheet && sheet.getLastRow() > 1) {
    var rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 9).getValues(); // A..I (§1.9: รวม Category_ID คอลัมน์ I)
    for (var r = 0; r < rows.length; r++) {
      if (!rows[r][0]) continue; // ข้ามแถวว่าง
      var rowSubj = String(rows[r][1]).trim().toUpperCase();
      if (cleanFilter !== "all" && rowSubj !== cleanFilter) continue;
      chunks.push({
        chunkId: rows[r][0],
        subject: rows[r][1],
        source: rows[r][2],
        heading: rows[r][3],
        chunk_md: rows[r][4],
        status: rows[r][6],
        categoryId: rows[r][8] // §1.9: คอลัมน์ I — client ใช้ join กับ q.category ทำ source-map routing (ว่างได้)
      });
    }
  }

  var payload = JSON.stringify({ result: 'success', chunks: chunks });
  putLargeCache(cacheKey, payload, 1800); // 30 นาที (คีย์ผูก version อยู่แล้ว)
  return ContentService.createTextOutput(payload).setMimeType(ContentService.MimeType.JSON);
}

/* =========================================================================
   FEATURE 2 — Glossary (root-word + Thai↔English, unified) §2.1–§2.6
   สองเส้นทาง:
   (1) tap/select miss-path = askGlossaryTerm (doPost standalone block, ด้านบน):
       rate-limit → dedup(cache, lock-free) → LLM lock-free → เขียน 1 แถวใต้ localized-15s lock.
   (2) batch idle-day = generateGlossaryForSubject / runGlossaryBatch (checkpointed) +
       runGlossaryBatchManual (admin-tier, ยิง LLM ใต้ admin lock = admin wall ที่แผนยอมรับ).
       trigger เว้นไว้ไม่ติดตั้ง (installGlossaryTrigger มีไว้แต่ไม่เรียก) — batch ชนกำแพงโทเคน/admin.
   เสิร์ฟผ่าน getGlossary (doGet, public, chunked cache — mirror getKBData).
   *** row shape (byte-identical: getGlossary payload = askGlossaryTerm return = client glossaryMap): ***
     { term_en, term_th, root_breakdown, root_parts(///-string), short_def_th, source_questionIds(///-string), status }
   ========================================================================= */

// normalize คีย์ศัพท์: ตัวเล็ก + trim + ตัดเครื่องหมายวรรคตอน/สัญลักษณ์หัวท้าย + ยุบช่องว่างซ้ำ
// *** ต้อง "เหมือน frontend normalizeGlossaryKey เป๊ะ" *** ไม่งั้นแถวที่ backend เขียนจะหาไม่เจอใน client map → ยิง LLM วนไม่จบ
function normalizeGlossaryTerm(s) {
  s = (s == null ? "" : String(s)).toLowerCase().trim();
  s = s.replace(/^[^a-z0-9฀-๿]+/, "").replace(/[^a-z0-9฀-๿]+$/, ""); // ตัด non-alnum (เก็บช่วงไทย) หัวท้าย
  return s.replace(/\s+/g, " ");
}

// กัน placeholder หลุดจาก template prompt ("prefix"/"root"/"suffix" ตรงตัว) — ทำ cluster §2.6 พัง
// แทนที่ด้วย "" (คงตำแหน่ง slot ไว้). ใช้ทั้งขาเขียน (parser) และขาเสิร์ฟ (rowToObj — self-heal แถวเก่าที่เขียนไปแล้ว)
function cleanGlossaryParts(parts) {
  return (parts || []).map(function (p) {
    p = String(p == null ? "" : p).trim();
    return /^(prefix|root|suffix)$/i.test(p) ? "" : p;
  });
}

// map แถว sheet (A..I) → object รูปแบบเดียวกับที่ frontend คาดหวัง (byte-identical keys). Subject_ID (col C) ไม่ส่งกลับ
function glossaryRowToObj(row) {
  return {
    term_en: row[0], term_th: row[1], root_breakdown: row[3],
    root_parts: cleanGlossaryParts(String(row[4] == null ? "" : row[4]).split("///")).join("///"), // ///-string — frontend split เอง (§2.6 clustering)
    short_def_th: row[5], source_questionIds: row[6], status: row[7]
  };
}

// idempotent: สร้างชีต Glossary ถ้ายังไม่มี (mirror setupKBChunksSheet) — สคีมา A..I ตาม §2.1 เป๊ะ
function setupGlossarySheet() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(GLOSSARY_SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(GLOSSARY_SHEET_NAME);
  var headers = ["Term_EN", "Term_TH", "Subject_ID", "Root_Breakdown", "Root_Parts", "Short_Def_TH", "Source_QuestionIds", "Status", "Updated_At"];
  if (!sheet.getRange(1, 1).getValue()) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#e6f7ff");
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// ล้าง glossary cache ของวิชา (บังคับ getGlossary ดึงใหม่หลังเขียน) — *** ต้องมี suffix "_chunks" ***
// เพราะเป็น marker key ของ putLargeCache (เหมือน invalidateKBCache) ลืมแล้ว remove จะ no-op → เสิร์ฟ stale
function invalidateGlossaryCache(subject) {
  var v = getVersionCached();
  var cleanFilter = subject ? String(subject).trim().toUpperCase() : "all";
  try { CacheService.getScriptCache().remove("glossary_" + v + "_" + cleanFilter + "_chunks"); } catch (e) {}
}

// อ่าน glossary ของวิชา (public, chunked cache) — mirror getKBData. degrade เป็น terms:[] เมื่อไม่มีชีต/แถว
function getGlossaryData(subject) {
  var v = getVersionCached();
  var cleanFilter = subject ? String(subject).trim().toUpperCase() : "all";
  var cacheKey = "glossary_" + v + "_" + cleanFilter;

  var cachedStr = getLargeCache(cacheKey);
  if (cachedStr != null) {
    return ContentService.createTextOutput(cachedStr).setMimeType(ContentService.MimeType.JSON);
  }

  var ss = SpreadsheetApp.openById(SHEET_ID);
  var terms = [];
  var sheet = ss.getSheetByName(GLOSSARY_SHEET_NAME);
  if (sheet && sheet.getLastRow() > 1) {
    var rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 9).getValues(); // A..I
    for (var r = 0; r < rows.length; r++) {
      if (!rows[r][0]) continue; // ข้ามแถวว่าง
      var rowSubj = String(rows[r][2]).trim().toUpperCase(); // col C = Subject_ID
      if (cleanFilter !== "all" && rowSubj !== cleanFilter) continue;
      terms.push(glossaryRowToObj(rows[r]));
    }
  }

  var payload = JSON.stringify({ result: 'success', terms: terms });
  putLargeCache(cacheKey, payload, 1800); // 30 นาที (คีย์ผูก version อยู่แล้ว)
  return ContentService.createTextOutput(payload).setMimeType(ContentService.MimeType.JSON);
}

// dedup แบบ lock-free ผ่าน getGlossary "cache" (ไม่ getDataRange ทุก tap — hot path) — คืน term object หรือ null
// match ทั้ง normalized Term_EN และ Term_TH (client map ก็คีย์สองทางนี้เหมือนกัน)
function glossaryLookupCached(subject, normKey) {
  if (!normKey) return null;
  var terms = [];
  try { terms = (JSON.parse(getGlossaryData(subject).getContent()) || {}).terms || []; } catch (e) { terms = []; }
  for (var i = 0; i < terms.length; i++) {
    if (normalizeGlossaryTerm(terms[i].term_en) === normKey ||
        normalizeGlossaryTerm(terms[i].term_th) === normKey) return terms[i];
  }
  return null;
}

// prompt §2.2 ปรับสำหรับ "คำเดียว" + ประโยคแวดล้อม (disambiguation) — สั่งให้คืน JSON object เดียว
function buildGlossaryTermPrompt(word, sentence) {
  return "สกัดความหมายศัพท์แพทย์ของคำที่กำหนด โดยใช้ประโยคแวดล้อมช่วย disambiguate. " +
    "ตอบเป็น JSON object เดียวเท่านั้น ห้ามมีข้อความอื่นหรือ markdown code fence:\n" +
    '{ "en": "คำอังกฤษ canonical (base form)", "th": "คำแปลไทย", ' +
    '"root": "การแตกรากศัพท์ เช่น hepato- (ตับ) + -megaly (โต)", ' +
    '"parts": ["peri-","cardi-","-itis"] (morphemes จริงของคำนั้น เรียงตำแหน่ง prefix/root/suffix, ช่องที่ไม่มีใส่ "" — ห้ามใส่คำว่า prefix/root/suffix ตรงๆ), ' +
    '"def_th": "นิยามสั้นๆ ไม่เกิน 1 บรรทัด" }\n' +
    'ถ้าไม่ใช่ศัพท์แพทย์จริง ให้ en เป็นคำเดิม, root=\"\" parts=[] และ def_th อธิบายความหมายทั่วไปสั้นๆ.\n' +
    'คำ: "' + String(word || "") + '"\n' +
    'ประโยคแวดล้อม: "' + String(sentence || "") + '"';
}

// parse คำตอบ LLM แบบ object เดียวแบบทน — ตัด code fence, คว้า {...} ก้อนแรก, coerce parts เป็น array. null เมื่อพัง
function parseGlossaryJson(raw) {
  if (!raw) return null;
  var text = String(raw).replace(/```json/gi, "").replace(/```/g, "").trim();
  var start = text.indexOf("{");
  var end = text.lastIndexOf("}");
  if (start < 0 || end <= start) return null;
  var obj;
  try { obj = JSON.parse(text.slice(start, end + 1)); } catch (e) { return null; }
  if (!obj || typeof obj !== "object") return null;
  var parts = obj.parts;
  if (!Array.isArray(parts)) parts = (parts == null || parts === "") ? [] : [String(parts)];
  return {
    en: String(obj.en || "").trim(),
    th: String(obj.th || "").trim(),
    root: String(obj.root || "").trim(),
    parts: cleanGlossaryParts(parts),
    def_th: String(obj.def_th || "").trim()
  };
}

// เขียน 1 แถว glossary จากผล LLM — เรียก "ใต้ localized-15s lock" เท่านั้น (ตรวจ dedup ซ้ำแบบ race-safe)
// re-read ชีตสด (ไม่ใช่ cache) เพื่อกัน 2 คำขอเขียนคำเดียวกันพร้อมกัน. คืน {term, cached}
function writeGlossaryRowLocked(subject, parsed, sourceQuestionIds) {
  var cleanSubject = String(subject || "").trim().toUpperCase();
  var enKey = normalizeGlossaryTerm(parsed.en);
  var sheet = setupGlossarySheet();
  var existing = sheet.getDataRange().getValues();
  for (var r = 1; r < existing.length; r++) {
    if (!existing[r][0]) continue;
    if (normalizeGlossaryTerm(existing[r][0]) === enKey) {
      return { term: glossaryRowToObj(existing[r]), cached: true }; // race: อีก request เขียนไปก่อนแล้ว
    }
  }
  var partsStr = (parsed.parts || []).join("///"); // §2.1 col E: prefix///root///suffix
  var nowIso = new Date().toISOString();
  // A..I: Term_EN, Term_TH, Subject_ID, Root_Breakdown, Root_Parts, Short_Def_TH, Source_QuestionIds, Status, Updated_At
  var row = [parsed.en, parsed.th, cleanSubject, parsed.root, partsStr, parsed.def_th, String(sourceQuestionIds || ""), "auto", nowIso];
  sheet.getRange(sheet.getLastRow() + 1, 1, 1, 9).setValues([row]);
  invalidateGlossaryCache(cleanSubject);
  return { term: glossaryRowToObj(row), cached: false };
}

/* ---- Batch generation (§2.2, idle-day) — ยิง LLM เป็นชุด ***spends tokens*** ---- */

// prompt §2.2 สกัดหลายศัพท์จากก้อนข้อความ → JSON array
function buildGlossaryBatchPrompt(text) {
  return "สกัดศัพท์แพทย์จากข้อความนี้ ตอบเป็น JSON array เท่านั้น (ห้ามมีข้อความอื่น/code fence). แต่ละตัว:\n" +
    '{ "en": "...", "th": "...", "root": "hepato- (ตับ) + -megaly (โต)", ' +
    '"parts": ["hepato-","","-megaly"], "def_th": "≤1 บรรทัด" }\n' +
    "เฉพาะศัพท์แพทย์จริง (กายวิภาค/พยาธิ/ยา/สรีระ). ข้ามคำทั่วไป.\n" +
    'ข้อความ: "' + String(text || "").slice(0, GLOSSARY_BATCH_CHARS) + '"';
}

// parse JSON array แบบทน — ตัด code fence, คว้า [...] ก้อนแรก, coerce แต่ละตัว
function parseGlossaryArray(raw) {
  if (!raw) return [];
  var text = String(raw).replace(/```json/gi, "").replace(/```/g, "").trim();
  var start = text.indexOf("[");
  var end = text.lastIndexOf("]");
  if (start < 0 || end <= start) return [];
  var arr;
  try { arr = JSON.parse(text.slice(start, end + 1)); } catch (e) { return []; }
  if (!Array.isArray(arr)) return [];
  return arr.map(function (o) {
    var parts = o && o.parts;
    if (!Array.isArray(parts)) parts = (parts == null || parts === "") ? [] : [String(parts)];
    return {
      en: String((o && o.en) || "").trim(), th: String((o && o.th) || "").trim(),
      root: String((o && o.root) || "").trim(),
      parts: cleanGlossaryParts(parts),
      def_th: String((o && o.def_th) || "").trim()
    };
  }).filter(function (o) { return o.en; });
}

// สกัด glossary จากทั้งวิชา (batch): อ่านข้อสอบ → strip รูป (reuse buildExplainPrompt logic) → ก้อนละ ~GLOSSARY_BATCH_CHARS
// → LLM หลายรอบ → dedup กับ Term_EN เดิม → เขียนแถวใหม่ Status:auto. คืนจำนวนแถวที่เขียน
function generateGlossaryForSubject(subjectId) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var qJson = getQuestionsDataCached(subjectId, ss).getContent();
  var questions = [];
  try { questions = JSON.parse(qJson) || []; } catch (e) { questions = []; }
  if (!questions.length) return 0;

  var sheet = setupGlossarySheet();
  var existing = sheet.getDataRange().getValues();
  var seen = {}; // normalized Term_EN ที่มีอยู่แล้ว → dedup ก่อนเขียน
  for (var r = 1; r < existing.length; r++) {
    if (existing[r][0]) seen[normalizeGlossaryTerm(existing[r][0])] = true;
  }

  // รวมข้อความทั้งวิชา (strip รูปเป็น [รูปภาพ] กันไป poison การสกัด) เป็นก้อน ~GLOSSARY_BATCH_CHARS ตัวอักษร
  var blobs = [];
  var cur = "";
  for (var i = 0; i < questions.length; i++) {
    var q = questions[i];
    var choiceArr = String(q.choices || "").split("///").map(function (c) {
      return (c.indexOf("http") === 0 || c.indexOf("<svg") === 0) ? "[รูปภาพ]" : c;
    });
    var piece = (q.problem || "") + " " + choiceArr.join(" ") + " " + (q.explain || "") + " " + (q.answer || "") + "\n";
    if (cur.length + piece.length > GLOSSARY_BATCH_CHARS && cur) { blobs.push(cur); cur = ""; }
    cur += piece;
  }
  if (cur) blobs.push(cur);

  var written = 0;
  var nowIso = new Date().toISOString();
  var cleanSubject = String(subjectId || "").trim().toUpperCase();
  for (var b = 0; b < blobs.length; b++) {
    var arr;
    try {
      var raw = executeChatbotQuery(buildGlossaryBatchPrompt(blobs[b]), GLOSSARY_MODEL, 1);
      arr = parseGlossaryArray(raw.content);
    } catch (e) { arr = []; }
    if (!arr || !arr.length) continue;
    var newRows = [];
    for (var a = 0; a < arr.length; a++) {
      var p = arr[a];
      var enKey = normalizeGlossaryTerm(p.en);
      if (!enKey || seen[enKey]) continue; // dedup ทั้งกับชีตเดิมและภายใน batch เดียวกัน
      seen[enKey] = true;
      newRows.push([p.en, p.th, cleanSubject, p.root, (p.parts || []).join("///"), p.def_th, "", "auto", nowIso]);
    }
    if (newRows.length) {
      sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, 9).setValues(newRows);
      written += newRows.length;
    }
  }
  if (written) invalidateGlossaryCache(cleanSubject);
  return written;
}

// nightly entry point (checkpointed) — mirror runQuestionRelationsBatch. ***ไม่มี trigger เรียกอัตโนมัติ***
function runGlossaryBatch() {
  var deadline = Date.now() + GLOSSARY_BATCH_BUDGET_MS;
  var srcVersion = getVersionCached();
  var props = PropertiesService.getScriptProperties();

  var ckpt = {};
  try { ckpt = JSON.parse(props.getProperty(GLOSSARY_CHECKPOINT_KEY) || "{}"); } catch (e) { ckpt = {}; }
  if (ckpt.srcVersion !== srcVersion) ckpt = { srcVersion: srcVersion, done: [] }; // version drift → เริ่มรอบใหม่
  var doneMap = {};
  for (var i = 0; i < (ckpt.done || []).length; i++) doneMap[ckpt.done[i]] = true;

  var subjects = getSubjectsSortedBySize();
  var processed = 0;
  for (var s = 0; s < subjects.length; s++) {
    var sid = subjects[s].subject;
    if (doneMap[sid]) continue;
    if (Date.now() > deadline) break; // หมดงบเวลา — resume รอบหน้า
    try { generateGlossaryForSubject(sid); }
    catch (err) { console.error("runGlossaryBatch: subject " + sid + " failed: " + err.message); }
    ckpt.done.push(sid); doneMap[sid] = true; processed++;
    props.setProperty(GLOSSARY_CHECKPOINT_KEY, JSON.stringify(ckpt));
  }
  return { srcVersion: srcVersion, processedThisRun: processed, totalDone: ckpt.done.length, totalSubjects: subjects.length };
}

// ติดตั้ง time-driven trigger รันทุกวัน ~ตี 4 (idempotent). *ไม่* เรียกอัตโนมัติ — เว้นไว้เหมือน relations
// (batch ยิง LLM ชนกำแพงโทเคน/admin; tap/select path เติม glossary เองแบบ lock-free ระหว่างนี้)
function installGlossaryTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'runGlossaryBatch') ScriptApp.deleteTrigger(triggers[i]);
  }
  ScriptApp.newTrigger('runGlossaryBatch').timeBased().everyDays(1).atHour(4).create();
  return 'installed';
}

/* =========================================================================
   FEATURE 3 — High-yield cram sheet (summary + mnemonics + keywords) §3.1–§3.6
   หน่วย = ต่อ "หมวด/หัวข้อ" (category) — ตรงกับขอบเขตที่นิสิตอ่านทวนก่อนสอบ 1 exam (§3.1)
   สองเส้นทาง:
   (1) lazy-generate miss-path = generateHighYield (doPost standalone block, ด้านบน):
       rate-limit → dedup(cache) → LLM lock-free (max_tokens สูง) → เขียน 1 แถวใต้ localized-15s lock.
   (2) batch idle-day = generateHighYieldForSubject / runHighYieldBatch (checkpointed) +
       runHighYieldBatchManual (admin-tier, ยิง LLM ใต้ admin lock = admin wall ที่แผนยอมรับ).
       trigger เว้นไว้ไม่ติดตั้ง (installHighYieldTrigger มีไว้แต่ไม่เรียก) — เหมือน glossary/relations.
   mnemonic vote = voteHighYieldMnemonic (doPost standalone block, ด้านบน) — อ่านคอลัมน์ H สดใต้ lock แล้ว mutate.
   เสิร์ฟผ่าน getHighYield (doGet, public, chunked cache — mirror getGlossary).
   คุณภาพ (§3.4): facts/keywords ควรมี second-model gate — แต่ glossary ยังไม่มี gate ให้ mirror → v1 เขียน Status:'auto'
     + frontend โชว์ badge "AI สร้าง"; mnemonics ใช้ community vote (คอลัมน์ H) แทนการ auto-judge.
   *** row shape (byte-identical: getHighYield payload = generateHighYield return = client highYieldCache): ***
     { category_id, summary_md, mnemonics[], keywords[], status, src_version, updated_at, mnemonic_votes{idx:net} }
   ========================================================================= */

// idempotent: สร้างชีต HighYield_Cache ถ้ายังไม่มี — สคีมา A..H ตาม §3.2 เป๊ะ
function setupHighYieldSheet() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(HIGHYIELD_SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(HIGHYIELD_SHEET_NAME);
  var headers = ["Category_ID", "Summary_MD", "Mnemonics", "Keywords", "Status", "Src_Version", "Updated_At", "Mnemonic_Votes"];
  if (!sheet.getRange(1, 1).getValue()) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#e6f7ff");
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// normalize คีย์หมวด — trim อย่างเดียว (category ID อาจ case-sensitive; ใช้ให้ตรงกันทุกที่: get/write/vote/invalidate)
function normalizeHighYieldCategory(c) {
  return String(c == null ? "" : c).trim();
}

// map แถว sheet (A..H) → object รูปแบบเดียวกับที่ frontend คาดหวัง (byte-identical keys)
function highYieldRowToObj(row) {
  var mnemonics = String(row[2] == null ? "" : row[2]).split("///").map(function (s) { return s.trim(); }).filter(Boolean);
  var keywords = String(row[3] == null ? "" : row[3]).split("///").map(function (s) { return s.trim(); }).filter(Boolean);
  var votes = {};
  try { votes = JSON.parse(String(row[7] || "{}")) || {}; } catch (e) { votes = {}; }
  return {
    category_id: row[0],
    summary_md: String(row[1] == null ? "" : row[1]),
    mnemonics: mnemonics,
    keywords: keywords,
    status: row[4],
    src_version: String(row[5] == null ? "" : row[5]),
    updated_at: row[6],
    mnemonic_votes: votes
  };
}

// ล้าง high-yield cache ของหมวด (บังคับ getHighYield ดึงใหม่หลังเขียน/โหวต) — *** ต้องมี suffix "_chunks" ***
// (marker key ของ putLargeCache เหมือน invalidateGlossaryCache; ลืมแล้ว remove จะ no-op → เสิร์ฟ stale)
function invalidateHighYieldCache(category) {
  var v = getVersionCached();
  var clean = normalizeHighYieldCategory(category) || "all";
  try { CacheService.getScriptCache().remove("highyield_" + v + "_" + clean + "_chunks"); } catch (e) {}
}

// อ่านชีทสรุปของหมวดหนึ่ง (public, chunked cache) — mirror getGlossaryData. degrade เป็น highyield:null เมื่อไม่มีชีต/แถว
function getHighYieldData(category) {
  var v = getVersionCached();
  var clean = normalizeHighYieldCategory(category);
  var cacheKey = "highyield_" + v + "_" + (clean || "all");

  var cachedStr = getLargeCache(cacheKey);
  if (cachedStr != null) {
    return ContentService.createTextOutput(cachedStr).setMimeType(ContentService.MimeType.JSON);
  }

  var ss = SpreadsheetApp.openById(SHEET_ID);
  var hy = null;
  var sheet = ss.getSheetByName(HIGHYIELD_SHEET_NAME);
  if (sheet && sheet.getLastRow() > 1 && clean) {
    var rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 8).getValues(); // A..H
    for (var r = 0; r < rows.length; r++) {
      if (!rows[r][0]) continue;
      if (normalizeHighYieldCategory(rows[r][0]) === clean) { hy = highYieldRowToObj(rows[r]); break; }
    }
  }

  var payload = JSON.stringify({ result: 'success', highyield: hy });
  putLargeCache(cacheKey, payload, 1800); // 30 นาที (คีย์ผูก version + invalidate ตอนเขียน/โหวต)
  return ContentService.createTextOutput(payload).setMimeType(ContentService.MimeType.JSON);
}

// dedup แบบ lock-free ผ่าน getHighYield "cache" (ไม่ getDataRange ทุกครั้ง) — คืน object หรือ null
function highYieldLookupCached(category) {
  var clean = normalizeHighYieldCategory(category);
  if (!clean) return null;
  try {
    var obj = JSON.parse(getHighYieldData(clean).getContent()) || {};
    return obj.highyield || null;
  } catch (e) { return null; }
}

// รวมข้อความข้อสอบของหมวดหนึ่งเป็น 1 ก้อน (สำหรับป้อน LLM) — subject resolve จาก categoryId ฝั่ง server (fallback = subjectHint)
// เลือกข้อ "ที่มีคำอธิบาย" ก่อน (มีค่ากับการสรุป) แล้วเติมข้อที่เหลือ; ตัดที่ MAX_QUESTIONS/MAX_CHARS; strip รูปเป็น [รูปภาพ]
function aggregateCategoryText(categoryId, subjectHint) {
  var clean = normalizeHighYieldCategory(categoryId);
  if (!clean) return "";
  var map = getCategoryToSubjectMapCached();
  var subject = map[clean] || (subjectHint ? String(subjectHint).trim().toUpperCase() : "");
  if (!subject) return "";

  var qJson = getQuestionsDataCached(subject).getContent();
  var questions = [];
  try { questions = JSON.parse(qJson) || []; } catch (e) { questions = []; }
  if (!questions.length) return "";

  // เก็บเฉพาะข้อของหมวดนี้
  var inCat = questions.filter(function (q) {
    var cats = Array.isArray(q.category) ? q.category : [q.category];
    return cats.some(function (c) { return normalizeHighYieldCategory(c) === clean; });
  });
  if (!inCat.length) return "";

  // จัดลำดับ: ข้อที่มี explain มาก่อน (มีค่ากับการสรุปมากกว่า) — §3.6 prioritize ones with explanations
  inCat.sort(function (a, b) {
    var ea = (a.explain && String(a.explain).trim()) ? 1 : 0;
    var eb = (b.explain && String(b.explain).trim()) ? 1 : 0;
    return eb - ea;
  });

  var pieces = [];
  var totalLen = 0;
  var used = 0;
  for (var i = 0; i < inCat.length && used < HIGHYIELD_MAX_QUESTIONS; i++) {
    var q = inCat[i];
    var choiceArr = String(q.choices || "").split("///").map(function (c) {
      return (c.indexOf("http") === 0 || c.indexOf("<svg") === 0) ? "[รูปภาพ]" : c; // reuse image-strip logic (Phase 0)
    });
    var piece = "โจทย์: " + (q.problem || "") +
      "\nตัวเลือก: " + choiceArr.join(" / ") +
      "\nเฉลย: " + (q.answer || "") +
      (q.explain ? ("\nคำอธิบาย: " + q.explain) : "") + "\n\n";
    if (totalLen + piece.length > HIGHYIELD_MAX_CHARS && pieces.length) break; // เต็มงบตัวอักษร
    pieces.push(piece);
    totalLen += piece.length;
    used++;
  }
  return pieces.join("").slice(0, HIGHYIELD_MAX_CHARS);
}

// prompt §3.3 — ตอบ JSON object เดียว, bound output (กัน truncation แม้ max_tokens สูงขึ้นแล้ว), ห้ามใส่ "///"
function buildHighYieldPrompt(text) {
  return "สร้างชีทสรุป high-yield สำหรับหัวข้อนี้ จากโจทย์+เฉลยด้านล่าง. " +
    "ตอบเป็น JSON object เดียวเท่านั้น ห้ามมีข้อความอื่นหรือ markdown code fence:\n" +
    '{ "summary_md": "ประเด็นสำคัญที่ต้องจำ เขียนเป็น markdown แบบหัวข้อย่อย (bullet) สั้นๆ ไม่เกิน ~12 ข้อ ครอบคลุมพอดี 1 หน้า", ' +
    '"mnemonics": ["ตัวช่วยจำที่ใช้ได้จริง", "..."], ' +
    '"keywords": ["ศัพท์ต้องรู้", "..."] }\n' +
    "อิงเฉพาะเนื้อหาที่ให้ ห้ามแต่งข้อมูลใหม่. mnemonics ต้องช่วยจำได้จริง ไม่มั่ว (สูงสุด " + HIGHYIELD_MAX_MNEMONICS + " ตัว). " +
    "keywords สูงสุด " + HIGHYIELD_MAX_KEYWORDS + " คำ. ตอบกระชับ อย่าใส่ตัวคั่น \"///\" ในข้อความใดๆ.\n" +
    "โจทย์+เฉลย:\n\"" + String(text || "").slice(0, HIGHYIELD_MAX_CHARS) + "\"";
}

// coerce list ของ string: trim, strip "///" (กันปน delimiter), ตัดว่าง, cap จำนวน
function sanitizeHighYieldList(arr, maxItems) {
  if (!Array.isArray(arr)) arr = (arr == null || arr === "") ? [] : [arr];
  var out = [];
  for (var i = 0; i < arr.length && out.length < maxItems; i++) {
    var s = String(arr[i] == null ? "" : arr[i]).replace(/\/\/\//g, "/").trim(); // strip literal ///
    if (s) out.push(s);
  }
  return out;
}

// parse คำตอบ LLM แบบ object เดียวแบบทน — ตัด code fence, คว้า {...} ก้อนแรก. null เมื่อพัง (frontend แจ้ง error+ลองใหม่)
function parseHighYieldJson(raw) {
  if (!raw) return null;
  var text = String(raw).replace(/```json/gi, "").replace(/```/g, "").trim();
  var start = text.indexOf("{");
  var end = text.lastIndexOf("}");
  if (start < 0 || end <= start) return null;
  var obj;
  try { obj = JSON.parse(text.slice(start, end + 1)); } catch (e) { return null; }
  if (!obj || typeof obj !== "object") return null;
  var summary = String(obj.summary_md == null ? "" : obj.summary_md).replace(/\/\/\//g, "/").trim();
  if (!summary) return null;
  return {
    summary_md: summary,
    mnemonics: sanitizeHighYieldList(obj.mnemonics, HIGHYIELD_MAX_MNEMONICS),
    keywords: sanitizeHighYieldList(obj.keywords, HIGHYIELD_MAX_KEYWORDS)
  };
}

// เขียน 1 แถว high-yield จากผล LLM — เรียก "ใต้ localized-15s lock" เท่านั้น. re-read ชีตสด (ไม่ใช่ cache) กัน 2 คำขอหมวดเดียวกันพร้อมกัน
// upsert: หมวดที่มีอยู่แล้ว → อัปเดต B..G และ ***รีเซ็ต Mnemonic_Votes เป็น {} (index mnemonic เปลี่ยน)***. คืน {row, cached}
function writeHighYieldRowLocked(category, parsed, srcVersion) {
  var clean = normalizeHighYieldCategory(category);
  var sheet = setupHighYieldSheet();
  var existing = sheet.getDataRange().getValues();
  var mnemStr = (parsed.mnemonics || []).join("///");
  var kwStr = (parsed.keywords || []).join("///");
  var nowIso = new Date().toISOString();

  for (var r = 1; r < existing.length; r++) {
    if (!existing[r][0]) continue;
    if (normalizeHighYieldCategory(existing[r][0]) === clean) {
      // race: อีก request (lazy path) เขียนไปก่อนแล้ว → คืนของเดิม ไม่ทับ (dedup lazy). batch เรียกผ่านทางอื่น
      return { row: highYieldRowToObj(existing[r]), cached: true };
    }
  }
  // A..H: Category_ID, Summary_MD, Mnemonics(///), Keywords(///), Status, Src_Version, Updated_At, Mnemonic_Votes(JSON)
  var row = [clean, parsed.summary_md, mnemStr, kwStr, "auto", String(srcVersion || ""), nowIso, "{}"];
  sheet.getRange(sheet.getLastRow() + 1, 1, 1, 8).setValues([row]);
  invalidateHighYieldCache(clean);
  return { row: highYieldRowToObj(row), cached: false };
}

// โหวต mnemonic — เรียก "ใต้ localized-15s lock" เท่านั้น. อ่านคอลัมน์ H "สดจากชีต" → mutate → เขียนกลับ → invalidate
// (advisor rule: ห้ามอ่านจาก cache ไม่งั้นโหวตพร้อมกันทับกัน). คืน {result, mnemonicIdx, netVotes}
function voteHighYieldMnemonicLocked(category, idx, delta) {
  var clean = normalizeHighYieldCategory(category);
  var sheet = setupHighYieldSheet();
  var values = sheet.getDataRange().getValues();
  for (var r = 1; r < values.length; r++) {
    if (!values[r][0]) continue;
    if (normalizeHighYieldCategory(values[r][0]) !== clean) continue;
    var mnemCount = String(values[r][2] == null ? "" : values[r][2]).split("///").map(function (s) { return s.trim(); }).filter(Boolean).length;
    if (idx >= mnemCount) return { result: 'error', message: 'ไม่พบ mnemonic ที่โหวต' };
    var votes = {};
    try { votes = JSON.parse(String(values[r][7] || "{}")) || {}; } catch (e) { votes = {}; }
    votes[idx] = (parseInt(votes[idx], 10) || 0) + delta;
    sheet.getRange(r + 1, 8).setValue(JSON.stringify(votes)); // คอลัมน์ H
    invalidateHighYieldCache(clean);
    return { result: 'success', mnemonicIdx: idx, netVotes: votes[idx] };
  }
  return { result: 'error', message: 'ยังไม่มีชีทสรุปของหัวข้อนี้' };
}

/* ---- Batch generation (§3.6, idle-day) — ยิง LLM เป็นชุด ***spends tokens*** ---- */

// สร้างชีทสรุปของ "ทุกหมวดในวิชาเดียว" (batch): หา category ทั้งหมดของวิชา → หมวดที่ยังไม่มีแถว → aggregate → LLM → เขียน
// คืนจำนวนหมวดที่เขียนใหม่. ***ยิง LLM หลาย call*** — ใช้กับวิชาเล็ก/ทดสอบ (admin) หรือ nightly checkpointed เท่านั้น
function generateHighYieldForSubject(subjectId) {
  var cleanSubject = String(subjectId || "").trim().toUpperCase();
  if (!cleanSubject) return 0;

  var map = getCategoryToSubjectMapCached();
  var cats = [];
  for (var catId in map) { if (map[catId] === cleanSubject) cats.push(catId); }
  if (!cats.length) return 0;

  var sheet = setupHighYieldSheet();
  var existing = sheet.getDataRange().getValues();
  var seen = {};
  for (var r = 1; r < existing.length; r++) {
    if (existing[r][0]) seen[normalizeHighYieldCategory(existing[r][0])] = true;
  }

  var srcVersion = getVersionCached();
  var written = 0;
  for (var c = 0; c < cats.length; c++) {
    var cat = normalizeHighYieldCategory(cats[c]);
    if (!cat || seen[cat]) continue; // dedup กับแถวที่มีอยู่แล้ว
    var text = aggregateCategoryText(cat, cleanSubject);
    if (!text) continue; // หมวดไม่มีข้อสอบ
    var parsed;
    try {
      var raw = executeChatbotQuery(buildHighYieldPrompt(text), HIGHYIELD_MODEL, 1, HIGHYIELD_MAX_TOKENS);
      parsed = parseHighYieldJson(raw.content);
    } catch (e) { parsed = null; }
    if (!parsed || !parsed.summary_md) continue;
    // เขียนตรง (batch อยู่ใต้ admin lock อยู่แล้ว — ไม่ต้อง localized lock ซ้อน)
    var row = [cat, parsed.summary_md, (parsed.mnemonics || []).join("///"), (parsed.keywords || []).join("///"),
      "auto", String(srcVersion || ""), new Date().toISOString(), "{}"];
    sheet.getRange(sheet.getLastRow() + 1, 1, 1, 8).setValues([row]);
    seen[cat] = true;
    invalidateHighYieldCache(cat);
    written++;
  }
  return written;
}

// nightly entry point (checkpointed) — mirror runGlossaryBatch. ***ไม่มี trigger เรียกอัตโนมัติ***
function runHighYieldBatch() {
  var deadline = Date.now() + HIGHYIELD_BATCH_BUDGET_MS;
  var srcVersion = getVersionCached();
  var props = PropertiesService.getScriptProperties();

  var ckpt = {};
  try { ckpt = JSON.parse(props.getProperty(HIGHYIELD_CHECKPOINT_KEY) || "{}"); } catch (e) { ckpt = {}; }
  if (ckpt.srcVersion !== srcVersion) ckpt = { srcVersion: srcVersion, done: [] };
  var doneMap = {};
  for (var i = 0; i < (ckpt.done || []).length; i++) doneMap[ckpt.done[i]] = true;

  var subjects = getSubjectsSortedBySize();
  var processed = 0;
  for (var s = 0; s < subjects.length; s++) {
    var sid = subjects[s].subject;
    if (doneMap[sid]) continue;
    if (Date.now() > deadline) break;
    try { generateHighYieldForSubject(sid); }
    catch (err) { console.error("runHighYieldBatch: subject " + sid + " failed: " + err.message); }
    ckpt.done.push(sid); doneMap[sid] = true; processed++;
    props.setProperty(HIGHYIELD_CHECKPOINT_KEY, JSON.stringify(ckpt));
  }
  return { srcVersion: srcVersion, processedThisRun: processed, totalDone: ckpt.done.length, totalSubjects: subjects.length };
}

// ติดตั้ง time-driven trigger รันทุกวัน ~ตี 4 (idempotent). *ไม่* เรียกอัตโนมัติ — เว้นไว้เหมือน glossary/relations
function installHighYieldTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'runHighYieldBatch') ScriptApp.deleteTrigger(triggers[i]);
  }
  ScriptApp.newTrigger('runHighYieldBatch').timeBased().everyDays(1).atHour(4).create();
  return 'installed';
}

/* =========================================================================
   FEATURE 6 — Frequently-tested keyword index (pre-exam review) §6.1–§6.4
   หน่วย = ต่อ "หมวด/หัวข้อ" (category). *** token-free 100% — ไม่ยิง LLM เลย ***
   term set = keywords ที่ HighYield สกัดไว้ (§3) ∪ Glossary terms ของวิชา (§2) — reuse 2 pass เดิม ไม่ pass ที่ 3
   นับความถี่ lexical (§6.2): EN → word-boundary regex, TH → substring (ไทยไม่มี word boundary = v1 characteristic)
   Freq = จำนวน "ข้อสอบที่ไม่ซ้ำ" ในหมวดที่ match; Source_QuestionIds = id ข้อเหล่านั้น; Source_KB_ChunkIds = KB ที่ match
   Fail_Weight (คอลัมน์ E) = 0 ใน v1 — ไม่มีข้อมูล per-answer ฝั่ง server (batchLog เก็บแค่ REPORT_OPEN/DOWNLOAD/IMG_ERROR
     ส่วน attemptCount/failCount นับใน IndexedDB ฝั่ง client เท่านั้น ไม่เคยส่งมา backend) → เก็บคอลัมน์ไว้ตาม schema เขียน 0
   เสิร์ฟผ่าน getKeywordIndex (doGet, public, chunked cache — คืน "list" ต่อหมวด, ไม่ใช่ object เดี่ยวแบบ HighYield)
   generation = admin เท่านั้น (runKeywordIndexBatchManual, §6.2 idle-day/admin) — ไม่มี public gen endpoint / ไม่ติดตั้ง trigger
   ========================================================================= */

// idempotent: สร้างชีต Keyword_Index ถ้ายังไม่มี — สคีมา A..I ตาม §6.1 เป๊ะ
function setupKeywordIndexSheet() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(KEYWORD_INDEX_SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(KEYWORD_INDEX_SHEET_NAME);
  var headers = ["Keyword_EN", "Keyword_TH", "Category_ID", "Freq", "Fail_Weight", "Source_QuestionIds", "Source_KB_ChunkIds", "Status", "Updated_At"];
  if (!sheet.getRange(1, 1).getValue()) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#e6f7ff");
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// map แถว sheet (A..I) → object รูปแบบเดียวกับที่ frontend คาดหวัง. source ids คงเป็น ///-string (frontend split เอง)
function keywordIndexRowToObj(row) {
  return {
    keyword_en: row[0],
    keyword_th: row[1],
    category_id: row[2],
    freq: parseInt(row[3], 10) || 0,
    fail_weight: parseInt(row[4], 10) || 0, // v1 = 0 เสมอ (ไม่มี per-answer data ฝั่ง server)
    source_questionIds: String(row[5] == null ? "" : row[5]),
    source_kb_chunkIds: String(row[6] == null ? "" : row[6]),
    status: row[7],
    updated_at: row[8]
  };
}

// ล้าง keyword-index cache ของหมวด (บังคับ getKeywordIndex ดึงใหม่หลังเขียน) — *** ต้องมี suffix "_chunks" ***
// (marker key ของ putLargeCache เหมือน invalidateHighYieldCache; ลืมแล้ว remove จะ no-op → เสิร์ฟ stale)
function invalidateKeywordIndexCache(category) {
  var v = getVersionCached();
  var clean = normalizeHighYieldCategory(category) || "all";
  try { CacheService.getScriptCache().remove("keywordindex_" + v + "_" + clean + "_chunks"); } catch (e) {}
}

// อ่านดัชนีคำสำคัญของหมวดหนึ่ง (public, chunked cache) — mirror getHighYieldData แต่คืน "list" ของทุก keyword ในหมวด
// จัดอันดับ Freq desc. degrade เป็น keywords:[] เมื่อไม่มีชีต/ไม่มีแถว (สะอาด — §6.3 review surface โชว์ "ยังไม่มี")
function getKeywordIndexData(category) {
  var v = getVersionCached();
  var clean = normalizeHighYieldCategory(category);
  var cacheKey = "keywordindex_" + v + "_" + (clean || "all");

  var cachedStr = getLargeCache(cacheKey);
  if (cachedStr != null) {
    return ContentService.createTextOutput(cachedStr).setMimeType(ContentService.MimeType.JSON);
  }

  var ss = SpreadsheetApp.openById(SHEET_ID);
  var list = [];
  var sheet = ss.getSheetByName(KEYWORD_INDEX_SHEET_NAME);
  if (sheet && sheet.getLastRow() > 1 && clean) {
    var rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 9).getValues(); // A..I
    for (var r = 0; r < rows.length; r++) {
      if (!rows[r][0] && !rows[r][1]) continue; // แถวว่าง (ไม่มีทั้ง EN/TH)
      if (normalizeHighYieldCategory(rows[r][2]) !== clean) continue; // คอลัมน์ C = Category_ID
      list.push(keywordIndexRowToObj(rows[r]));
    }
    list.sort(function (a, b) { return (b.freq || 0) - (a.freq || 0); }); // §6.3 rank Freq desc
  }

  var payload = JSON.stringify({ result: 'success', keywords: list });
  putLargeCache(cacheKey, payload, 1800); // 30 นาที (คีย์ผูก version + invalidate ตอนเขียน)
  return ContentService.createTextOutput(payload).setMimeType(ContentService.MimeType.JSON);
}

/* ---- PURE นับความถี่ (node-testable, ไม่พึ่ง GAS ใดๆ) — หัวใจของ Feature 6 ---- */

// สร้างข้อความค้นหา 1 ก้อนต่อข้อ (lowercased): problem + choices(strip รูป) + answer + explain
function keywordSearchText_(q) {
  var choiceArr = String(q.choices || "").split("///").map(function (c) {
    c = String(c);
    return (c.indexOf("http") === 0 || c.indexOf("<svg") === 0) ? "" : c; // ตัดรูป (URL/SVG) กัน match ขยะ
  });
  return [String(q.problem || ""), choiceArr.join(" "), String(q.answer || ""), String(q.explain || "")]
    .join(" ").toLowerCase();
}

// สร้าง matcher(text)->bool จาก term {en, th}: match ถ้า "surface form ใดฟอร์มหนึ่ง" ตรง
//   - ฟอร์มที่มีอักษรละติน → word-boundary regex (กัน substring inflation, §6.4) — escape regex metachars
//   - ฟอร์มไทยล้วน → substring (ไทยไม่มี word boundary) — ต้องยาว ≥ KEYWORD_MIN_LEN_TH กัน match มั่ว
// คืน null ถ้าไม่มี surface form ใดผ่าน min-length (term ถูกข้าม)
function buildKeywordMatcher_(term) {
  var forms = [];
  [term.en, term.th].forEach(function (raw) {
    var f = String(raw == null ? "" : raw).trim();
    if (!f) return;
    var hasLatin = /[A-Za-z]/.test(f);
    var minLen = hasLatin ? KEYWORD_MIN_LEN_EN : KEYWORD_MIN_LEN_TH;
    if (f.length < minLen) return;
    if (hasLatin) {
      var esc = f.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
      // word boundary แบบคุมเอง: ขอบเป็นต้น/ท้ายข้อความ หรืออักขระที่ไม่ใช่ alnum ละติน (ไทยถือเป็นขอบ = ยอมรับใน v1)
      forms.push({ re: new RegExp("(^|[^a-z0-9])" + esc + "($|[^a-z0-9])", "i") });
    } else {
      forms.push({ sub: f.toLowerCase() });
    }
  });
  if (!forms.length) return null;
  return function (text) {
    for (var i = 0; i < forms.length; i++) {
      if (forms[i].sub != null) { if (text.indexOf(forms[i].sub) >= 0) return true; }
      else if (forms[i].re.test(text)) return true;
    }
    return false;
  };
}

// PURE: (terms, questions, kbChunks) -> rows. terms=[{en,th,key}] (คัด/รวมมาแล้ว), questions/kbChunks = เฉพาะที่จะสแกน
// Freq = จำนวน "ข้อที่ไม่ซ้ำ" ที่ match (นับข้อละครั้งต่อ term); KB match ไม่นับ Freq แต่เก็บเป็น Source_KB_ChunkIds
// คืนเฉพาะ term ที่ Freq>=1 (§6.2). ไม่พึ่ง SpreadsheetApp → unit-test ใน node ได้ (จุดที่ verify list ของ spec ไม่ครอบคลุม)
function computeKeywordIndexRows(terms, questions, kbChunks) {
  var qTexts = (questions || []).map(function (q) { return { id: q.questionId, text: keywordSearchText_(q) }; });
  var kbTexts = (kbChunks || []).map(function (c) {
    return { id: c.chunkId, text: String((c.chunk_md || "") + " " + (c.heading || "")).toLowerCase() };
  });

  var out = [];
  for (var t = 0; t < (terms || []).length; t++) {
    var term = terms[t];
    var matcher = buildKeywordMatcher_(term);
    if (!matcher) continue; // ไม่มี surface form ผ่าน min-length

    var qids = [];
    for (var i = 0; i < qTexts.length; i++) {
      if (matcher(qTexts[i].text)) qids.push(String(qTexts[i].id));
    }
    if (!qids.length) continue; // §6.2 เขียนเฉพาะ Freq>=1

    var cids = [];
    for (var j = 0; j < kbTexts.length; j++) {
      if (matcher(kbTexts[j].text)) cids.push(String(kbTexts[j].id));
    }
    out.push({
      keyword_en: term.en || "",
      keyword_th: term.th || "",
      freq: qids.length,
      source_questionIds: qids,
      source_kb_chunkIds: cids
    });
  }
  return out;
}

/* ---- GAS wrapper: ประกอบ term set จาก cached getters (ไม่ getDataRange) แล้วเรียก pure fn + upsert ---- */

// สร้าง keyword index ของหมวดเดียว (token-free). คืนจำนวนแถวที่เขียน/อัปเดต
// term set: (a) HighYield keywords ของหมวด (highYieldLookupCached — cache) ∪ (b) Glossary terms ของวิชา (getGlossaryData — cache)
// dedup key = normalizeGlossaryTerm (เหมือน Glossary Term_EN); glossary เติม TH ให้ keyword ที่ตรงกัน (match ได้ทั้ง 2 ฟอร์ม)
function generateKeywordIndexForCategory(categoryId) {
  var clean = normalizeHighYieldCategory(categoryId);
  if (!clean) return 0;
  var map = getCategoryToSubjectMapCached();
  var subject = map[clean] || "";
  if (!subject) return 0;

  // (a) HighYield keywords ของหมวดนี้ (cached lookup — ไม่อ่านชีตตรง)
  var termMap = {}; // normKey -> {en, th, key}
  var hy = highYieldLookupCached(clean);
  if (hy && hy.keywords) {
    hy.keywords.forEach(function (kw) {
      var key = normalizeGlossaryTerm(kw);
      if (!key) return;
      if (!termMap[key]) termMap[key] = { en: kw, th: "", key: key };
    });
  }
  // (b) Glossary terms ของวิชา (cached getter — subject-filtered แล้ว)
  var gloss = [];
  try { gloss = (JSON.parse(getGlossaryData(subject).getContent()) || {}).terms || []; } catch (e) { gloss = []; }
  gloss.forEach(function (term) {
    var en = term.term_en, th = term.term_th;
    var key = normalizeGlossaryTerm(en) || normalizeGlossaryTerm(th);
    if (!key) return;
    if (termMap[key]) {
      if (!termMap[key].th && th) termMap[key].th = th;                                  // glossary เติม TH ให้ keyword EN
      if ((!termMap[key].en || !/[A-Za-z]/.test(termMap[key].en)) && en) termMap[key].en = en; // เติม EN canonical ถ้าเดิมเป็นไทย
    } else {
      termMap[key] = { en: en || "", th: th || "", key: key };
    }
  });
  var terms = [];
  for (var k in termMap) terms.push(termMap[k]);
  if (!terms.length) return 0;

  // questions ของหมวด (cached) — กรองด้วย category เดียวกับ aggregateCategoryText
  var questions = [];
  try { questions = JSON.parse(getQuestionsDataCached(subject).getContent()) || []; } catch (e) { questions = []; }
  var inCat = questions.filter(function (q) {
    var cats = Array.isArray(q.category) ? q.category : [q.category];
    return cats.some(function (c) { return normalizeHighYieldCategory(c) === clean; });
  });

  // KB chunks ของวิชา (cached; ว่างใน prod) — ถ้า chunk มี categoryId ให้กรองเฉพาะหมวดนี้ (§1.9), ไม่มี = subject-wide
  var kbChunks = [];
  try { kbChunks = (JSON.parse(getKBData(subject).getContent()) || {}).chunks || []; } catch (e) { kbChunks = []; }
  var kbInCat = kbChunks.filter(function (c) {
    return !c.categoryId || normalizeHighYieldCategory(c.categoryId) === clean;
  });

  var rows = computeKeywordIndexRows(terms, inCat, kbInCat);
  if (!rows.length) return 0;
  return writeKeywordIndexRows(clean, rows);
}

// upsert แถว keyword index ของหมวด — อ่านชีต "สด" ครั้งเดียว (เฉพาะ Keyword_Index; getter อื่นใช้ cache หมด)
// dedup/upsert key = (normalizeGlossaryTerm(Keyword_EN), Category_ID). preserve Status='reviewed' ของเดิมถ้ามี. คืนจำนวนแถว
function writeKeywordIndexRows(categoryClean, rows) {
  var sheet = setupKeywordIndexSheet();
  var existing = sheet.getDataRange().getValues();
  var idxByKey = {}; // normKey -> sheet row number (1-based)
  for (var r = 1; r < existing.length; r++) {
    if (normalizeHighYieldCategory(existing[r][2]) !== categoryClean) continue;
    var ek = normalizeGlossaryTerm(existing[r][0]) || normalizeGlossaryTerm(existing[r][1]);
    if (ek) idxByKey[ek] = r + 1;
  }

  var nowIso = new Date().toISOString();
  var appended = [];
  var n = 0;
  for (var i = 0; i < rows.length; i++) {
    var row = rows[i];
    var key = normalizeGlossaryTerm(row.keyword_en) || normalizeGlossaryTerm(row.keyword_th);
    var status = "auto";
    if (key && idxByKey[key]) {
      var prevStatus = existing[idxByKey[key] - 1][7];
      if (String(prevStatus) === "reviewed") status = "reviewed"; // คงสถานะที่มนุษย์ review แล้ว
    }
    // A..I: Keyword_EN, Keyword_TH, Category_ID, Freq, Fail_Weight(=0 v1), Source_QuestionIds, Source_KB_ChunkIds, Status, Updated_At
    var vals = [
      row.keyword_en, row.keyword_th, categoryClean,
      row.freq, 0,
      (row.source_questionIds || []).join("///"),
      (row.source_kb_chunkIds || []).join("///"),
      status, nowIso
    ];
    if (key && idxByKey[key]) {
      sheet.getRange(idxByKey[key], 1, 1, 9).setValues([vals]);
    } else {
      appended.push(vals);
    }
    n++;
  }
  if (appended.length) {
    sheet.getRange(sheet.getLastRow() + 1, 1, appended.length, 9).setValues(appended);
  }
  invalidateKeywordIndexCache(categoryClean);
  return n;
}

// สร้าง keyword index ของ "ทุกหมวดในวิชาเดียว" (§6.2 idle-day/admin) — token-free จึงไม่ต้อง checkpoint
// ⚠️ big-subject cost: วน (ทุกหมวด × ทุก term × ทุกข้อ) "ใต้ admin lock 25s" — วิชาใหญ่ (CVS 2516 ข้อ) อาจถือ lock นาน
//    ยังไม่มี trigger nightly (เว้นไว้เหมือน glossary/highyield); เรียกผ่าน runKeywordIndexBatchManual (admin) เท่านั้น
function generateKeywordIndexForSubject(subjectId) {
  var cleanSubject = String(subjectId || "").trim().toUpperCase();
  if (!cleanSubject) return { categories: 0, rows: 0 };

  var map = getCategoryToSubjectMapCached();
  var cats = [];
  for (var catId in map) { if (map[catId] === cleanSubject) cats.push(catId); }

  var totalRows = 0, processed = 0;
  for (var c = 0; c < cats.length; c++) {
    try { totalRows += generateKeywordIndexForCategory(cats[c]); processed++; }
    catch (err) { console.error("generateKeywordIndexForCategory " + cats[c] + " failed: " + err.message); }
  }
  return { categories: processed, rows: totalRows };
}