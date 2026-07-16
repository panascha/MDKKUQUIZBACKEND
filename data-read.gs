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
    serverTime: Date.now(), // seed lastSyncTs ฝั่ง client สำหรับ getAdminSync delta
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

// getAdminSync — combined delta-sync endpoint สำหรับแดชบอร์ดแอดมิน (ยิงจาก doPost lock-free tier, auth แล้ว)
// clientVer ตรงกับเวอร์ชันปัจจุบัน ⇒ NOT_MODIFIED; ไม่ตรง ⇒ ส่ง small slices ทั้งก้อน + question delta (ไม่ส่ง questions เต็ม)
function getAdminSyncData(clientVer, sinceStr) {
  // Stamp เวลา "ก่อน" อ่านทุกอย่าง — write ที่ landing ระหว่างประมวลผลจะถูกเก็บใน delta รอบถัดไปเสมอ (overlap = idempotent)
  var syncTime = new Date().getTime();
  var v = getVersionCached();
  if (clientVer && String(clientVer) === String(v)) {
    return ContentService.createTextOutput(JSON.stringify({
      status: 'NOT_MODIFIED', v: v, serverTime: syncTime
    })).setMimeType(ContentService.MimeType.JSON);
  }

  // Question delta — เรียกแบบไม่ filter subject เพื่อให้ "id อยู่ใน changedIds แต่ไม่มีแถว" = ถูกลบ เสมอ
  var delta = JSON.parse(getChangedSinceTimestamp(sinceStr, '').getContent());

  // Small slices (ทุกอย่างยกเว้น questions ~1.6MB raw) — cache ผูกเวอร์ชัน TTL 1800 เหมือน getAllData
  var smallKey = "admin_sync_small_" + v;
  var smallStr = getLargeCache(smallKey);
  var small;
  if (smallStr) {
    small = JSON.parse(smallStr);
  } else {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var adminsSafe = getSheetDataJSON('Admins', ss).map(function (admin) {
      var safeAdmin = {};
      for (var key in admin) {
        if (key !== 'Password') safeAdmin[key] = admin[key];
      }
      return safeAdmin;
    });
    small = {
      structure: getSheetDataJSON('Structure', ss),
      category: getSheetDataJSON('Category', ss),
      report: getSheetDataJSON('Report', ss),
      votes: getSheetDataJSON('Votes', ss),
      logs: getLogsTailJSON(ss, 300),
      admins: adminsSafe,
      announcements: getSheetDataJSON('Announcements', ss)
    };
    putLargeCache(smallKey, JSON.stringify(small), 1800);
  }

  return ContentService.createTextOutput(JSON.stringify({
    result: 'success',
    v: v,
    serverTime: syncTime,
    structure: small.structure,
    category: small.category,
    report: small.report,
    votes: small.votes,
    logs: small.logs,
    admins: small.admins,
    announcements: small.announcements,
    changedQuestions: delta.changed,
    changedIds: delta.changedIds
  })).setMimeType(ContentService.MimeType.JSON);
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
        // targetId อาจเป็น comma-joined หลาย qid (IMPORT/BULK_CATEGORIZE) — split ให้เป็นรายข้อ
        targetId.split(",").forEach(function (tid) {
          tid = tid.trim();
          if (tid) changedIds[tid] = true;
        });
      }
    }
  }

  var changedIdKeys = Object.keys(changedIds);

  if (changedIdKeys.length === 0) {
    return ContentService.createTextOutput(JSON.stringify({
      changed: [],
      changedIds: [],
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
    // qid ทุกตัวที่ log ระบุว่าเปลี่ยน — id ที่อยู่ใน changedIds แต่ไม่มีแถวใน changed ⇒ ถูกลบ (client drop ได้)
    changedIds: changedIdKeys,
    serverTime: new Date().getTime(),
    count: changedQuestions.length
  })).setMimeType(ContentService.MimeType.JSON);
}

