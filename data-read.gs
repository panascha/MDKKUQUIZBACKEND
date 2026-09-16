function getAdminsList() {
    return getSheetDataJSON('Admins');
}

function getAllDataForAdmin(startTime) {
  startTime = startTime || Date.now();
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

  assertNotTimedOut_(startTime, 'getAllDataForAdmin:admins');
  getOrCreateAnnouncementsSheet(ss); // Ensure sheet exists

  assertNotTimedOut_(startTime, 'getAllDataForAdmin:questions');
  // getQuestionsArray คืน array ตรง ไม่ต้องผ่าน stringify→parse ของ getQuestionsData (เดิม materialize คำถาม 26MB ซ้ำ 2 รอบก่อนถูก stringify รอบสุดท้ายด้านล่าง)
  var questionsArr = getQuestionsArray('', ss, startTime);

  assertNotTimedOut_(startTime, 'getAllDataForAdmin:structure_category');
  var structureArr = getSheetDataJSON('Structure', ss);
  var categoryArr = getSheetDataJSON('Category', ss);

  assertNotTimedOut_(startTime, 'getAllDataForAdmin:report_votes');
  var reportArr = getSheetDataJSON('Report', ss);
  var votesArr = getSheetDataJSON('Votes', ss);

  assertNotTimedOut_(startTime, 'getAllDataForAdmin:logs');
  var logsArr = getLogsTailJSON(ss, 300); // จำกัดเฉพาะ 300 แถวล่าสุด (Logs โตไม่จำกัด) — โหลดเต็มผ่าน action=getLogsPage

  var data = {
    v: getVersionCached(), // แทรกเวอร์ชันปัจจุบันเพื่อให้ฝั่งไคลเอนต์ใช้ซิงค์ในรอบเดี่ยวได้โดยไม่ต้องยิง checkVersion แยก
    serverTime: Date.now(), // seed lastSyncTs ฝั่ง client สำหรับ getAdminSync delta
    questions: questionsArr,
    structure: structureArr,
    category: categoryArr,
    report: reportArr,
    votes: votesArr,
    logs: logsArr,
    admins: adminsSafe,
    announcements: getSheetDataJSON('Announcements', ss)
  };
  assertNotTimedOut_(startTime, 'getAllDataForAdmin:before_stringify');
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
function getLogsPageData(offsetStr, limitStr, startTime) {
    if (startTime) assertNotTimedOut_(startTime, 'getLogsPageData');
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

function getPendingVotesData(qid, startTime) {
  var v = getVotesVersionCached();
  var cacheKey = "pending_votes_" + v + "_" + qid;
  var cached = getLargeCache(cacheKey);
  if (cached != null) {
    return ContentService.createTextOutput(cached).setMimeType(ContentService.MimeType.JSON);
  }

  if (startTime) assertNotTimedOut_(startTime, 'getPendingVotesData');
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

function getPendingReportsData(qid, startTime) {
  var v = getVotesVersionCached();
  var cacheKey = "pending_reports_" + v + "_" + qid;
  var cached = getLargeCache(cacheKey);
  if (cached != null) {
    return ContentService.createTextOutput(cached).setMimeType(ContentService.MimeType.JSON);
  }

  if (startTime) assertNotTimedOut_(startTime, 'getPendingReportsData');
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
function getPendingVotesReportsData(subjectParam, startTime) {
  var v = getVotesVersionCached();
  var cleanFilter = subjectParam ? String(subjectParam).trim().toUpperCase() : "all";
  var cacheKey = "pending_vr_" + v + "_" + cleanFilter;
  var cached = getLargeCache(cacheKey);
  if (cached != null) {
    return ContentService.createTextOutput(cached).setMimeType(ContentService.MimeType.JSON);
  }

  if (startTime) assertNotTimedOut_(startTime, 'getPendingVotesReportsData:start');
  var ss = SpreadsheetApp.openById(SHEET_ID);

  // สร้างเซ็ตของ qid ที่อยู่ในวิชานี้ (ผ่าน category->subject map + คำถามที่แคชไว้)
  // ใช้ตัวคำถามเป็นเกณฑ์เพื่อความถูกต้อง โดยไม่ขึ้นกับความกำกวมของคอลัมน์ subject ในชีต Votes/Report
  var subjectQids = null; // null = รับทุก qid (กรณี subject = all)
  if (cleanFilter !== "all") {
    var catToSubj = getCategoryToSubjectMapCached(ss, startTime);
    var qData = getAllQuestionsCached(ss, startTime);
    subjectQids = {};
    for (var qi = 0; qi < qData.length; qi++) {
      if (qi % 500 === 0 && startTime) assertNotTimedOut_(startTime, 'getPendingVotesReportsData:qData_loop');
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

  if (startTime) assertNotTimedOut_(startTime, 'getPendingVotesReportsData:before_votes');
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

  if (startTime) assertNotTimedOut_(startTime, 'getPendingVotesReportsData:before_reports');
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

function getStructureData(filterSubject, startTime) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var cleanFilter = filterSubject ? String(filterSubject).trim().toUpperCase() : "";

  var rows = getStructureSheetDataCached(ss, startTime);
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

  var catRows = getCategorySheetDataCached(ss, startTime);
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

// คืน array ของคำถามตรง (ไม่ผ่าน ContentService) — ให้ caller ในกระบวนการเดียวกัน (เช่น getAllDataForAdmin)
// ใส่ลง object รวมแล้ว stringify ครั้งเดียวตอนท้าย แทนที่จะ stringify ที่นี่แล้วต้อง parse กลับซ้ำ
function getQuestionsArray(filterSubject, ss, startTime) {
  if (!ss) ss = SpreadsheetApp.openById(SHEET_ID);
  var cleanFilter = filterSubject ? String(filterSubject).trim().toUpperCase() : "";

  var categoryToSubjectMap = getCategoryToSubjectMapCached(ss, startTime);

  var qData = getAllQuestionsCached(ss, startTime);
  if (qData.length === 0) return [];

  // นอก .map callback เสมอ — ข้างในมี catch(err) ของตัวเองที่จะกลืน throw ของ assertNotTimedOut_ ถ้าเช็คในนั้น
  if (startTime) assertNotTimedOut_(startTime, 'getQuestionsArray:before_map');

  return qData.map(function (row) {
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
}

function getQuestionsData(filterSubject, ss, startTime) {
  var questions = getQuestionsArray(filterSubject, ss, startTime);
  if (startTime) assertNotTimedOut_(startTime, 'getQuestionsData:end');
  return ContentService.createTextOutput(JSON.stringify(questions)).setMimeType(ContentService.MimeType.JSON);
}

// Prior Year Audit — เตือนแอดมินถ้าวิชาใดยังไม่มีข้อสอบ "รุ่นก่อนหน้า" + เทียบกลุ่มข้อสอบ (MCQ1/MCQ2/FMT ฯลฯ) ระหว่างรุ่นล่าสุดกับรุ่นก่อนหน้า
// "รุ่น" (year) ในที่นี้คือเลขรุ่นสอบที่ฝังอยู่ใน categoryId (เช่น CVS_52FMT1 = รุ่น 52) — ไม่ใช่คอลัมน์ year ของ Structure sheet (นั่นคือชั้นปีหลักสูตร 1-6)
// เฉพาะ Structure/Category sheet เท่านั้น ไม่แตะ Questions sheet เลย — ข้อมูลเล็ก ไม่มีความเสี่ยง timeout
function getPriorYearAuditData(startTime) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var catRows = getCategorySheetDataCached(ss, startTime);
  var structRows = getStructureSheetDataCached(ss, startTime);

  var subjectNameById = {};
  for (var s = 1; s < structRows.length; s++) {
    var sid = String(structRows[s][1] || '').trim();
    if (sid && !subjectNameById[sid]) subjectNameById[sid] = String(structRows[s][2] || '').trim();
  }

  // subjectId -> { year(number) -> { groupKey: true } }
  var bySubject = {};
  // กลุ่มข้อสอบ = ตัวเลขรุ่น 2 หลัก + ตัวอักษรประเภท + เลขกลุ่ม(ถ้ามี) — ยอมรับส่วนขยายท้าย เช่น "_Alltopics"/"_AnatomyPhysiology" (ของจริงในชีท)
  // แต่ตัด junk suffix ที่ไม่ใช่หัวข้อจริง (_Extracted, _Modified, by AI) ออกก่อนเช็ค — ไม่งั้นจะกลายเป็นกลุ่มปลอมซ้ำ
  // เจตนาไม่รับ "xx" (รหัสปีที่ไม่ทราบ เช่น COMMED2_xxMCQ1) — audit นี้เทียบรุ่น N กับ N-1
  // ข้อสอบที่ไม่รู้ปีจึงเทียบไม่ได้ ต้องข้ามไป ห้ามแก้เป็น (\d{2}|xx)
  // เพราะ parseInt("xx") = NaN แล้วจะเกิด bucket bySubject[subj][NaN] ทำให้ผลเทียบเพี้ยน
  var groupRe = /^(\d{2})_?([A-Za-z]+)(\d*)/;
  var junkSuffixRe = /_Extracted|_Modified|by AI/i;

  for (var i = 1; i < catRows.length; i++) {
    if (i % 500 === 0 && startTime) assertNotTimedOut_(startTime, 'getPriorYearAuditData:cat_loop');
    var categoryId = String(catRows[i][0] || '').trim();
    var subjectRef = String(catRows[i][1] || '').trim();
    if (!categoryId || !subjectRef) continue;

    var prefix = subjectRef + '_';
    if (categoryId.indexOf(prefix) !== 0) continue;

    var rest = categoryId.slice(prefix.length);
    if (junkSuffixRe.test(rest)) continue;

    var m = groupRe.exec(rest);
    if (!m) continue;

    var year = parseInt(m[1], 10);
    var groupKey = (m[2] + m[3]).toUpperCase();

    if (!bySubject[subjectRef]) bySubject[subjectRef] = {};
    if (!bySubject[subjectRef][year]) bySubject[subjectRef][year] = {};
    bySubject[subjectRef][year][groupKey] = true;
  }

  var subjects = [];
  var missingPriorCount = 0;
  var mismatchCount = 0;

  Object.keys(bySubject).sort().forEach(function (subjectId) {
    var yearMap = bySubject[subjectId];
    var years = Object.keys(yearMap).map(Number).sort(function (a, b) { return b - a; });
    if (years.length === 0) return;

    var latestYear = years[0];
    var priorYear = latestYear - 1;
    var priorYearExists = yearMap.hasOwnProperty(priorYear);

    var latestGroups = Object.keys(yearMap[latestYear]).sort();
    var priorGroups = priorYearExists ? Object.keys(yearMap[priorYear]).sort() : [];

    var structuralDiff = null;
    if (priorYearExists) {
      var onlyInLatest = latestGroups.filter(function (g) { return priorGroups.indexOf(g) === -1; });
      var onlyInPrior = priorGroups.filter(function (g) { return latestGroups.indexOf(g) === -1; });
      var matched = latestGroups.filter(function (g) { return priorGroups.indexOf(g) !== -1; });
      structuralDiff = { onlyInLatest: onlyInLatest, onlyInPrior: onlyInPrior, matched: matched };
      if (onlyInLatest.length > 0 || onlyInPrior.length > 0) mismatchCount++;
    } else {
      missingPriorCount++;
    }

    subjects.push({
      subjectId: subjectId,
      subjectName: subjectNameById[subjectId] || subjectId,
      years: years,
      latestYear: latestYear,
      priorYear: priorYear,
      priorYearExists: priorYearExists,
      latestGroups: latestGroups,
      priorGroups: priorGroups,
      structuralDiff: structuralDiff
    });
  });

  return ContentService.createTextOutput(JSON.stringify({
    status: 'success',
    serverTime: Date.now(),
    summary: {
      totalSubjects: subjects.length,
      missingPriorCount: missingPriorCount,
      mismatchCount: mismatchCount
    },
    subjects: subjects
  })).setMimeType(ContentService.MimeType.JSON);
}

// คืนแถว Questions เฉพาะ qid ที่เปลี่ยน
// เหตุผล: updateVersion() เปลี่ยน v ทุกครั้งที่บันทึกข้อสอบ ทำให้ key "all_questions_raw_<v>" เป็นของใหม่เสมอ
// → getAllQuestionsCached() cache-miss 100% ทุกครั้งที่ client sync หลังบันทึก แล้วต้องอ่าน+บีบอัดชีตทั้งใบ (24k แถว)
// เพื่อดึงข้อเดียว ทางนี้อ่านเฉพาะคอลัมน์ A หา row index แล้วดึงเฉพาะแถวที่ตรง
// ถ้าจำนวนข้อที่เปลี่ยนเยอะ (import/bulk) การดึงทีละแถวจะแพงกว่า จึงถอยกลับไปใช้ cache ก้อนใหญ่ตามเดิม
var CHANGED_ROWS_TARGETED_LIMIT = 50;

function getChangedQuestionRows_(ss, changedIds, changedCount, startTime) {
  var cached = getLargeCache("all_questions_raw_" + getVersionCached());
  if (cached) return JSON.parse(cached);
  if (changedCount > CHANGED_ROWS_TARGETED_LIMIT) return getAllQuestionsCached(ss, startTime);

  var qSheet = ss.getSheetByName('Questions');
  var lastRow = qSheet.getLastRow();
  if (lastRow <= 1) return [];

  if (startTime) assertNotTimedOut_(startTime, 'getChangedQuestionRows_:before_id_scan');
  var idCol = qSheet.getRange(2, 1, lastRow - 1, 1).getValues();
  var rows = [];
  for (var i = 0; i < idCol.length; i++) {
    if (!changedIds[String(idCol[i][0]).trim()]) continue;
    rows.push(qSheet.getRange(i + 2, 1, 1, 7).getValues()[0]);
    if (rows.length >= changedCount) break;
  }
  return rows;
}

function getChangedSinceTimestamp(sinceStr, filterSubject, startTime) {
  if (startTime) assertNotTimedOut_(startTime, 'getChangedSinceTimestamp:start');
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sinceMs = parseInt(sinceStr) || 0;
  var cleanFilter = filterSubject ? String(filterSubject).trim().toUpperCase() : "";

  // --- Build Category → Subject map ---
  var catToSubjectMap = getCategoryToSubjectMapCached(ss, startTime);

  // --- Scan Logs sheet using cached or fresh data ---
  var logDataJson = getLargeCache("logs_data_cache");
  var logData;
  if (logDataJson) {
    logData = JSON.parse(logDataJson);
    // Cache may hold only a last-1000-rows sample primed by a recent-since caller.
    // If its oldest row is newer than this caller's sinceMs, the window isn't covered — read fresh.
    if (sinceMs > 0 && logData.length > 1 && new Date(logData[1][0]).getTime() > sinceMs) {
      logData = null;
    }
  }
  if (!logData) {
    if (startTime) assertNotTimedOut_(startTime, 'getChangedSinceTimestamp:before_logs');
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
        putLargeCache("logs_data_cache", JSON.stringify(logData), 15, startTime); // Cache for 15s to block stamps
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
      if (i % 1000 === 0 && startTime) assertNotTimedOut_(startTime, 'getChangedSinceTimestamp:log_loop');
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
  var qData = getChangedQuestionRows_(ss, changedIds, changedIdKeys.length, startTime);
  var changedQuestions = [];

  for (var i = 0; i < qData.length; i++) {
    if (i % 500 === 0 && startTime) assertNotTimedOut_(startTime, 'getChangedSinceTimestamp:qData_loop');
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

