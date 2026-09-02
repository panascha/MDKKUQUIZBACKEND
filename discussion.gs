/* =========================================================================
   Feature 4 (main-task): Peer Discussion Thread (per-question)
   Idea/active/peer-discussion-thread.md — 10 locked decisions, do not re-derive.
   getDiscussion (doGet) รวม comments + reports (Report sheet) + revisions (Logs sheet)
   ต่อ qid ในเรียกเดียว, cache disc_<qid> 5 นาที (DISCUSSION_CACHE_TTL_SEC)
   ========================================================================= */

// idempotent: สร้างชีต Discussion ถ้ายังไม่มี (mirror setupAiFeedbackSheet)
function setupDiscussionSheet() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(DISCUSSION_SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(DISCUSSION_SHEET_NAME);

  var headers = ["Timestamp", "QuestionID", "Email", "Nickname", "Tag", "Text", "Status"];
  if (!sheet.getRange(1, 1).getValue()) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#e6f7ff");
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// tag 4 ตัวอักษรจาก SHA-256(email) — คงที่ต่ออีเมล กันปลอมฝั่ง client, ไม่ส่งอีเมลจริงกลับ
function computeEmailTag_(email) {
  var bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(email || "").trim().toLowerCase());
  var hex = bytes.map(function (b) {
    var v = (b < 0 ? b + 256 : b).toString(16);
    return v.length === 1 ? "0" + v : v;
  }).join("");
  return hex.slice(0, 4);
}

function readDiscussionComments_(qid, ss) {
  if (!ss) ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(DISCUSSION_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) return [];
  var rows = sheet.getDataRange().getValues();
  var out = [];
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][1]).trim() !== qid || rows[i][6] !== "visible") continue;
    var ts = rows[i][0] instanceof Date ? rows[i][0].toISOString() : String(rows[i][0]);
    out.push({ timestamp: ts, nickname: rows[i][3], tag: rows[i][4], text: rows[i][5] });
  }
  return out; // เก่าสุดก่อนอยู่แล้ว (ลำดับ appendRow) — ตาม decision #10
}

// Report sheet columns: From,Category,QuestionID,Question,Image,Choices,SuggestedAnswer,ReportDetail,Time,Status,AdminNote,Done,SuggestedExplain,VoteCount
// decision #7: เปิด ReportDetail/SuggestedAnswer/Time/Status/VoteCount, ซ่อน From (ผู้รายงาน)
function readQuestionReports_(qid, ss) {
  if (!ss) ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName("Report");
  if (!sheet || sheet.getLastRow() < 2) return [];
  var rows = sheet.getDataRange().getValues();
  var out = [];
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][2]) !== qid) continue;
    var t = rows[i][8] instanceof Date ? rows[i][8].toISOString() : String(rows[i][8]);
    out.push({
      reportDetail: rows[i][7],
      suggestedAnswer: rows[i][6],
      time: t,
      status: rows[i][9],
      voteCount: rows[i][13] || 0
    });
  }
  return out;
}

// Logs columns: Timestamp,User,Role,ActionGroup,ActionType,TargetID,Details,OldValue,NewValue,Metadata
// decision #8: filter ActionGroup=QUESTION/ActionType=EDIT/TargetID=qid, diff แบบย่อ, ซ่อนชื่อแอดมิน
// ใช้ logs_data_cache (15s TTL) แบบเดียวกับ getChangedSinceTimestamp — กัน full-scan Logs ซ้ำทุกครั้งที่เปิด discussion
function readQuestionRevisions_(qid, ss, startTime) {
  if (!ss) ss = SpreadsheetApp.openById(SHEET_ID);
  var logDataJson = getLargeCache("logs_data_cache");
  var logData;
  if (logDataJson) {
    logData = JSON.parse(logDataJson);
  } else {
    if (startTime) assertNotTimedOut_(startTime, 'readQuestionRevisions_:before_logs');
    var sheet = ss.getSheetByName("Logs");
    if (!sheet || sheet.getLastRow() < 2) return [];
    logData = sheet.getDataRange().getValues();
    putLargeCache("logs_data_cache", JSON.stringify(logData), 15, startTime);
  }
  var out = [];
  for (var i = 1; i < logData.length; i++) {
    if (i % 1000 === 0 && startTime) assertNotTimedOut_(startTime, 'readQuestionRevisions_:loop');
    if (logData[i][3] !== "QUESTION" || logData[i][4] !== "EDIT" || String(logData[i][5]) !== qid) continue;
    var t = logData[i][0] instanceof Date ? logData[i][0].toISOString() : String(logData[i][0]);
    out.push({ time: t, diff: describeQuestionEditDiff_(logData[i][7], logData[i][8]) });
  }
  return out;
}

function normalizeQuestionDiffKeys_(obj) {
  var out = {};
  for (var k in obj) {
    if (obj.hasOwnProperty(k)) out[String(k).toLowerCase()] = obj[k];
  }
  return out;
}

// สรุป field ที่เปลี่ยนแบบย่อจาก OldValue/NewValue (JSON string, เขียนโดย writeAdminLog ตอน editQuestion)
// คีย์ฝั่ง old (จาก header ชีต) กับ new (จาก data.data) case ไม่ตรงกัน → normalize เป็น lowercase ก่อนเทียบ
function describeQuestionEditDiff_(oldValStr, newValStr) {
  try {
    var o = normalizeQuestionDiffKeys_(JSON.parse(oldValStr));
    var n = normalizeQuestionDiffKeys_(JSON.parse(newValStr));
    var labels = [];
    if (o.answer !== undefined && n.answer !== undefined && String(o.answer) !== String(n.answer)) {
      labels.push("แก้เฉลย: " + o.answer + "→" + n.answer);
    }
    if (o.choices !== undefined && n.choices !== undefined && String(o.choices) !== String(n.choices)) {
      labels.push("แก้ตัวเลือก");
    }
    if (o.explain !== undefined && n.explain !== undefined && String(o.explain) !== String(n.explain)) {
      labels.push("แก้คำอธิบาย");
    }
    if (o.problem !== undefined && n.problem !== undefined && String(o.problem) !== String(n.problem)) {
      labels.push("แก้โจทย์");
    }
    if (o.img !== undefined && n.img !== undefined && String(o.img) !== String(n.img)) {
      labels.push("แก้รูปภาพ");
    }
    if (o.category !== undefined && n.category !== undefined && JSON.stringify(o.category) !== JSON.stringify(n.category)) {
      labels.push("แก้หมวดหมู่");
    }
    return labels.length ? labels.join(", ") : "แก้ไขคำถาม";
  } catch (e) {
    return "แก้ไขคำถาม";
  }
}

// doGet handler: comments + reports + revisions รวมในเรียกเดียว, cache 5 นาทีต่อ qid
function getDiscussionData(qid, startTime) {
  qid = String(qid || "").trim();
  if (!qid) {
    return ContentService.createTextOutput(JSON.stringify({ result: "error", message: "missing qid" }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  var cache = CacheService.getScriptCache();
  var cacheKey = "disc_" + qid;
  var cached = cache.get(cacheKey);
  if (cached) return ContentService.createTextOutput(cached).setMimeType(ContentService.MimeType.JSON);

  if (startTime) assertNotTimedOut_(startTime, 'getDiscussionData');
  var ss = SpreadsheetApp.openById(SHEET_ID); // เปิดครั้งเดียว ส่งต่อให้ sub-helpers ทั้ง 3 ตัว กัน openById ซ้ำ
  if (startTime) assertNotTimedOut_(startTime, 'getDiscussionData:after_openById');
  var comments = readDiscussionComments_(qid, ss);
  if (startTime) assertNotTimedOut_(startTime, 'getDiscussionData:after_comments');
  var reports = readQuestionReports_(qid, ss);
  if (startTime) assertNotTimedOut_(startTime, 'getDiscussionData:after_reports');
  var revisions = readQuestionRevisions_(qid, ss, startTime);
  if (startTime) assertNotTimedOut_(startTime, 'getDiscussionData:after_revisions');
  var payload = JSON.stringify({
    result: "success",
    comments: comments,
    reports: reports,
    revisions: revisions
  });
  cache.put(cacheKey, payload, DISCUSSION_CACHE_TTL_SEC); // เขียน cache แม้ผลว่างเปล่า กัน full-scan Logs ซ้ำทุกครั้งที่เปิด
  return ContentService.createTextOutput(payload).setMimeType(ContentService.MimeType.JSON);
}

// เรียกใต้ localized-15s lock เท่านั้น — re-count เพดาน 100 comment "ใต้ lock" กัน race (สอง request ผ่านพร้อมกันที่ 99)
function postDiscussionCommentLocked_(qid, email, nickname, text) {
  var sheet = setupDiscussionSheet();
  var rows = sheet.getDataRange().getValues();
  var count = 0;
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][1]).trim() === qid && rows[i][6] === "visible") count++;
  }
  if (count >= DISCUSSION_MAX_COMMENTS) {
    return { ok: false, message: "กระทู้เต็มแล้ว (สูงสุด " + DISCUSSION_MAX_COMMENTS + " ความคิดเห็น)" };
  }
  var tag = computeEmailTag_(email);
  var now = new Date();
  sheet.appendRow([now, qid, email, nickname, tag, text, "visible"]);
  CacheService.getScriptCache().remove("disc_" + qid);
  return { ok: true, comment: { timestamp: now.toISOString(), nickname: nickname, tag: tag, text: text } };
}

// เรียกใต้ localized-15s lock เท่านั้น — self-delete (email ตรง) หรือ admin (isAdmin=true)
// purge cache ด้วย qid ที่อ่านจากแถวจริง ไม่ใช่จาก client (กัน client ส่ง qid ผิดแล้ว cache thread อื่นเพี้ยน)
function deleteDiscussionCommentLocked_(qid, timestamp, requestorEmail, isAdmin) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(DISCUSSION_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) return { ok: false, message: "ไม่พบความคิดเห็น" };
  var rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    var rowTs = rows[i][0] instanceof Date ? rows[i][0].toISOString() : String(rows[i][0]);
    if (String(rows[i][1]).trim() !== qid || rowTs !== timestamp) continue;
    if (!isAdmin && rows[i][2] !== requestorEmail) {
      return { ok: false, message: "ไม่มีสิทธิ์ลบความคิดเห็นนี้" };
    }
    sheet.getRange(i + 1, 7).setValue("deleted");
    CacheService.getScriptCache().remove("disc_" + rows[i][1]);
    return { ok: true };
  }
  return { ok: false, message: "ไม่พบความคิดเห็น" };
}
