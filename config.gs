var SHEET_ID = '12rN8vcykEwgcPFK4LoOj18PEhj7JPhwMfz6uUkKrhJU';
var DRIVE_FOLDER_ID = '1nzLH2ia2lL2TMxfrr6Kv-5fhsWwOWSCm'; 

var VOTE_THRESHOLD_CONFIRM = 2;

var REPORT_VOTE_THRESHOLD = 5;

// ชื่อผู้รับโอนจริง (PromptPay) — ใช้ที่เดียวทั้ง prompt OCR สลิป + ตรวจ recipient match (donations.gs)
// ★ นิยามครั้งเดียวที่นี่ ห้าม hardcode ซ้ำใน prompt/การตรวจแยกกัน
var DONATION_RECIPIENT_NAME = 'ปาณัสม์ จังตระกูล';

// salt สำหรับ SHA-256(studentId) — dedup รีวิว 1 คน/วิชา (reviews.gs::hashStudentId_)
// ★ อยู่ฝั่ง server เท่านั้น: ห้ามส่งออก client, ห้ามหมุน (หมุน = hash เปลี่ยน = คนเดิมรีวิวซ้ำได้)
var STUDENT_ID_SALT = 'mdkku_reviews_salt_2026_v1';

// ────────────────────────────────────────────────────────────────────────────
// EXECUTION BUDGET — GAS ตัด execution ที่ ~6 นาที (หน้า Executions โชว์ 369.99s ซ้ำๆ = ชนเพดาน)
// UrlFetchApp ไม่มีพารามิเตอร์ timeout ให้ตั้ง → กันชนเพดานได้ทางเดียวคือ "ไม่เริ่มรอบใหม่"
//   เมื่องบเวลาใกล้หมด. global var ถูก evaluate ใหม่ทุก execution จึงใช้เป็นจุดเริ่มนับได้
// ลูปไหนที่ยิง UrlFetchApp หรือ retry ซ้ำ ต้องเช็ค execRemainingMs_() ก่อนเริ่มรอบถัดไป
// ────────────────────────────────────────────────────────────────────────────
var EXEC_START_MS = Date.now();
var EXEC_BUDGET_MS = 300000;        // 5 นาที — เหลืออีก 1 นาทีให้ serialize + ตอบกลับก่อนโดนตัด
var EXEC_FETCH_RESERVE_MS = 60000;  // กันไว้ให้ UrlFetchApp 1 ครั้งที่ช้าที่สุด

function execRemainingMs_() {
  return EXEC_BUDGET_MS - (Date.now() - EXEC_START_MS);
}

// true = เวลาเหลือไม่พอเริ่ม network call รอบใหม่
function execBudgetExhausted_() {
  return execRemainingMs_() < EXEC_FETCH_RESERVE_MS;
}

// ────────────────────────────────────────────────────────────────────────────
// AUDIT LOG — แยก spreadsheet ออกจาก SHEET_ID (คลังข้อสอบ) โดยสิ้นเชิง
// เหตุผล: log การใช้งาน/prompt AI โตไม่จำกัด → ถ้าเขียนลง SHEET_ID เดียวกันจะดัน
//   จำนวนเซลล์ชนเพดาน 10M ต่อไฟล์ แล้วทำให้ "ทั้งฐานข้อมูล" (คำถาม/โหวต/รายงาน) เขียนไม่ได้
// วิธีแก้: สร้างไฟล์ audit เดี่ยว 1 ครั้ง เก็บ ID ไว้ใน ScriptProperties (auto-provision)
//   overflow/พังของไฟล์ audit จะกระทบเฉพาะ audit ไม่แตะคลังข้อสอบ
// ────────────────────────────────────────────────────────────────────────────
// PRIVACY: ไฟล์ audit เก็บเฉพาะสถิติ "ไม่ระบุตัวตน" — ไม่มี email / studentId / clientId /
//   userAgent / ข้อความ prompt ดิบ เลย. เก็บแค่ intent tag + ฟีเจอร์ + app + เวลา
//   (จำแนก intent ฝั่ง client; ข้อความดิบไม่เคยออกจากอุปกรณ์)
function getAuditSheetId() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('AUDIT_SHEET_ID');
  if (id) return id;

  // provision ครั้งเดียว — สร้าง spreadsheet ใหม่ + 2 แท็บ (Features / AI_Intents)
  // ★ ชื่อแท็บใหม่ (Features/AI_Intents) ตั้งใจให้ไม่ชนกับ schema เก่า (Interactions/AI_Prompts)
  //   → ถ้าไฟล์ audit เคยถูกสร้างด้วย schema เก่า getAuditTab_ จะสร้างแท็บใหม่สดเสมอ ไม่เขียนผิดคอลัมน์
  var ss = SpreadsheetApp.create('MDKKUQUIZ_Audit');
  var features = ss.getSheets()[0].setName('Features');
  features.appendRow(["Timestamp", "AppId", "FeatureName"]);
  features.getRange(1, 1, 1, 3).setFontWeight("bold").setBackground("#e6f7ff");

  var intents = ss.insertSheet('AI_Intents');
  intents.appendRow(["Timestamp", "AppId", "IntentTag", "Model"]);
  intents.getRange(1, 1, 1, 4).setFontWeight("bold").setBackground("#ffe6f0");

  id = ss.getId();
  props.setProperty('AUDIT_SHEET_ID', id);
  return id;
}

// คืนแท็บชื่อ name (สร้างพร้อม header ถ้ายังไม่มี) — กันกรณีไฟล์ audit ถูก provision
//   ด้วย schema เก่าไปแล้ว: self-heal โดยไม่ต้องลบไฟล์/property เอง
function getAuditTab_(ss, name, headers) {
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(headers);
    sh.getRange(1, 1, 1, headers.length).setFontWeight("bold");
  }
  return sh;
}

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('🛠️ MDKKU Tools')
      .addItem('🔄 สั่งเรียงลำดับ Category ใหม่', 'sortCategorySheet')
      .addItem('✅ Verify ทั้งหมด (เฉพาะที่มี 1 Cat & ไม่ใช่ AI)', 'verifyAllSingleCategoryVotes')
      .addItem('📂 แยกกลุ่มวิชาอัตโนมัติ (Extracted) ทั้งหมด', 'runManualSplitExtraction')
      .addItem('📃 ตรวจสอบ Image url', 'generateImageVerificationReport')
      .addItem('🖼️ จัดระเบียบ/เปลี่ยนชื่อรูปตามวิชา', 'promptMigrateSubjectImages')
      .addToUi();
}

