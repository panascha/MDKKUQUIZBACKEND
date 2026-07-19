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
      .addItem('🖼️ จัดระเบียบ/เปลี่ยนชื่อรูปตามวิชา', 'promptMigrateSubjectImages')
      .addToUi();
}

