
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
  if (action == 'setupAIConfig') return setupAIConfigSheet(); // idempotent one-off: สร้าง AI_Models + migrate AI_Config เป็นโครง per-model quota
  if (action == 'aiConfigStatus') return getAIConfigStatus(); // read-only diagnostic (keys masked)
  if (action == 'recoverAIConfigKeys') return recoverAIConfigKeys(e.parameter.before); // กู้ key จาก revision history (idempotent)


  return ContentService.createTextOutput("Action not defined").setMimeType(ContentService.MimeType.TEXT);
}

