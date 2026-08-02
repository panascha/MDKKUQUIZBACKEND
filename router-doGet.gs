
// เพดานเวลาต่อ doGet execution — เช็ค checkpoint ระหว่างขั้นตอนหนักๆ (getAllData/getQuestions cache miss)
// เพื่อตัดจบก่อน Google ฆ่า container ที่ 360s (เห็นจาก Executions log: 4-5 execution ค้าง 100-311s, ตัวหนึ่ง timeout ที่ 369.985s)
var DOGET_TIMEOUT_MS = 90 * 1000;
var DOGET_TIMEOUT_MARK = 'DOGET_TIMEOUT';

function assertNotTimedOut_(startTime, context) {
  if (Date.now() - startTime > DOGET_TIMEOUT_MS) {
    throw new Error(DOGET_TIMEOUT_MARK + ':' + (context || ''));
  }
}

function doGet(e) {
  try {
    return doGet_(e);
  } catch (err) {
    var msg = err.toString();
    if (msg.indexOf(DOGET_TIMEOUT_MARK) !== -1) {
      return ContentService.createTextOutput(JSON.stringify({
        result: 'error',
        message: 'Request timeout - please filter by subject or use Delta Sync'
      })).setMimeType(ContentService.MimeType.JSON);
    }
    return ContentService.createTextOutput(JSON.stringify({
      'result': 'error',
      'message': msg
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet_(e) {
  var __startTime = Date.now();
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
  if (action == 'getStructure') return getStructureDataCached(e.parameter.subject, __startTime);
  if (action == 'getQuestions') {
    // ไม่มี caller ปัจจุบัน (REAL) เรียกโดยไม่มี subject — บล็อกเส้นทางดึงคำถามทั้งหมดแบบไม่กรองที่ไม่เคยถูกใช้จริง
    if (!e.parameter.subject) {
      return ContentService.createTextOutput(JSON.stringify({
        result: 'error',
        message: 'subject is required for getQuestions - please filter by subject'
      })).setMimeType(ContentService.MimeType.JSON);
    }
    return getQuestionsDataCached(e.parameter.subject, null, __startTime);
  }
  if (action == 'getPendingVotes') return getPendingVotesData(e.parameter.qid, __startTime);
  if (action == 'getPendingReports') return getPendingReportsData(e.parameter.qid, __startTime);
  if (action == 'getPendingVotesReports') return getPendingVotesReportsData(e.parameter.subject, __startTime);
  if (action == 'getAllData') return getAllDataForAdminCached(__startTime);
  if (action == 'getLogsPage') return getLogsPageData(e.parameter.offset, e.parameter.limit, __startTime);
  if (action == 'getPendingReportCount') return getPendingReportCount(e.parameter.subject, __startTime);
  if (action == 'getChangedSince') return getChangedSinceTimestamp(e.parameter.since, e.parameter.subject, __startTime);
  if (action == 'getRelatedQuestions') return getRelatedQuestionsData(e.parameter.subject, __startTime); // Feature 4: relations map ต่อวิชา (อ่านอย่างเดียว, chunked cache)
  if (action == 'getKB') return getKBData(e.parameter.subject, __startTime); // §1.8 KB corpus: chunks ต่อวิชา (public read, chunked cache)
  if (action == 'getGlossary') return getGlossaryData(e.parameter.subject, __startTime); // Feature 2: glossary ต่อวิชา (public read, chunked cache)
  if (action == 'getHighYield') return getHighYieldData(e.parameter.category, __startTime); // Feature 3: ชีทสรุป high-yield ต่อหมวด (public read, chunked cache)
  if (action == 'getKeywordIndex') return getKeywordIndexData(e.parameter.category, __startTime); // Feature 6: คำสำคัญที่ออกบ่อย ต่อหมวด (public read, chunked cache — list)
  if (action == 'getDiscussion') return getDiscussionData(e.parameter.qid, __startTime); // Feature 4 (main-task): comments+reports+revisions ต่อ qid (public read, cache disc_<qid> 5 นาที)
  if (action == 'setupIntelSphere') return setupIntelSphereSheet(); // idempotent one-off: สร้าง tab IntelSphere_Keys ถ้ายังไม่มี
  if (action == 'setupAIConfig') return setupAIConfigSheet(); // idempotent one-off: สร้าง AI_Models + migrate AI_Config เป็นโครง per-model quota
  if (action == 'aiConfigStatus') return getAIConfigStatus(); // read-only diagnostic (keys masked)
  if (action == 'discoverGeminiModels') {
    // เรียก Gemini จริงกินโควต้า key ที่บริจาค — public ไม่มี session token ให้ผูก key เลยจำกัดรวมทั้งระบบ
    if (!checkActionRateLimit('rl_discovergemini_', 'global', 5)) {
      return ContentService.createTextOutput(JSON.stringify({
        result: 'error', message: 'Rate limited (max 5/hour)'
      })).setMimeType(ContentService.MimeType.JSON);
    }
    return discoverGeminiModels(); // read-only: live models.list (แหล่งความจริงของ model IDs)
  }
  if (action == 'getAIModels') return getAIModels(); // read-only: ทะเบียน AI_Models สำหรับ admin panel (P2-Q6)
  // NOTE: gemini-sync ที่ mutate/กิน quota (reconcile/enable/purge/runGeminiModelSync/runGeminiToolProbe/installGeminiSyncTriggers)
  // เจตนา NOT exposed ทาง doGet — deployment นี้ public no-auth (frontend นิสิตเรียก getStructure ฯลฯ). Trigger เรียก fn ตรง,
  // manual ก็รันจาก Apps Script editor. verify actions (probeGeminiTier/verifyRpmCooldown/verifyToolSmokeTest) ก็เอาออก (scaffolding + verifyRpmCooldown burst pool ที่นิสิตใช้ร่วม).
  if (action == 'recoverAIConfigKeys') return recoverAIConfigKeys(e.parameter.before); // กู้ key จาก revision history (idempotent)


  return ContentService.createTextOutput("Action not defined").setMimeType(ContentService.MimeType.TEXT);
}

