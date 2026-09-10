/*
   =========================================
   supabase-mirror.gs — dual-write ไป Postgres แบบ best-effort (D14)
   แผน: Idea/active/supabase-migration-plan-v3.md §7 Phase 1, §9.11, §9.12
   =========================================

   ── หลักการที่ห้ามผิด ─────────────────────────────────────────────────────
   1. ชีทคือความจริง Postgres คือเงา. mirror พังไม่ว่ากรณีใด ต้องไม่ทำให้ handler พัง
      ⇒ sbFlush_() กลืน error ทุกชนิด แล้วบันทึก POSTGRES_MIRROR_FAIL ลง Logs
   2. ห้ามเรียก UrlFetchApp ใต้ LockService (D14). โครงจึงแยกเป็นสองจังหวะ:
        - ใน lock: จัดคิว + ถ่ายภาพชีท (getValues ครั้งเดียว = atomic)
        - นอก lock: ยิง HTTP
   3. ไม่มีคีย์ใน ScriptProperties = ทุกทางเข้าเป็น no-op เงียบๆ
      ⇒ ไฟล์นี้ deploy ได้ก่อนจะมีคีย์ โดยไม่เปลี่ยนพฤติกรรมอะไรเลยแม้แต่นิดเดียว

   ── ทำไมถ่ายภาพชีท "ใน lock" ถึงเป็นเรื่องคอขาดบาดตาย ────────────────────
   replace_*_all ลบทุกแถวที่ไม่ได้ส่งไป การ์ด §Q ของ 004 กันได้แค่ "array ว่าง"
   แต่ถ้าอ่านชีทตอน doPost อีกตัวกำลัง deleteRow อยู่ จะได้ snapshot ที่ขาดบางแถว
   ซึ่ง "ไม่ว่าง" จึงผ่านการ์ด แล้วแถวที่ขาดจะถูกลบจริงใน Postgres
   ⇒ อ่านใน lock เสมอ ยิงทีหลังได้ แต่ห้ามอ่านทีหลัง

   ── ตั้งค่าครั้งเดียว ────────────────────────────────────────────────────
   ScriptProperties: SUPABASE_URL = https://<ref>.supabase.co
                     SUPABASE_SERVICE_KEY = sb_secret_…  (ห้ามใช้ anon key)
   แล้วรัน setupSupabaseMirror() หนึ่งครั้งเพื่อสร้าง trigger ของ sweep
*/

var SB_RPC_PATH = '/rest/v1/rpc/';
var SB_QUESTION_CHUNK = 400;       // แถวต่อหนึ่ง POST — กัน payload บวมตอน adminImport
var SB_SWEEP_CURSOR_PROP = 'SB_SWEEP_CURSOR';
var SB_LOG_TAIL_ROWS = 2000;       // แถว Logs ท้ายสุดที่ sweep จะส่อง

// ชีท → RPC ที่ใช้แทนที่ทั้ง slice (004). ครอบคลุมทุก handler ที่แตะสามชีทนี้
// โดยไม่ต้อง map ฟิลด์รายตัวต่อ handler — ซึ่งเป็นจุดที่พลาดง่ายที่สุด
var SB_SLICE_RPC = {
  'Category':      'replace_categories_all',
  'Structure':     'replace_subjects_all',
  'Announcements': 'replace_announcements_all',
  'Votes':         'replace_votes_all',
  'Report':        'replace_reports_all'
};

// คิวต่อหนึ่ง execution — GAS ให้ global ใหม่ทุกครั้งที่รัน จึงไม่ต้องล้างเอง
var SB_QUEUE_ = [];
var SB_DIRTY_SHEETS_ = {};

// ────────────────────────────────────────────────────────────────────────────
// ชั้นล่าง
// ────────────────────────────────────────────────────────────────────────────

// อ่าน ScriptProperties ครั้งเดียวต่อ execution — sbEnabled_() ถูกเรียกที่ทุกทางเข้า
// รวมถึงในลูปของ bulkAddQuestionCategories ที่วิ่ง 100 รอบ
var SB_CFG_CACHED_;
function sbConfig_() {
  if (SB_CFG_CACHED_ === undefined) {
    var props = PropertiesService.getScriptProperties();
    var url = props.getProperty('SUPABASE_URL');
    var key = props.getProperty('SUPABASE_SERVICE_KEY');
    SB_CFG_CACHED_ = (url && key)
      ? { url: String(url).replace(/\/+$/, ''), key: String(key) }
      : null;
  }
  return SB_CFG_CACHED_;
}

function sbEnabled_() {
  return sbConfig_() !== null;
}

/** จัดคิวเรียก RPC หนึ่งครั้ง — ถูกเสมอที่จะเรียกใน lock เพราะไม่มี network */
function sbQueue_(fn, payload) {
  if (!sbEnabled_()) return;
  if (!payload) return;
  if (Array.isArray(payload) && payload.length === 0) return;   // §Q: อย่าส่ง array ว่างไปให้ RAISE เปล่าๆ
  SB_QUEUE_.push({ fn: fn, payload: payload });
}

/**
 * มาร์คว่าชีทนี้เปลี่ยนแล้ว — "ยังไม่อ่าน" ตรงนี้
 * เพราะ handler หลายตัวยังแก้ชีทต่ออีกหลังบรรทัดที่เรียก (เช่น addCategory → sortCategorySheet)
 * การอ่านจริงเกิดครั้งเดียวใน sbSnapshotDirtySheets_() ตอนท้ายสุดที่ยังถือ lock อยู่
 */
function sbMarkSheet_(sheetName) {
  if (!sbEnabled_()) return;
  if (!SB_SLICE_RPC[sheetName]) return;
  SB_DIRTY_SHEETS_[sheetName] = true;
}

/**
 * ★ ต้องเรียก "ก่อน releaseLock() เสมอ" ★
 * อ่านชีทที่ถูกมาร์คทั้งหมด ชีทละหนึ่ง getDataRange().getValues() แล้วจัดคิวเป็น replace_*_all
 */
function sbSnapshotDirtySheets_() {
  if (!sbEnabled_()) return;
  var names = Object.keys(SB_DIRTY_SHEETS_);
  if (!names.length) return;
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    for (var i = 0; i < names.length; i++) {
      var rows = getSheetDataJSON(names[i], ss);   // header-keyed + Date→ISO อยู่แล้ว (§9.7)
      if (!rows || !rows.length) continue;          // ชีทว่าง = อ่านพลาด ⇒ ไม่ส่ง ปล่อยให้ sweep เก็บ
      sbQueue_(SB_SLICE_RPC[names[i]], rows);
    }
  } catch (e) {
    console.error('sbSnapshotDirtySheets_: ' + e.message);
  }
  SB_DIRTY_SHEETS_ = {};
}

/**
 * ★ ต้องเรียก "หลัง releaseLock() เสมอ" ★  (D14: ห้าม UrlFetchApp ใต้ lock)
 * กลืน error ทุกชนิด — ถ้า throw ออกไป ค่าที่ handler กำลังจะ return จะถูกแทนที่ด้วย exception
 */
function sbFlush_() {
  if (!SB_QUEUE_.length) return;
  var queue = SB_QUEUE_;
  SB_QUEUE_ = [];

  var cfg = sbConfig_();
  if (!cfg) return;

  var failures = [];
  for (var i = 0; i < queue.length; i++) {
    // งบเวลาหมด = ปล่อยให้ sweep รอบหน้าเก็บ ดีกว่าโดน GAS ตัดกลางคัน (config.gs §EXECUTION BUDGET)
    if (execBudgetExhausted_()) {
      failures.push(queue[i].fn + ': ข้ามเพราะงบเวลา execution ใกล้หมด');
      continue;
    }
    var chunks = sbChunk_(queue[i].fn, queue[i].payload);
    for (var k = 0; k < chunks.length; k++) {
      var err = sbCall_(cfg, queue[i].fn, chunks[k]);
      if (err) failures.push(queue[i].fn + ': ' + err);
    }
  }

  if (failures.length) {
    try {
      writeAdminLog('SYSTEM', 'MIRROR', 'SYSTEM', 'POSTGRES_MIRROR_FAIL',
                    '', failures.length + ' rpc call(s) failed', '',
                    failures.slice(0, 5).join(' | ').slice(0, 4000), '');
    } catch (e) { console.error('POSTGRES_MIRROR_FAIL log failed: ' + e.message); }
  }
}

/** แบ่ง payload คำถามเป็นก้อน — ชุดอื่นเล็กพอที่จะไปทั้งก้อน */
function sbChunk_(fn, payload) {
  if (fn !== 'upsert_questions_batch' || !Array.isArray(payload)) return [payload];
  if (payload.length <= SB_QUESTION_CHUNK) return [payload];
  var out = [];
  for (var i = 0; i < payload.length; i += SB_QUESTION_CHUNK) {
    out.push(payload.slice(i, i + SB_QUESTION_CHUNK));
  }
  return out;
}

/** ยิง RPC หนึ่งครั้ง — คืน null ถ้าสำเร็จ, คืนข้อความ error ถ้าไม่ */
function sbCall_(cfg, fn, payload) {
  try {
    // เส้นทาง inline คือ "แอดมินเพิ่งบันทึกของจริง" ⇒ ปั้นหมวดที่ยังไม่มีได้ตามค่าเริ่มต้น
    // ต่างจาก sweep ที่เล่นซ้ำย้อนหลังและต้องปิด autocreate (§9.11 ข้อ 5, ดู sbCallNow_)
    var body = (fn === 'soft_delete_questions')
      ? { p_ids: payload }
      : { p_rows: payload };

    var res = UrlFetchApp.fetch(cfg.url + SB_RPC_PATH + fn, {
      method: 'post',
      contentType: 'application/json',
      headers: {
        // ทั้งสองหัวข้อจำเป็น: apikey ไม่มี = PostgREST ปฏิเสธก่อนถึง Authorization (§9.9)
        'apikey': cfg.key,
        'Authorization': 'Bearer ' + cfg.key
      },
      payload: JSON.stringify(body),
      muteHttpExceptions: true
    });
    var code = res.getResponseCode();
    if (code >= 200 && code < 300) return null;
    return 'HTTP ' + code + ' ' + String(res.getContentText()).slice(0, 300);
  } catch (e) {
    return e.message;
  }
}

/** เรียก RPC ตรงๆ นอกคิว — ใช้ใน sweep ที่ต้องอ่านผลลัพธ์ */
function sbCallNow_(fn, body) {
  var cfg = sbConfig_();
  if (!cfg) return null;
  try {
    var res = UrlFetchApp.fetch(cfg.url + SB_RPC_PATH + fn, {
      method: 'post',
      contentType: 'application/json',
      headers: { 'apikey': cfg.key, 'Authorization': 'Bearer ' + cfg.key },
      payload: JSON.stringify(body || {}),
      muteHttpExceptions: true
    });
    if (res.getResponseCode() >= 300) {
      return { error: 'HTTP ' + res.getResponseCode() + ' ' + String(res.getContentText()).slice(0, 300) };
    }
    return JSON.parse(res.getContentText() || 'null');
  } catch (e) {
    return { error: e.message };
  }
}

// ────────────────────────────────────────────────────────────────────────────
// ทางเข้าระดับ handler — ทุกตัวเป็น no-op เมื่อไม่มีคีย์
// ────────────────────────────────────────────────────────────────────────────

/**
 * คำถามหนึ่งข้อจาก editQuestion
 * ⚠️ ส่ง data.data.category (array ดิบ) ไม่ใช่ catToSave ที่ผ่าน JSON.stringify มาแล้ว —
 *    ตัวหลังทำงานได้ด้วยความบังเอิญของ path double-quote ใน parser §5.1 เท่านั้น
 *
 * splitRes = ค่าคืนจาก autoCreateSplitCategories (อาจเป็น undefined เมื่อมันคืนก่อนกำหนด)
 *   - finalCategories: ขั้นที่ 7 ของมันเขียนทับคอลัมน์ category "หลังจาก" handler เขียนไปแล้ว
 *     ⇒ ถ้ามีค่านี้ต้องใช้แทน q.category ไม่งั้น mirror ส่งรายการหมวดเก่าไป
 *   - sheetsChanged: มันเพิ่งเพิ่มแถวใน Category และ/หรือ Structure ⇒ ต้อง replace ทั้งสอง slice
 */
function sbMirrorQuestion_(q, splitRes) {
  if (!sbEnabled_() || !q || !q.id) return;
  sbQueue_('upsert_questions_batch', [{
    questionId: String(q.id),
    problem: q.problem,
    img: q.img,
    choices: q.choices,
    answer: q.answer,
    explain: q.explain,
    category: (splitRes && splitRes.finalCategories) || q.category
  }]);
  if (splitRes && splitRes.sheetsChanged) {
    sbMarkSheet_('Category');
    sbMarkSheet_('Structure');
  }
}

/** หลายข้อพร้อมกัน — [{questionId, category}] จาก bulkAddQuestionCategories */
function sbMirrorQuestionRows_(rows) {
  if (!sbEnabled_() || !rows || !rows.length) return;
  sbQueue_('upsert_questions_batch', rows);
}

/** adminImport: แถวดิบของชีท Questions → object ตามลำดับคอลัมน์ 0..6 (getQuestionsArray) */
function sbMirrorQuestionSheetRows_(sheetRows) {
  if (!sbEnabled_() || !sheetRows || !sheetRows.length) return;
  var out = [];
  for (var i = 0; i < sheetRows.length; i++) {
    var r = sheetRows[i];
    var qid = String(r[0] || '').trim();
    if (!qid) continue;
    out.push({
      questionId: qid,
      problem: r[1],
      img: r[2],
      choices: r[3],
      answer: r[4],
      explain: r[5],
      category: r[6]        // สตริง pseudo-JSON ก็ได้ — RPC parse ให้ตาม §5.1
    });
  }
  sbQueue_('upsert_questions_batch', out);
}

function sbMirrorQuestionDeleted_(questionId) {
  if (!sbEnabled_() || !questionId) return;
  sbQueue_('soft_delete_questions', [String(questionId)]);
}

/** โหวตหนึ่งแถวที่ "เขียนค่าสุดท้ายแล้ว" — ไม่ใช่ delta (§9.11 หมายเหตุ upsert_votes_batch) */
function sbMirrorVoteRow_(row) {
  if (!sbEnabled_() || !row) return;
  sbQueue_('upsert_votes_batch', [row]);
}

function sbMirrorVoteDeleted_(questionId, category) {
  if (!sbEnabled_()) return;
  sbQueue_('delete_votes_batch', [{ QuestionID: String(questionId), SuggestedTopic: String(category) }]);
}

/** รายงานหนึ่งแถว — คีย์คือ Time (§9.11 ข้อ 7) ⇒ ไม่มี Time = ไม่ต้องส่ง */
function sbMirrorReportRow_(row) {
  if (!sbEnabled_() || !row || !row.Time) return;
  sbQueue_('upsert_reports_batch', [row]);
}

/** เรียกจาก writeAdminLog จุดเดียว ครอบคลุมทุก action ของแอดมินรวมถึง REPORT_AUTOFIX */
function sbMirrorLogRow_(row) {
  if (!sbEnabled_() || !row) return;
  sbQueue_('insert_logs_batch', [row]);
}

// ────────────────────────────────────────────────────────────────────────────
// SWEEP — คืนสภาพสิ่งที่ inline hook พลาด (D14) รันจาก trigger ทุก 10 นาที
//
// ทำไมต้องมีทั้งที่ inline hook ครบแล้ว:
//   - POST ที่ล้มเหลวไม่มีใคร retry
//   - processVotes()/processReports() รันนอก lock และแก้ชีทเอง ไม่มี hook (ตามเจตนา)
//   - deleteQuestion ลบแถวจริงในชีท ⇒ ไม่มีอะไรใน "สภาพปัจจุบัน" ที่บอกได้ว่ามันเคยมี
//     ต้องอ่านย้อนจาก feed ของ Logs เท่านั้น
// ────────────────────────────────────────────────────────────────────────────

function runSupabaseMirrorSweep() {
  if (!sbEnabled_()) return;

  var lock = LockService.getScriptLock();
  if (!lock.tryLock(25000)) return;    // ชนกับ doPost อยู่ — รอบหน้าค่อยว่ากัน

  var snapshot = null;
  var cursorMs = parseInt(PropertiesService.getScriptProperties().getProperty(SB_SWEEP_CURSOR_PROP)) || 0;
  try {
    snapshot = sbSweepRead_(cursorMs);   // อ่านทุกอย่างใน lock ครั้งเดียว
  } catch (e) {
    console.error('sweep read: ' + e.message);
  } finally {
    lock.releaseLock();
  }
  if (!snapshot) return;

  // ตั้งแต่บรรทัดนี้ไปคือนอก lock แล้ว ยิง HTTP ได้
  var failed = false;

  // 0) คีย์ซ้ำ: ตรวจจาก snapshot ที่ถืออยู่แล้ว ไม่มีการอ่านชีทเพิ่ม และไม่แตะ failed
  //    (แถวซ้ำไม่ใช่ของค้างที่ sweep รอบหน้าจะเคลียร์ได้ ⇒ ห้ามใช้มันหยุด cursor)
  sbWarnDupSlices_(snapshot.slices);

  // 1) ลบก่อนเสมอ: ถ้า upsert วิ่งก่อนแล้ว delete ล้ม ข้อที่ลบแล้วจะโผล่กลับมาให้นักศึกษาเห็น
  if (snapshot.deletedQids.length) {
    var dRes = sbCallNow_('soft_delete_questions', { p_ids: snapshot.deletedQids });
    if (!dRes || dRes.error) failed = true;
  }

  // 2) ข้อที่ถูกแก้/นำเข้า — p_autocreate=false: ชีทยังอ้างหมวดที่ deleteCategory เพิ่งลบไป
  //    ถ้าปั้นคืนจะได้หมวดผีกลับมาทุกรอบ sweep (§9.11 ข้อ 5) ลิงก์ที่ข้ามคืนมาใน skippedLinks
  for (var i = 0; i < snapshot.questionRows.length; i += SB_QUESTION_CHUNK) {
    if (execBudgetExhausted_()) { failed = true; break; }
    var qRes = sbCallNow_('upsert_questions_batch', {
      p_rows: snapshot.questionRows.slice(i, i + SB_QUESTION_CHUNK),
      p_autocreate: false
    });
    if (!qRes || qRes.error) failed = true;
  }

  // 3) slice เล็กทั้งห้า — แทนที่ทั้งก้อน ครอบคลุมทั้งการแก้และการลบในตัวเอง
  var slices = [
    ['replace_categories_all',    snapshot.slices.Category],
    ['replace_subjects_all',      snapshot.slices.Structure],
    ['replace_announcements_all', snapshot.slices.Announcements],
    ['replace_votes_all',         snapshot.slices.Votes],
    ['replace_reports_all',       snapshot.slices.Report]
  ];
  for (var s = 0; s < slices.length; s++) {
    var rows = slices[s][1];
    if (!rows || !rows.length) continue;        // ว่าง = อ่านพลาด ⇒ ไม่ส่ง (§Q จะ RAISE อยู่ดี)
    if (execBudgetExhausted_()) { failed = true; break; }
    var sRes = sbCallNow_(slices[s][0], { p_rows: rows });
    if (!sRes || sRes.error) failed = true;
  }

  // 4) log rows เอง
  if (snapshot.logRows.length) {
    var lRes = sbCallNow_('insert_logs_batch', { p_rows: snapshot.logRows });
    if (!lRes || lRes.error) failed = true;
  }

  // 5) ★ cursor มาจาก "เวลาสูงสุดของแถว Logs ที่ประมวลผลจริง" ไม่ใช่ Date.now() ★
  //    นาฬิกาเครื่องที่เดินเร็วกว่า timestamp ในชีทจะทำให้รอบหน้าข้ามทุกแถวในช่วงต่างนั้น
  //    และล้มทั้งรอบ = ไม่ขยับ cursor เลย เพื่อให้รอบหน้าลองใหม่ทั้งหมด
  if (!failed && snapshot.maxLogMs > cursorMs) {
    PropertiesService.getScriptProperties()
      .setProperty(SB_SWEEP_CURSOR_PROP, String(snapshot.maxLogMs));
  }
  if (failed) {
    try {
      writeAdminLog('SYSTEM', 'MIRROR', 'SYSTEM', 'POSTGRES_MIRROR_FAIL', '',
                    'sweep รอบนี้ไม่ครบ — ไม่ขยับ cursor', '', '', '');
    } catch (e) { console.error(e.message); }
  }
}

/** อ่านทุกอย่างที่ sweep ต้องใช้ ในครั้งเดียวขณะยังถือ lock อยู่ */
function sbSweepRead_(cursorMs) {
  var ss = SpreadsheetApp.openById(SHEET_ID);

  // --- Logs tail: หา qid ที่เปลี่ยนตั้งแต่ cursor + แถว log เองไปด้วย ---
  var deleted = {}, touched = {}, logRows = [], maxLogMs = cursorMs;
  var logSheet = ss.getSheetByName('Logs');
  if (logSheet) {
    var lastRow = logSheet.getLastRow();
    if (lastRow > 1) {
      var take = Math.min(SB_LOG_TAIL_ROWS, lastRow - 1);
      var logVals = logSheet.getRange(lastRow - take + 1, 1, take, 10).getValues();
      for (var i = 0; i < logVals.length; i++) {
        var r = logVals[i];
        var ts = (r[0] instanceof Date) ? r[0].getTime() : new Date(r[0]).getTime();
        if (!ts || ts <= cursorMs) continue;
        if (ts > maxLogMs) maxLogMs = ts;

        logRows.push({
          Timestamp: (r[0] instanceof Date) ? r[0].toISOString() : String(r[0]),
          User: String(r[1] || ''), Role: String(r[2] || ''),
          ActionGroup: String(r[3] || ''), ActionType: String(r[4] || ''),
          TargetID: String(r[5] || ''), Details: String(r[6] || ''),
          OldValue: String(r[7] || ''), NewValue: String(r[8] || ''),
          Metadata: String(r[9] || '')
        });

        if (String(r[3]) !== 'QUESTION') continue;
        // TargetID เป็น qid เดี่ยว หรือหลาย qid คั่นด้วย comma (bulk/import — ดู router-doPost)
        var ids = String(r[5] || '').split(',');
        for (var k = 0; k < ids.length; k++) {
          var qid = ids[k].trim();
          if (!qid) continue;
          if (String(r[4]) === 'DELETE') deleted[qid] = true;
          else touched[qid] = true;
        }
      }
    }
  }

  // แถวที่ถูกลบทีหลังชนะเสมอ — upsert ไม่เคยล้าง deleted_at อยู่แล้ว (§9.11 ข้อ 4)
  // แต่ยังต้องกันไม่ให้ส่ง qid เดียวกันไปทั้งสองทางในรอบเดียว
  for (var d in deleted) { if (touched[d]) delete touched[d]; }

  // --- อ่านค่าปัจจุบันของ qid ที่ถูกแตะ จากชีท Questions ---
  var questionRows = [];
  var touchedIds = Object.keys(touched);
  if (touchedIds.length) {
    var qSheet = ss.getSheetByName('Questions');
    if (qSheet && qSheet.getLastRow() > 1) {
      var qVals = qSheet.getRange(2, 1, qSheet.getLastRow() - 1, 7).getValues();
      var want = {};
      for (var t = 0; t < touchedIds.length; t++) want[touchedIds[t]] = true;
      for (var j = 0; j < qVals.length; j++) {
        var id = String(qVals[j][0] || '').trim();
        if (!id || !want[id]) continue;
        questionRows.push({
          questionId: id, problem: qVals[j][1], img: qVals[j][2], choices: qVals[j][3],
          answer: qVals[j][4], explain: qVals[j][5], category: qVals[j][6]
        });
        delete want[id];
      }
      // qid ที่ log บอกว่าเปลี่ยนแต่หาไม่เจอในชีท = ถูกลบไปแล้วโดยไม่มี log DELETE (เช่นลบด้วยมือ)
      for (var missing in want) deleted[missing] = true;
    }
  }

  return {
    deletedQids: Object.keys(deleted),
    questionRows: questionRows,
    logRows: logRows,
    maxLogMs: maxLogMs,
    slices: {
      Category:      getSheetDataJSON('Category', ss),
      Structure:     getSheetDataJSON('Structure', ss),
      Announcements: getSheetDataJSON('Announcements', ss),
      Votes:         getSheetDataJSON('Votes', ss),
      Report:        getSheetDataJSON('Report', ss)
    }
  };
}

// ────────────────────────────────────────────────────────────────────────────
// ติดตั้ง / ตรวจสภาพ — รันจากตัวแก้ไข GAS
// ────────────────────────────────────────────────────────────────────────────

function setupSupabaseMirror() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'runSupabaseMirrorSweep') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  ScriptApp.newTrigger('runSupabaseMirrorSweep').timeBased().everyMinutes(10).create();
  console.log('sweep trigger ทุก 10 นาที: ติดตั้งแล้ว');
  console.log(sbEnabled_()
    ? 'พบ SUPABASE_URL + SUPABASE_SERVICE_KEY — mirror ทำงาน'
    : '⚠️ ยังไม่มี SUPABASE_URL / SUPABASE_SERVICE_KEY ใน ScriptProperties — mirror เป็น no-op');
}

/**
 * นับแถวของ view ผ่าน PostgREST — ใช้ Prefer: count=exact + Range: 0-0
 * คืนตัวเลข หรือ { error } ไม่เคย throw
 */
function sbCountNow_(view) {
  var cfg = sbConfig_();
  if (!cfg) return { error: 'no config' };
  try {
    // view อาจมี filter ติดมาแล้ว (เช่น 'v_categories?Status=eq.auto_created')
    // ต่อ '?select=*' ดื้อๆ จะได้ '?' สองตัว แล้ว PostgREST อ่าน filter เพี้ยนทั้งเส้น
    var sep = (view.indexOf('?') === -1) ? '?' : '&';
    var res = UrlFetchApp.fetch(cfg.url + '/rest/v1/' + view + sep + 'select=*', {
      method: 'get',
      headers: {
        'apikey': cfg.key,
        'Authorization': 'Bearer ' + cfg.key,
        'Prefer': 'count=exact',
        'Range': '0-0'
      },
      muteHttpExceptions: true
    });
    if (res.getResponseCode() >= 300) {
      return { error: 'HTTP ' + res.getResponseCode() + ' ' + String(res.getContentText()).slice(0, 200) };
    }
    // GAS ให้ header มาโดยไม่รับประกันตัวพิมพ์ — กวาดหาแบบไม่สนตัวพิมพ์
    var all = res.getAllHeaders(), cr = null;
    for (var h in all) if (String(h).toLowerCase() === 'content-range') cr = all[h];
    if (!cr) return { error: 'ไม่มี Content-Range' };
    var total = String(cr).split('/')[1];
    if (!total || total === '*') return { error: 'Content-Range ไม่มียอดรวม: ' + cr };
    return parseInt(total, 10);
  } catch (e) {
    return { error: e.message };
  }
}

/**
 * แกนกลางของสถิติคีย์ — รับ "คีย์ที่สกัดมาแล้ว" เรียงตามลำดับแถวข้อมูล
 * index i ⇒ แถวจริงในชีท i + 2 (แถว 1 เป็น header) ⇒ ผู้เรียกต้องส่งมาครบทุกแถว ห้ามกรองก่อน
 * คีย์ว่าง = ข้าม ให้ตรงกับ RPC ฝั่ง postgres ที่ CONTINUE เมื่อค่าคีย์ว่าง
 * คืน { rows, unique, dups[], rowsOfDup{} }
 */
function sbKeyStatsFromKeys_(keys) {
  // prefix กัน key ชนกับ prototype ('constructor', '__proto__', ...)
  var seen = {}, dups = [], rowsOfDup = {}, uniq = 0;
  for (var i = 0; i < keys.length; i++) {
    var k = keys[i];
    if (!k) continue;
    var kk = 'k:' + k;
    if (seen[kk]) {
      if (dups.indexOf(k) === -1) { dups.push(k); rowsOfDup[k] = [seen[kk]]; }
      rowsOfDup[k].push(i + 2);                       // เลขแถวจริงในชีท (1-based, มี header)
    } else {
      seen[kk] = i + 2;
      uniq++;
    }
  }
  return { rows: keys.length, unique: uniq, dups: dups, rowsOfDup: rowsOfDup };
}

/**
 * ต่อคีย์จากหลายคอลัมน์เป็นคีย์เดียว — ส่วนไหนว่าง = ทั้งแถวไม่มีคีย์ (คืน '')
 * ⚠️ ห้ามคืน '|' หรือ '' แบบนับเป็นคีย์จริง: upsert_subjects_batch ข้ามแถวที่ SubjectID
 *    หรือ AccordionGroup ว่าง ⇒ ถ้านับเข้ามาด้วย สองแถวว่างจะกลายเป็น "คีย์ซ้ำ" ที่ไม่มีอยู่จริง
 */
function sbJoinKey_(parts) {
  var out = [];
  for (var i = 0; i < parts.length; i++) {
    var p = String(parts[i] == null ? '' : parts[i]).trim();
    if (!p) return '';
    out.push(p);
  }
  return out.join('|');
}

/**
 * นับแถว + คีย์ซ้ำของชีทหนึ่งใบ รองรับคีย์ประกอบหลายคอลัมน์ (Structure = SubjectID|AccordionGroup)
 *
 * ⚠️ คีย์ซ้ำคือจุดบอดของการเทียบ "จำนวนแถว" เฉยๆ: ชีทมี 2 แถว, PK ฝั่ง postgres
 *    ยุบเหลือ 1 ⇒ ยอดไม่ตรงตลอดกาลและ sweep รอบถัดไปก็ไม่ช่วย เพราะไม่ใช่ของค้าง
 *    (เจอจริง 2026-09-10: Category ซ้ำหนึ่งคู่ ⇒ ชีท 1439 / postgres 1440)
 * ⚠️ อ่านเฉพาะช่วงคอลัมน์ที่เป็นคีย์ ไม่ใช่ getDataRange() — ชีทใหญ่ๆ การดึง
 *    problem/choices/explain ทั้งใบมาเพื่อ "นับคีย์" คือทางลัดไปชนลิมิต 6 นาที
 * คืน { rows, unique, dups[], rowsOfDup{} } หรือ { error }
 */
function sbSheetKeyStats_(ss, sheetName, keyHeaders) {
  var sh = ss.getSheetByName(sheetName);
  if (!sh) return { error: 'ไม่พบชีท ' + sheetName };
  var last = sh.getLastRow();
  if (last < 2) return { rows: 0, unique: 0, dups: [], rowsOfDup: {} };

  var header = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  var cols = [];
  for (var i = 0; i < keyHeaders.length; i++) {
    var col = -1;
    for (var j = 0; j < header.length; j++) {
      if (String(header[j]).trim() === keyHeaders[i]) { col = j; break; }
    }
    if (col < 0) return { error: 'ไม่พบคอลัมน์ ' + keyHeaders[i] + ' ในชีท ' + sheetName };
    cols.push(col);
  }

  var lo = Math.min.apply(null, cols), hi = Math.max.apply(null, cols);
  var values = sh.getRange(2, lo + 1, last - 1, hi - lo + 1).getValues();
  var keys = [];
  for (var r = 0; r < values.length; r++) {
    var parts = [];
    for (var c = 0; c < cols.length; c++) parts.push(values[r][cols[c] - lo]);
    keys.push(sbJoinKey_(parts));
  }
  return sbKeyStatsFromKeys_(keys);
}

/**
 * เตือนคีย์ซ้ำจาก slice ที่ sweep ถืออยู่ในหน่วยความจำแล้ว — ไม่อ่านชีทเพิ่มแม้แถวเดียว
 *
 * ทำไมต้องอยู่ใน sweep: checkSupabaseMirror() ไม่มีใครเรียกและไม่มี trigger
 *   ⇒ ถ้าเช็คแต่ในนั้น คีย์ซ้ำจะไม่มีวันถูกพบเองเลย (เจอ 2026-09-10 เพราะไล่มือ)
 * ทำไมไม่ตั้ง failed: failed = ไม่ขยับ cursor แต่แถวซ้ำมีแต่คนเท่านั้นที่ลบได้
 *   ⇒ mirror จะค้างตลอดกาลรอสิ่งที่โค้ดแก้เองไม่ได้
 * ทำไมไม่ writeAdminLog: trigger ทุก 10 นาที × ซ้ำที่ยังไม่มีใครลบ = ชีท Logs บวมไม่จบ
 * ทำไมไม่มี Votes/Report: PK ฝั่ง postgres เป็น identity ⇒ ไม่มีการยุบแถว ไม่มีคีย์ให้ซ้ำ
 */
function sbWarnDupSlices_(slices) {
  var SPEC = [
    ['Category',      ['CategoryID']],
    ['Structure',     ['SubjectID', 'AccordionGroup']],
    ['Announcements', ['Id']]
  ];
  for (var i = 0; i < SPEC.length; i++) {
    var name = SPEC[i][0], fields = SPEC[i][1], rows = slices[name];
    if (!rows || !rows.length) continue;

    var keys = [];
    for (var r = 0; r < rows.length; r++) {
      var parts = [];
      for (var f = 0; f < fields.length; f++) parts.push(rows[r][fields[f]]);
      keys.push(sbJoinKey_(parts));
    }
    var st = sbKeyStatsFromKeys_(keys);
    if (!st.dups.length) continue;

    var sample = [];
    for (var d = 0; d < Math.min(st.dups.length, 5); d++) {
      sample.push(st.dups[d] + ' (แถว ' + st.rowsOfDup[st.dups[d]].join(', ') + ')');
    }
    console.warn('⚠️ คีย์ซ้ำในชีท ' + name + ' ' + st.dups.length + ' คีย์ — postgres ยุบเหลือแถวเดียวเสมอ: ' +
                 sample.join(' | ') + ' — ลบแถวซ้ำให้เหลือใบเดียว แล้วรอ sweep รอบถัดไป');
  }
}

/** พิมพ์รายการคีย์ซ้ำของชีทหนึ่งใบ คืนจำนวน issue ที่ต้องบวกเข้า bad (0 หรือ 1) */
function sbPrintDups_(sheetName, st) {
  if (!st.dups.length) return 0;
  console.log('  ⚠️ คีย์ซ้ำในชีท ' + sheetName + ' ' + st.dups.length + ' คีย์ — postgres ยุบเหลือแถวเดียวเสมอ');
  for (var d = 0; d < Math.min(st.dups.length, 10); d++) {
    var k = st.dups[d];
    console.log('     ' + k + ' → แถว ' + st.rowsOfDup[k].join(', '));
  }
  console.log('  แก้ที่ชีท: ลบแถวซ้ำให้เหลือใบเดียว แล้วรอ sweep รอบถัดไป (10 นาที)');
  return 1;
}

/**
 * เช็คสุขภาพ: เทียบทุก slice ที่ mirror ดูแล ระหว่างชีทกับ postgres
 * ไม่เขียนอะไรทั้งนั้น เรียกได้ทุกเมื่อ
 *
 * เทียบ "คีย์ไม่ซ้ำ" ไม่ใช่ "จำนวนแถว" เพราะฝั่ง postgres มี PK ⇒ แถวซ้ำในชีทยุบหายไปเสมอ
 * Category ฝั่ง postgres มีแถว auto_created ที่ไม่มีในชีทโดยเจตนา (004 §R) จึงหักออกก่อนเทียบ
 */
function checkSupabaseMirror() {
  if (!sbEnabled_()) { console.log('mirror ปิดอยู่ (ไม่มีคีย์)'); return; }

  var v = sbCallNow_('data_version', {});
  if (!v || v.error) { console.log('data_version ล้มเหลว: ' + (v && v.error)); return; }

  var ss = SpreadsheetApp.openById(SHEET_ID);
  var bad = 0;

  // ── questions: data_version() นับให้แล้ว (ไม่รวมที่ถูกลบอ่อน)
  // ⚠️ อ่านคอลัมน์ A อย่างเดียว — 23,905 แถว × getDataRange() = ลาก problem/choices/explain
  //    ทั้งชีทมาเพื่อ "นับ id" เสี่ยงชนลิมิต 6 นาทีโดยไม่ได้อะไรเพิ่มเลย
  // ⚠️ เทียบด้วย "id ไม่ซ้ำ" ไม่ใช่จำนวนแถว: questions.question_id เป็น PK ⇒ id ซ้ำในชีท
  //    ยุบเหลือแถวเดียวฝั่ง postgres เหมือนที่ Category เคยเจอ
  var qSh = ss.getSheetByName('Questions');
  var qLast = qSh ? qSh.getLastRow() : 0;
  var qKeys = [];
  if (qLast > 1) {
    var qCol = qSh.getRange(2, 1, qLast - 1, 1).getValues();
    for (var qi = 0; qi < qCol.length; qi++) qKeys.push(sbJoinKey_([qCol[qi][0]]));
  }
  var qst = sbKeyStatsFromKeys_(qKeys);

  // data_version() นับเฉพาะแถวที่ยังไม่ถูกลบอ่อน แต่ชีทยังเก็บแถวนั้นไว้และ GAS ยังเสิร์ฟอยู่
  // ⇒ ต้องบวกกลับก่อนเทียบ ไม่งั้นรายงาน "ไม่ตรง" ตลอดกาลจนคนเลิกอ่าน
  // (ถ้าเรียกไม่ได้ก็แค่ไม่บวก — ไม่ทำให้ check ล้ม แต่ต้องบอก ไม่งั้นกลายเป็น "ไม่ตรง" ลวงๆ เงียบๆ)
  var del = sbCountNow_('questions?deleted_at=not.is.null');
  if (del && del.error) console.log('  (นับข้อที่ถูกลบอ่อนไม่ได้ — ' + del.error + ' ⇒ ไม่ได้บวกกลับ ยอดอาจแจ้งไม่ตรงลวงๆ)');
  var delN = (del && del.error) ? 0 : Number(del) || 0;
  var qExpected = Number(v.questionCount) + delN;

  console.log('cursor  : ' + v.questions);
  console.log('Questions      sheet ' + qst.rows + ' แถว / id ไม่ซ้ำ ' + qst.unique +
              ' / postgres ' + v.questionCount +
              (delN ? ' (+ ลบอ่อน ' + delN + ' = ' + qExpected + ')' : '') +
              (qExpected === qst.unique ? '  ตรงกัน' : '  ⚠️ ไม่ตรง'));
  if (qExpected !== qst.unique) bad++;
  bad += sbPrintDups_('Questions', qst);
  if (delN) {
    console.log('  หมายเหตุ: ' + delN + ' ข้อถูกลบอ่อนใน postgres แต่ยังอยู่ในชีท —');
    console.log('  คนที่อ่านผ่าน Supabase จะไม่เห็น ส่วนคนที่ตกไป GAS จะยังเห็น');
  }

  // ── slice ที่เหลือ: [ชีท, view, คอลัมน์คีย์ (ประกอบได้)]
  // Structure ใช้คีย์ประกอบตาม PK ฝั่ง postgres (subject_id, accordion_group)
  // Votes/Report ไม่อยู่ในนี้: PK เป็น identity ⇒ ไม่มีคีย์จากชีทให้เทียบ และยังไม่มี read path
  var SLICES = [
    ['Category',      'v_categories',    ['CategoryID']],
    ['Structure',     'v_structure',     ['SubjectID', 'AccordionGroup']],
    ['Announcements', 'v_announcements', ['Id']]
  ];

  for (var i = 0; i < SLICES.length; i++) {
    var sheetName = SLICES[i][0], view = SLICES[i][1], keyCol = SLICES[i][2];

    var st = sbSheetKeyStats_(ss, sheetName, keyCol);
    if (st.error) { console.log(sheetName + ': อ่านชีทไม่ได้ — ' + st.error); bad++; continue; }

    var pg = sbCountNow_(view);
    if (pg && pg.error) { console.log(sheetName + ': นับ ' + view + ' ไม่ได้ — ' + pg.error); bad++; continue; }

    // Category: หักแถว auto_created ออก (ไม่มีต้นทางในชีท จึงไม่ควรถูกนับว่าเกิน)
    var extra = 0;
    if (sheetName === 'Category') {
      var auto = sbCountNow_('v_categories?Status=eq.auto_created');
      // เงียบตรงนี้ = รายงาน "ไม่ตรง" ทั้งที่ระบบปกติ เพราะหักแถว auto_created ไม่ออก
      if (auto && auto.error) console.log('  (นับ auto_created ไม่ได้ — ' + auto.error + ' ⇒ ไม่ได้หักออก)');
      else extra = auto;
    }

    var expected = pg - extra;
    var okRow = (expected === st.unique);
    console.log(sheetName + '       sheet ' + st.rows + ' แถว / คีย์ไม่ซ้ำ ' + st.unique +
                ' / postgres ' + pg + (extra ? ' (auto_created ' + extra + ')' : '') +
                (okRow ? '  ตรงกัน' : '  ⚠️ ไม่ตรง'));
    if (!okRow) bad++;
    bad += sbPrintDups_(sheetName, st);
  }

  console.log('');
  console.log(bad === 0 ? 'sum: OK - all slices match' : 'sum: WARN - ' + bad + ' issue(s)');
}
