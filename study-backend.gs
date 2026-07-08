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
