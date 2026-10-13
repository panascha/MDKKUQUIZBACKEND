/* =========================================
   Student "Verify answer key with Multi-AI" — qid-only wrapper over verifyQuestionBatch
   =========================================
   Lock-free (CacheService writes only; LLM calls never under a lock).
   Auth: verifyAnySession (any KKU login). Client sends ONLY qid — the question is
   loaded server-side, so no client text reaches a model (no LLM proxy / cache poison / injection).
   Order: hit -> inflight -> per-user limit -> daily budget -> set inflight -> run.
   VERIFY_SOLVER_A/B + VERIFY_JUDGE come from verify-questions.gs (not redefined here).
*/

var VERIFY_STUDENT_RATE_LIMIT = 5;      // cache-MISS runs per hour per email
var VERIFY_STUDENT_DAILY_BUDGET = 150;  // cache-MISS runs per day, whole app
var VERIFY_CACHE_TTL = 21600;           // 6h (CacheService max)
var VERIFY_NEG_TTL = 600;               // inconclusive verdicts
var VERIFY_INFLIGHT_TTL = 150;
var VERIFY_PROMPT_VER = 'v1';

function verifyStudentErr_(message, extra) {
  var o = { result: 'error', message: message };
  if (extra) for (var k in extra) o[k] = extra[k];
  return o;
}

function verifyStudentMd5_(s) {
  var d = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, s, Utilities.Charset.UTF_8);
  var out = '';
  for (var i = 0; i < d.length; i++) out += ('0' + (d[i] & 0xff).toString(16)).slice(-2);
  return out;
}

// คืน {q} หรือ {err} — q ไม่มี suggestedAnswer/reportDetail โดยตั้งใจ
function buildStudentVerifyQuestion_(row) {
  if (String(row[2] == null ? '' : row[2]).trim()) return { err: verifyStudentErr_('not_verifiable', { reason: 'has_image' }) };
  var choicesArr = String(row[3] == null ? '' : row[3]).split('///').map(function (c) { return c.trim(); });
  var mediaRe = /^<svg|drive\.google\.com|^https?:|\[IMAGE_PENDING\]|googleusercontent/i;
  for (var i = 0; i < choicesArr.length; i++) {
    if (mediaRe.test(choicesArr[i])) return { err: verifyStudentErr_('not_verifiable', { reason: 'media_choice' }) };
  }
  var dbIdx = choicesArr.indexOf(String(row[4] == null ? '' : row[4]).trim());
  if (choicesArr.length < 2 || dbIdx < 0) return { err: verifyStudentErr_('not_verifiable', { reason: 'bad_data' }) };
  return { q: { qid: String(row[0]), questionText: String(row[1]), choices: choicesArr, correctAnswer: dbIdx } };
}

function verifyStudentVerdictTag_(v, dbIdx) {
  if (v.confidence === 'consensus-verified') return 'matches_db';
  if (v.confidence === 'debate-resolved') return v.verifiedAnswer === dbIdx ? 'matches_db' : 'disagrees_db';
  if (v.confidence === 'solver-consensus-unjudged') return 'disagrees_db';
  return 'inconclusive'; // single-model, judge-failed-fallback
}

function verifyStudentTrim_(s) {
  return s == null ? s : String(s).slice(0, 200);
}

function verifyQuestionStudent(data) {
  var qid = String(data.qid || '').trim();
  if (!qid || qid.length > 100) return verifyStudentErr_('bad_request');

  var user = verifyAnySession(data.sessionToken);
  if (!user || !user.email) return verifyStudentErr_('session_expired');

  var rows = getAllQuestionsCached(null, null);
  var row = null;
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i][0]) === qid) { row = rows[i]; break; }
  }
  if (!row) return verifyStudentErr_('not_found');

  var built = buildStudentVerifyQuestion_(row);
  if (built.err) return built.err;
  var q = built.q;

  var hash = verifyStudentMd5_([VERIFY_PROMPT_VER, VERIFY_SOLVER_A, VERIFY_SOLVER_B, VERIFY_JUDGE,
    String(row[1]), String(row[3]), String(row[4])].join('\u0001'));
  var key = 'vfy_' + hash;
  var flKey = 'vfyfl_' + hash;
  var cache = CacheService.getScriptCache();

  var hit = cache.get(key);
  if (hit) {
    var cachedRes = JSON.parse(hit);
    cachedRes.cached = true;
    cachedRes.qid = qid;
    if (cachedRes.verified) cachedRes.verified.qid = qid;
    return cachedRes;
  }

  if (cache.get(flKey)) return { result: 'pending', retryAfterSec: 8 };

  if (!checkActionRateLimit('rl_vfystu_', user.email, VERIFY_STUDENT_RATE_LIMIT)) {
    return verifyStudentErr_('rate_limited', { retryAfterSec: 3600 });
  }

  // daily counter ต้อง TTL 86400 — checkActionRateLimit ใช้ 3600 สั้นเกินไป
  var budgetKey = 'vfy_budget_' + Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyyMMdd');
  var used = parseInt(cache.get(budgetKey) || '0', 10);
  if (used >= VERIFY_STUDENT_DAILY_BUDGET) return verifyStudentErr_('budget_exhausted');
  cache.put(budgetKey, String(used + 1), 86400);

  cache.put(flKey, '1', VERIFY_INFLIGHT_TTL);
  var res;
  try {
    res = verifyQuestionBatch([q], user);
  } finally {
    cache.remove(flKey);
  }

  if (!res || res.result !== 'success' || !res.verified || !res.verified.length) {
    var detail = (res && res.errors && res.errors[0] && res.errors[0].error) || (res && res.message) || '';
    return verifyStudentErr_('ai_failed', { detail: verifyStudentTrim_(detail) });
  }

  var v = res.verified[0];
  v.qid = qid;
  if (v.solverErrors) {
    v.solverErrors = v.solverErrors.map(function (e) { return { model: e.model, error: verifyStudentTrim_(e.error) }; });
  }
  if (v.judgeError) v.judgeError = verifyStudentTrim_(v.judgeError);

  var tag = verifyStudentVerdictTag_(v, q.correctAnswer);
  var payload = {
    result: 'success',
    cached: false,
    verifiedAt: new Date().toISOString(),
    qid: qid,
    verdictTag: tag,
    dbAnswerIndex: q.correctAnswer,
    choices: q.choices,
    verified: v
  };
  try {
    cache.put(key, JSON.stringify(payload), tag === 'inconclusive' ? VERIFY_NEG_TTL : VERIFY_CACHE_TTL);
  } catch (e) { /* cache failure must not fail the request */ }
  return payload;
}
