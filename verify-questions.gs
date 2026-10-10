/* =========================================
   Question Verification — 2-Model Debate + Arbiter (admin-tier, lock-free LLM)
   =========================================
   Flow:
   1. Parallel solve: DeepSeek-V4-Pro + Claude-Sonnet-4.5 via executeChatbotQuery
   2. Auto-approve only when both match AND match DB answer
   3. Else escalate to Gemini judge (callGeminiAI, parametric memory only)
   Auth: verifyAdmin OR verifySessionToken (triple-auth pattern)
   Lock: 25s admin lock ONLY around orchestration; LLM calls outside lock.
*/

// Solver models อ้างอิง PROVIDER_MODEL_MAP (maintenance.gs) — แพลตฟอร์มเปลี่ยนชื่อรุ่นเมื่อไหร่
// ไฟล์นี้ตามไปเอง ไม่ต้องแก้สองที่. maintenance.gs โหลดก่อน verify-questions.gs (เรียงตามชื่อไฟล์)
// และ var hoisting ทำให้ typeof guard ปลอดภัยแม้ลำดับโหลดเปลี่ยน — ตกไปใช้ literal เดิม
var VERIFY_SOLVER_A = (typeof PROVIDER_MODEL_MAP !== 'undefined' && PROVIDER_MODEL_MAP && PROVIDER_MODEL_MAP["Deepseek"]) || "deepseek-v4-pro";
var VERIFY_SOLVER_B = (typeof PROVIDER_MODEL_MAP !== 'undefined' && PROVIDER_MODEL_MAP && PROVIDER_MODEL_MAP["Claude"]) || "claude-sonnet-5.5";
// judge ไม่ derive จาก map โดยตั้งใจ — PROVIDER_MODEL_MAP.Gemini คือ 3.6-flash แต่หมายเหตุ
// Grounding OFF ผูกกับ 3.5-flash ตัวนี้โดยเฉพาะ
var VERIFY_JUDGE = "gemini-3.5-flash"; // cheap, high RPD; Grounding OFF (commented at ai-gemini.gs:1212)

/**
 * Orchestrator: verify a batch of questions
 * @param {Array} questions — [{qid, questionText, choices:[...], correctAnswer: N}]
 * @param {Object} adminUser — {email, role} from verifySession/verifyAdmin
 * @returns {Object} {result, verified: [...], errors: [...]}
 */
function verifyQuestionBatch(questions, adminUser) {
  if (!Array.isArray(questions) || questions.length === 0) {
    return { result: 'error', message: 'questions array required' };
  }
  var verified = [];
  var errors = [];

  for (var i = 0; i < questions.length; i++) {
    var q = questions[i];
    // งบ execution ใกล้หมด — ข้อนี้ต้องยิง LLM อย่างน้อย 2 ครั้ง เริ่มไปก็โดน GAS ตัดกลางคัน
    // ใช้ continue ไม่ใช่ break เพื่อให้ทุกข้อที่เหลือมี entry ใน errors[] (frontend อ่าน errors[0].error)
    if (execRemainingMs_() < 30000) {
      errors.push({ qid: q.qid, error: 'งบเวลาประมวลผลใกล้หมด — ข้อนี้ยังไม่ได้ตรวจ กรุณาลองใหม่อีกครั้ง' });
      continue;
    }
    try {
      // 1) Solve with both models (LLM calls outside lock)
      // แยก try ต่อโมเดล — provider เดียวล่ม/timeout ต้องไม่ลากอีกตัวที่ตอบสำเร็จตกไปด้วย
      var solveA = null, solveB = null;
      var solverErrors = [];
      try {
        solveA = solveWithModel_(q, VERIFY_SOLVER_A);
      } catch (eA) {
        solverErrors.push({ model: VERIFY_SOLVER_A, error: eA.message || String(eA) });
      }
      try {
        solveB = solveWithModel_(q, VERIFY_SOLVER_B);
      } catch (eB) {
        solverErrors.push({ model: VERIFY_SOLVER_B, error: eB.message || String(eB) });
      }

      if (!solveA && !solveB) {
        throw new Error('ทั้งสองโมเดลตอบไม่สำเร็จ — ' + solverErrors.map(function (x) {
          return x.model + ': ' + x.error;
        }).join(' | '));
      }

      // เหลือโมเดลเดียว = ไม่มีสัญญาณ consensus เลย ห้ามส่งเข้า judge
      // (buildJudgePrompt_ ออกแบบมาสำหรับสองความเห็น จะตัดสินมั่นใจเกินจริงจากความเห็นเดียว)
      if (!solveA || !solveB) {
        var only = solveA || solveB;
        verified.push({
          qid: q.qid,
          verifiedAnswer: only.choice,
          confidence: 'single-model',
          models: [only.model],
          judgeUsed: false,
          solvers: [{ model: only.model, choice: only.choice, rationale: only.rationale }],
          solverErrors: solverErrors,
          rationale: only.rationale || ''
        });
        continue;
      }

      var dbAnswer = Number(q.correctAnswer);

      // 2) Auto-approve path: both match AND agree with DB
      if (solveA.choice === solveB.choice && solveA.choice === dbAnswer) {
        verified.push({
          qid: q.qid,
          verifiedAnswer: solveA.choice,
          confidence: 'consensus-verified',
          models: [VERIFY_SOLVER_A, VERIFY_SOLVER_B],
          judgeUsed: false,
          solvers: [
            { model: VERIFY_SOLVER_A, choice: solveA.choice, rationale: solveA.rationale },
            { model: VERIFY_SOLVER_B, choice: solveB.choice, rationale: solveB.rationale }
          ],
          rationale: solveA.rationale || ''
        });
        continue;
      }

      // 3) Escalate to judge
      // judgeDisagreement_ throw เสมอเมื่อล้ม (โควต้าหมด / JSON เพี้ยน) ไม่เคยคืน null
      // ถ้าปล่อยหลุดขึ้นไป catch ด้านนอก rationale ของ solver ทั้งสองตัวจะหายไปทั้งหมด
      var judgeRes = null, judgeError = null;
      try {
        judgeRes = judgeDisagreement_(q, solveA, solveB);
      } catch (eJ) {
        judgeError = eJ.message || String(eJ);
      }

      if (!judgeRes) {
        // A===B ≠ DB + judge ล่ม → เชื่อ solver consensus (ทั้งคู่เห็นตรงกัน) ไม่ fallback ไป DB ที่โมเดลปฏิเสธ
        // frontend: confidence นี้ขึ้น badge เหลือง + เปิดปุ่ม Apply (admin ยังต้องกดยืนยัน)
        if (solveA.choice === solveB.choice && solveA.choice !== dbAnswer) {
          verified.push({
            qid: q.qid,
            verifiedAnswer: solveA.choice,
            confidence: 'solver-consensus-unjudged',
            judgeError: judgeError,
            models: [VERIFY_SOLVER_A, VERIFY_SOLVER_B],
            judgeModel: VERIFY_JUDGE,
            judgeUsed: false,
            solvers: [
              { model: VERIFY_SOLVER_A, choice: solveA.choice, rationale: solveA.rationale },
              { model: VERIFY_SOLVER_B, choice: solveB.choice, rationale: solveB.rationale }
            ],
            rationale: solveA.rationale || solveB.rationale || ''
          });
          continue;
        }
        // A≠B + judge ล่ม = ยังไม่ verified. verifiedAnswer คืนเฉลยเดิมใน DB ไว้เป็น
        // placeholder ให้ frontend เรนเดอร์ได้เท่านั้น ห้ามตีความว่าผ่านการตรวจ
        // (frontend ต้องอ่าน confidence นี้แล้วขึ้น badge เตือน + ซ่อนปุ่ม Apply)
        verified.push({
          qid: q.qid,
          verifiedAnswer: dbAnswer,
          confidence: 'judge-failed-fallback',
          judgeError: judgeError,
          models: [VERIFY_SOLVER_A, VERIFY_SOLVER_B],
          judgeModel: VERIFY_JUDGE,
          judgeUsed: false,
          solvers: [
            { model: VERIFY_SOLVER_A, choice: solveA.choice, rationale: solveA.rationale },
            { model: VERIFY_SOLVER_B, choice: solveB.choice, rationale: solveB.rationale }
          ],
          rationale: ''   // ไหลเข้า data('mv').explanation — ห้ามใส่ข้อความที่ยังไม่ผ่านการตัดสิน
        });
        continue;
      }

      verified.push({
        qid: q.qid,
        verifiedAnswer: judgeRes.verifiedAnswer,
        // confidence = สถานะ flow (frontend อ่านเพื่อเลือก badge); ความมั่นใจของ arbiter อยู่ที่ judgeConfidence
        confidence: 'debate-resolved',
        judgeConfidence: judgeRes.confidence,
        models: [VERIFY_SOLVER_A, VERIFY_SOLVER_B],
        judgeModel: VERIFY_JUDGE,
        judgeUsed: true,
        solvers: [
          { model: VERIFY_SOLVER_A, choice: solveA.choice, rationale: solveA.rationale },
          { model: VERIFY_SOLVER_B, choice: solveB.choice, rationale: solveB.rationale }
        ],
        distractors: judgeRes.distractors || {},
        rationale: judgeRes.correctRationale || ''
      });
    } catch (err) {
      errors.push({ qid: q.qid, error: err.message || String(err) });
    }
  }

  return { result: 'success', verified: verified, errors: errors };
}

/**
 * Parse LLM JSON tolerantly — strip ```json fences, grab first {...},
 * and if truncated, still pull verifiedAnswer/choice via regex.
 * Mirrors parseGlossaryJson / geminiFetchOnce_ expectJson recovery.
 * @returns {Object|null}
 */
function parseVerifyJson_(raw) {
  if (raw == null) return null;
  var text = String(raw).replace(/```json/gi, '').replace(/```/g, '').trim();
  // Drop leading prose before first { (some models preamble then JSON)
  var start = text.indexOf('{');
  var end = text.lastIndexOf('}');
  if (start >= 0 && end > start) {
    try { return JSON.parse(text.slice(start, end + 1)); } catch (e) { /* fall through */ }
  }
  // Truncated JSON — recover the fields we must have for a decision
  var out = null;
  var mAns = text.match(/"verifiedAnswer"\s*:\s*(\d+)/);
  var mChoice = text.match(/"choice"\s*:\s*(\d+|"[A-Za-z]"|"\d+")/);
  var mRat = text.match(/"correctRationale"\s*:\s*"((?:[^"\\]|\\.)*)"/) ||
             text.match(/"rationale"\s*:\s*"((?:[^"\\]|\\.)*)"/);
  var mConf = text.match(/"confidence"\s*:\s*"(high|moderate)"/i);
  if (mAns || mChoice) {
    out = {};
    if (mAns) out.verifiedAnswer = Number(mAns[1]);
    if (mChoice) {
      var c = mChoice[1];
      out.choice = (c.charAt(0) === '"') ? c.slice(1, -1) : Number(c);
    }
    if (mRat) {
      try { out.correctRationale = out.rationale = JSON.parse('"' + mRat[1] + '"'); }
      catch (e2) { out.correctRationale = out.rationale = mRat[1]; }
    }
    if (mConf) out.confidence = mConf[1].toLowerCase();
    out.distractors = {};
  }
  return out;
}

/**
 * Call one solver model via IntelSphere executeChatbotQuery
 * @returns {Object} {choice: N, rationale: "..."}
 */
function solveWithModel_(q, model) {
  var prompt = buildVerifyPrompt_(q);
  var raw = executeChatbotQuery(prompt, model, 1); // attempt=1
  var parsed = parseVerifyJson_(raw.content);
  if (!parsed) {
    throw new Error('Solver ' + model + ' returned non-JSON: ' + String(raw.content || '').slice(0, 200));
  }
  var idx = normalizeChoiceIndex_(parsed.choice, q.choices.length);
  if (idx === null) {
    throw new Error('Solver ' + model + ' returned invalid choice: ' + parsed.choice);
  }
  return { model: model, choice: idx, rationale: String(parsed.rationale || '') };
}

/**
 * แปลงค่า choice ที่โมเดลคืนมาให้เป็น index 0-based
 * รับได้: number, ตัวเลขในรูป string ("2"), ตัวอักษร A/B/C
 * หมายเหตุ: ค่าที่อยู่ในช่วงอยู่แล้วจะไม่ถูกแตะ — เลข 1 แบบ 0-based กับแบบ 1-based
 * แยกจากกันไม่ได้ ถ้าเดาแล้วลบ 1 จะทำให้คำตอบที่ถูกอยู่แล้วกลายเป็นผิด
 * กรณีเดียวที่ยืนยันได้ว่าเป็น 1-based คือค่าเท่ากับจำนวนตัวเลือกพอดี (เกินช่วงบนไป 1)
 * @returns {number|null} index 0-based หรือ null ถ้าตีความไม่ได้
 */
function normalizeChoiceIndex_(raw, len) {
  var n = null;
  var fromLetter = false;
  if (typeof raw === 'number') {
    n = raw;
  } else if (typeof raw === 'string') {
    var t = raw.trim();
    if (/^[0-9]+$/.test(t)) {
      n = Number(t);
    } else if (/^[A-Za-z]$/.test(t)) {
      n = t.toUpperCase().charCodeAt(0) - 65; // A -> 0, B -> 1, C -> 2
      fromLetter = true;
    }
  }
  if (n === null || isNaN(n) || n !== Math.floor(n)) return null;
  if (n >= 0 && n < len) return n;
  // ตัวอักษรไม่มีปัญหา 0-based/1-based — 'E' ของข้อ 4 ตัวเลือกคือเกินช่วงจริง ไม่ใช่ off-by-one
  if (!fromLetter && n === len && len > 0) return n - 1; // 1-based ชัดเจน (4 ตัวเลือก แล้วตอบ 4)
  return null;
}

/**
 * Judge disagreement via Gemini (callGeminiAI, no Grounding)
 * @returns {Object} {verifiedAnswer: N, correctRationale: "...", distractors: {...}, confidence: "high"|"moderate"}
 */
function judgeDisagreement_(q, solveA, solveB) {
  var prompt = buildJudgePrompt_(q, solveA, solveB);
  // Use Gemini pool (ai-gemini.gs) — cheap flash tier
  var keyInfo = getAvailableAIKey("Gemini", VERIFY_JUDGE);
  if (!keyInfo) throw new Error('โควต้า Gemini หมดแล้วสำหรับวันนี้');
  var raw = callGeminiAI(prompt, keyInfo, null);
  var parsed = parseVerifyJson_(raw);
  if (!parsed) {
    throw new Error('Judge returned non-JSON: ' + String(raw || '').slice(0, 200));
  }
  // รับ number หรือ string ("4"/"E") — ผ่าน normalize เดียวกับ solver
  var ans = normalizeChoiceIndex_(
    (typeof parsed.verifiedAnswer !== 'undefined') ? parsed.verifiedAnswer : parsed.choice,
    q.choices.length
  );
  if (ans === null) {
    throw new Error('Judge returned invalid verifiedAnswer: ' + parsed.verifiedAnswer);
  }
  return {
    verifiedAnswer: ans,
    correctRationale: parsed.correctRationale || parsed.rationale || '',
    distractors: parsed.distractors || {},
    confidence: parsed.confidence || 'moderate'
  };
}

/** Build solver prompt (clinical reasoning style) */
function buildVerifyPrompt_(q) {
  var lines = [
    'You are a Medical Education Expert. Analyze this question and choose the single best answer.',
    '',
    'Question: ' + q.questionText,
    'Choices:'
  ];
  for (var i = 0; i < q.choices.length; i++) {
    lines.push(i + ') ' + q.choices[i]);
  }
  // Inject reported-discrepancy context from edit-modal flow (askMultiAIForEditModal)
  var hasReported = (q.suggestedAnswer !== undefined && q.suggestedAnswer !== null && q.suggestedAnswer !== '') ||
                    (q.reportDetail && String(q.reportDetail).trim() !== '');
  if (hasReported) {
    lines.push('');
    lines.push('Reported Discrepancy:');
    if (q.suggestedAnswer !== undefined && q.suggestedAnswer !== null && q.suggestedAnswer !== '') {
      lines.push('- DB answer: ' + q.correctAnswer + ' — Student suggested: ' + q.suggestedAnswer);
    }
    if (q.reportDetail && String(q.reportDetail).trim() !== '') {
      lines.push('- Student report detail: ' + q.reportDetail);
    }
    lines.push('Explicitly compare the DB answer against the student\'s suggested correction. If the student is correct, pick the student\'s choice and cite why the DB is wrong.');
  }
  lines.push('');
  lines.push('Provide:');
  lines.push('1. Your answer (choice number, 0-based).');
  lines.push('2. Clinical rationale (pathophysiology, differential, guidelines) — max 5 sentences.');
  lines.push('');
  lines.push('Write all rationales in Thai mixed with English medical terminology in a single continuous paragraph (no bullet points or newlines).');
  lines.push('CRITICAL: Return ONLY raw JSON. No markdown fences. No prose. Schema: {"choice":N,"rationale":"..."}');
  return lines.join('\n');
}

/** Build judge prompt (arbiter transcript) */
function buildJudgePrompt_(q, solveA, solveB) {
  var lines = [
    'You are a Clinical Arbiter. Two models solved the same medical question and disagreed.',
    '',
    'Question: ' + q.questionText,
    'Choices: ' + JSON.stringify(q.choices),
    '',
    'Model A (' + VERIFY_SOLVER_A + ') chose ' + solveA.choice + ': "' + solveA.rationale + '"',
    'Model B (' + VERIFY_SOLVER_B + ') chose ' + solveB.choice + ': "' + solveB.rationale + '"'
  ];
  // Inject reported-discrepancy context so judge weighs student's correction
  var hasReported = (q.suggestedAnswer !== undefined && q.suggestedAnswer !== null && q.suggestedAnswer !== '') ||
                    (q.reportDetail && String(q.reportDetail).trim() !== '');
  if (hasReported) {
    lines.push('');
    lines.push('Reported Discrepancy:');
    if (q.suggestedAnswer !== undefined && q.suggestedAnswer !== null && q.suggestedAnswer !== '') {
      lines.push('- DB answer: ' + q.correctAnswer + ' — Student suggested: ' + q.suggestedAnswer);
    }
    if (q.reportDetail && String(q.reportDetail).trim() !== '') {
      lines.push('- Student report detail: ' + q.reportDetail);
    }
    lines.push('Explicitly compare the DB answer against the student\'s suggested correction. If the student is correct, pick the student\'s choice and cite why the DB is wrong.');
  }
  lines.push('');
  lines.push('Which answer is clinically correct? Provide:');
  lines.push('1. Verified answer (choice number, 0-based).');
  lines.push('2. Why the correct answer is right (causal mechanism) — max 4 sentences.');
  lines.push('3. Why each wrong choice is a distractor — one short sentence each.');
  lines.push('4. Confidence: "high" (clear guideline) or "moderate" (clinical judgment call).');
  lines.push('');
  lines.push('Write all rationales in Thai mixed with English medical terminology in a single continuous paragraph (no bullet points or newlines).');
  lines.push('CRITICAL OUTPUT RULES:');
  lines.push('- Return ONLY a raw JSON object. No markdown. No ``` fences. No prose before/after.');
  lines.push('- Keep correctRationale under 600 characters and each distractor under 200 characters.');
  lines.push('- verifiedAnswer MUST be a bare number (e.g. 4), never a string.');
  lines.push('Schema:');
  lines.push('{"verifiedAnswer":N,"correctRationale":"...","distractors":{"0":"...","1":"..."},"confidence":"high"|"moderate"}');
  return lines.join('\n');
}
