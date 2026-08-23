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

// Hardcoded models for v1 (not user-selectable)
var VERIFY_SOLVER_A = "deepseek-v4-pro";
var VERIFY_SOLVER_B = "claude-sonnet-4.5";
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
    try {
      // 1) Solve with both models (LLM calls outside lock)
      var solveA = solveWithModel_(q, VERIFY_SOLVER_A);
      var solveB = solveWithModel_(q, VERIFY_SOLVER_B);

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
      var judgeRes = judgeDisagreement_(q, solveA, solveB);
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
 * Call one solver model via IntelSphere executeChatbotQuery
 * @returns {Object} {choice: N, rationale: "..."}
 */
function solveWithModel_(q, model) {
  var prompt = buildVerifyPrompt_(q);
  var raw = executeChatbotQuery(prompt, model, 1); // attempt=1
  var parsed;
  try {
    parsed = JSON.parse(raw.content);
  } catch (e) {
    throw new Error('Solver ' + model + ' returned non-JSON: ' + raw.content.slice(0, 200));
  }
  if (typeof parsed.choice !== 'number' || parsed.choice < 0 || parsed.choice >= q.choices.length) {
    throw new Error('Solver ' + model + ' returned invalid choice: ' + parsed.choice);
  }
  return { choice: parsed.choice, rationale: String(parsed.rationale || '') };
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
  var parsed;
  try {
    parsed = JSON.parse(raw);
  } catch (e) {
    throw new Error('Judge returned non-JSON: ' + raw.slice(0, 200));
  }
  if (typeof parsed.verifiedAnswer !== 'number') {
    throw new Error('Judge returned invalid verifiedAnswer');
  }
  return {
    verifiedAnswer: parsed.verifiedAnswer,
    correctRationale: parsed.correctRationale || '',
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
  lines.push('1. Your answer (choice number only).');
  lines.push('2. Clinical rationale (pathophysiology, differential diagnosis, guideline references).');
  lines.push('');
  lines.push('Format as JSON: {"choice": N, "rationale": "..."}');
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
  lines.push('1. Verified answer (choice number).');
  lines.push('2. Why the correct answer is right (causal mechanism).');
  lines.push('3. Why each wrong choice is a distractor (trap, edge case, outdated guideline).');
  lines.push('4. Confidence: "high" (clear guideline) or "moderate" (clinical judgment call).');
  lines.push('');
  lines.push('Format as JSON:');
  lines.push('{');
  lines.push('  "verifiedAnswer": N,');
  lines.push('  "correctRationale": "...",');
  lines.push('  "distractors": {"0": "...", "1": "...", ...},');
  lines.push('  "confidence": "high" | "moderate"');
  lines.push('}');
  return lines.join('\n');
}
