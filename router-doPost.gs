/* 
   =========================================
   ส่วนที่ 2: การบันทึกข้อมูล (POST)
   =========================================
*/
function doPost(e) {
  try {
    var doc = SpreadsheetApp.openById(SHEET_ID);
    var contents = e.postData.contents;
    var data = JSON.parse(contents);
    var action = data.action;

    // ----------------------------------------------------
    // LOCK-FREE GROUP (Processes instantly, no write queue overhead)
    // ----------------------------------------------------
    if (action === 'verifySession') {
      var userObj = verifySessionToken(data.sessionToken);
      if (userObj) {
        return ContentService.createTextOutput(JSON.stringify({
          'result': 'success',
          'user': userObj
        })).setMimeType(ContentService.MimeType.JSON);
      } else {
        return ContentService.createTextOutput(JSON.stringify({
          'result': 'error',
          'message': 'session_expired'
        })).setMimeType(ContentService.MimeType.JSON);
      }
    }

    // บริจาค/อัปเดต IntelSphere key หนึ่งใบ (idempotent by API_Key) — public donation form, lock-free
    // Rate limit ก่อน validate — กันคนใช้ endpoint นี้เป็นเครื่องเดา key (key-testing oracle)
    if (action === 'seedIntelSphereKey') {
      if (!checkActionRateLimit('rl_seed_', data.sessionToken || 'anon', 3)) {
        return ContentService.createTextOutput(JSON.stringify({
          result: 'error', message: 'คุณส่งคำขอบริจาคบ่อยเกินไป (สูงสุด 3 ครั้ง/ชั่วโมง) กรุณาลองใหม่ภายหลัง'
        })).setMimeType(ContentService.MimeType.JSON);
      }
      return seedIntelSphereKey(data.apiKey, data.donorName, data.notes);
    }

    // อ่าน catalog โมเดล IntelSphere + รายชื่อผู้บริจาค (read-only, ไม่ต้อง auth) — lock-free
    if (action === 'listModels') {
      return ContentService.createTextOutput(JSON.stringify({
        result: 'success',
        catalog: getIntelSphereModelCatalog(),
        donors: getDonorCredits()
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // agentQuery — owner-only agentic traffic จาก claude-kkuintelsphere-router proxy (lock-free: อ่าน + quota-column write แบบเดียวกับ askAIExpert)
    // แยกจาก askAIExpert เพราะ (1) checkRateLimit 15/hr ต่ำเกินไปมากสำหรับ agentic coding
    // (2) ต้องบังคับ session token (branch IntelSphere ของ askAIExpert เป็น public)
    // (3) priority เป็น Claude-first เพื่อคุณภาพ ไม่ใช่ Deepseek-first เพื่อ cost แบบ INTELSPHERE_PROVIDER_PRIORITY
    if (action === 'agentQuery') {
      // Auth: non-expiring owner secret (preferred) OR legacy 30-day session token.
      // The secret's hash is kept in Script Properties — see generateAgentQueryOwnerSecret().
      var agentAuthed = verifyAgentQueryOwnerSecret(data.ownerSecret) || !!verifySessionToken(data.sessionToken);
      if (!agentAuthed) {
        return ContentService.createTextOutput(JSON.stringify({
          result: 'error', message: 'session_expired'
        })).setMimeType(ContentService.MimeType.JSON);
      }
      try {
        var agentResult = executeAgentQuery(data.request || {});
        return ContentService.createTextOutput(JSON.stringify({
          result: 'success', provider: agentResult.provider, completion: agentResult.completion
        })).setMimeType(ContentService.MimeType.JSON);
      } catch (agentErr) {
        return ContentService.createTextOutput(JSON.stringify({
          result: 'error', message: agentErr.message
        })).setMimeType(ContentService.MimeType.JSON);
      }
    }

    if (action === 'askAIExpert') {
      // --- IntelSphere shared-pool branch: public, rate-limited, ไม่ใช้ admin auth ---
      if (data.provider === "IntelSphere") {
        var rlToken = data.sessionToken || "guest_user";
        if (!checkRateLimit(rlToken)) {
          return ContentService.createTextOutput(JSON.stringify({
            result: 'error',
            message: 'คุณส่งคำถามถึง AI เกินกำหนด (สูงสุด 15 ครั้งต่อชั่วโมง) กรุณารอสักครู่แล้วลองใหม่ เพื่อช่วยแบ่งโควต้าให้เพื่อนๆ ด้วยนะครับ'
          })).setMimeType(ContentService.MimeType.JSON);
        }

        var isModel = data.model; // ไม่มี default — frontend ส่งโมเดลที่เลือกจาก catalog เสมอ
        if (!isModel) {
          return ContentService.createTextOutput(JSON.stringify({
            result: 'error', message: 'กรุณาเลือกโมเดล AI ก่อนส่งคำถาม'
          })).setMimeType(ContentService.MimeType.JSON);
        }

        try {
          var aiResult = executeChatbotQuery(data.prompt, isModel, 1);
          return ContentService.createTextOutput(JSON.stringify({
            result: 'success', answer: aiResult.content, servedModel: aiResult.servedModel, switched: aiResult.switched
          })).setMimeType(ContentService.MimeType.JSON);
        } catch (isErr) {
          return ContentService.createTextOutput(JSON.stringify({
            result: 'error', message: isErr.message
          })).setMimeType(ContentService.MimeType.JSON);
        }
      }

      // --- เส้นทาง admin เดิม (Gemini pool) ---
      var userObj = null;
      if (data.sessionToken) {
        userObj = verifySessionToken(data.sessionToken);
      } else if (data.googleIdToken) {
        var payload = verifyGoogleToken(data.googleIdToken);
        if (payload) userObj = findAdminByEmail(payload.email);
      } else {
        userObj = verifyAdmin(data.username, data.adminPass);
      }

      if (!userObj) {
        return ContentService.createTextOutput(JSON.stringify({
          'result': 'error',
          'message': 'Session หมดอายุ กรุณาล็อกอินใหม่'
        })).setMimeType(ContentService.MimeType.JSON);
      }

      var provider = data.provider || "Gemini";
      var apiKeyInfo = getAvailableAIKey(provider);

      if (!apiKeyInfo) {
        return ContentService.createTextOutput(JSON.stringify({
          'result': 'error', 'message': 'ขณะนี้ไม่มี API Key ที่พร้อมใช้งาน (โควต้าเต็มทุก Key หรือยังไม่ได้ตั้งค่า)'
        })).setMimeType(ContentService.MimeType.JSON);
      }

      try {
        var aiResponse = callGeminiAI(data.prompt, apiKeyInfo, data.images);
        return ContentService.createTextOutput(JSON.stringify({
          'result': 'success',
          'answer': aiResponse,
          'quota': (apiKeyInfo.usage + 1) + "/" + apiKeyInfo.limit
        })).setMimeType(ContentService.MimeType.JSON);
      } catch (err) {
        return ContentService.createTextOutput(JSON.stringify({
          'result': 'error', 'message': err.message
        })).setMimeType(ContentService.MimeType.JSON);
      }
    }

    // Helper function for mapping batch logs to rows
    function toActivityRow(entry) {
      return [
        entry.timestamp ? new Date(entry.timestamp) : new Date(),
        entry.session || "N/A",
        entry.action || "",
        entry.target || "",
        entry.result || "",
        entry.timeSpent || 0,
        entry.metadata || ""
      ];
    }

    // ----------------------------------------------------
    // T0.1: batchLog — จัดการ "ก่อน" ขอ Lock ใดๆ เพื่อไม่ให้ analytics (write ถี่สุดของนักเรียน) ไปบล็อกโหวต/รายงาน
    // ใช้ appendRow ต่อแถว (atomic ในตัว ไม่ต้องพึ่ง LockService) แทน getRange(getLastRow()+1).setValues()
    // ----------------------------------------------------
    if (action === 'batchLog') {
      var logs = data.logs || [];
      if (logs.length > 0) {
        var activitySheet = doc.getSheetByName("UserActivity") || doc.insertSheet("UserActivity");
        if (activitySheet.getLastRow() === 0) {
          activitySheet.appendRow(["Timestamp", "SessionID", "Action", "TargetID", "Result", "TimeSpent", "Metadata"]);
          activitySheet.getRange(1, 1, 1, 7).setFontWeight("bold").setBackground("#e6f7ff");
        }
        for (var li = 0; li < logs.length; li++) {
          activitySheet.appendRow(toActivityRow(logs[li]));
        }
      }
      return ContentService.createTextOutput(JSON.stringify({ 'result': 'success' })).setMimeType(ContentService.MimeType.JSON);
    }

    // ----------------------------------------------------
    // submitAiFeedback — append-only เดี่ยวแบบเดียวกับ batchLog จึงอยู่ lock-free tier
    // rate limit 30/ชม. ต่อ token; เกิน limit ตอบ success (dropped) เพราะ frontend เป็น fire-and-forget
    // ----------------------------------------------------
    if (action === 'submitAiFeedback') {
      if (!checkActionRateLimit('rl_fb_', data.sessionToken || 'anon', 30)) {
        return ContentService.createTextOutput(JSON.stringify({ result: 'success', dropped: true }))
          .setMimeType(ContentService.MimeType.JSON);
      }
      return submitAiFeedbackRow(data);
    }

    // ----------------------------------------------------
    // §1.8 ingestKB — เขียน KB_Chunks (logged-in-only). auth + rate-limit ทำ "นอก lock" (อ่านล้วน)
    // เพื่อไม่ให้ garbage-token flood ไปแย่ง shared localized lock ของ vote/report และไม่ให้บังคับ
    // reject-log ไม่จำกัด; ล็อกเฉพาะช่วงเขียนจริง → การเขียนยังอยู่ localized-15s tier ตามแผน (ไม่มี UrlFetchApp ใต้ lock)
    // ----------------------------------------------------
    if (action === 'ingestKB') {
      // rate-limit ก่อน (keyed by token/'anon') — กัน flood ทั้ง lock และ reject-log
      if (!checkActionRateLimit('rl_kb_', data.sessionToken || 'anon', 10)) {
        return ContentService.createTextOutput(JSON.stringify({
          result: 'error', message: 'อัปโหลดบ่อยเกินไป (สูงสุด 10 ครั้ง/ชั่วโมง) กรุณาลองใหม่ภายหลัง'
        })).setMimeType(ContentService.MimeType.JSON);
      }
      // AUTH GATE (server-side, ก่อนเขียน KB_Chunks ทุกกรณี) — ต้องมี session token ที่ยัง valid
      var kbUser = verifySessionToken(data.sessionToken);
      if (!kbUser) {
        try {
          writeAdminLog(data.sessionToken ? 'INVALID_TOKEN' : 'NO_TOKEN', 'GUEST', 'KB', 'INGEST_REJECT',
            'KB_Chunks', 'Unauthenticated ingestKB attempt (subject=' + (data.subject || '') +
            ', source=' + (data.source || '') + ')', '', '', '');
        } catch (kbLogErr) {}
        return ContentService.createTextOutput(JSON.stringify({
          result: 'error', message: 'ต้องเข้าสู่ระบบก่อนอัปโหลด'
        })).setMimeType(ContentService.MimeType.JSON);
      }
      // ผ่าน auth แล้ว → ล็อก localized-15s เฉพาะช่วงเขียน
      var kbLock = LockService.getScriptLock();
      if (!kbLock.tryLock(15000)) {
        return ContentService.createTextOutput(JSON.stringify({
          result: 'error', message: 'เซิร์ฟเวอร์ไม่ตอบสนองเนื่องจากโหลดสูง (Lock Timeout)'
        })).setMimeType(ContentService.MimeType.JSON);
      }
      try {
        var kbRes = ingestKBChunks(data.subject, data.source, data.markdown, data.categoryId); // §1.9: categoryId optional
        return ContentService.createTextOutput(JSON.stringify(kbRes)).setMimeType(ContentService.MimeType.JSON);
      } finally {
        kbLock.releaseLock();
      }
    }

    // ----------------------------------------------------
    // §2.5 askGlossaryTerm — tap/select miss-path (public, self-populating). โครงเดียวกับ ingestKB:
    // rate-limit + LLM ทำ "นอก lock"; ล็อกเฉพาะช่วงเขียน 1 แถว. ***ห้ามยิง LLM ใต้ lock (advisor rule)***
    // ไม่อยู่ใน admin tier — guest ต้อง tap แปลได้ (นี่คือ UX หลักของ feature). dedup ก่อนยิง LLM = ประหยัดโทเคน
    // NOTE (shared-bucket tradeoff): guest ทุกคน key เป็น 'anon' → rl_glossary_anon เป็นถังเดียว "รวม" 20/ชม.
    //   ทั้ง guest ทุกคน; ผู้ล็อกอินได้ถังของตัวเอง (key=token). ยอมรับได้เพราะ hit path เป็น client-side ไม่จำกัด
    //   → feature degrade เป็น instant-lookup เมื่อ glossary แน่นขึ้น และผู้ใช้ KKU ส่วนใหญ่ล็อกอิน
    // ----------------------------------------------------
    if (action === 'askGlossaryTerm') {
      var glToken = data.sessionToken || 'anon';
      if (!checkActionRateLimit('rl_glossary_', glToken, GLOSSARY_ASK_RATE_LIMIT)) {
        return ContentService.createTextOutput(JSON.stringify({
          result: 'error', message: 'ค้นหาคำศัพท์บ่อยเกินไป (สูงสุด ' + GLOSSARY_ASK_RATE_LIMIT + ' ครั้ง/ชั่วโมง) กรุณาลองใหม่ภายหลัง'
        })).setMimeType(ContentService.MimeType.JSON);
      }
      var glWordKey = normalizeGlossaryTerm(data.word);
      if (!glWordKey) {
        return ContentService.createTextOutput(JSON.stringify({
          result: 'error', message: 'ไม่พบคำที่เลือก'
        })).setMimeType(ContentService.MimeType.JSON);
      }
      // (1) dedup ก่อนยิง LLM (อ่านผ่าน getGlossary cache — ไม่ getDataRange ทุก tap) → ไม่เปลืองโทเคนถ้ามีแล้ว
      var glHit = glossaryLookupCached(data.subject, glWordKey);
      if (glHit) {
        return ContentService.createTextOutput(JSON.stringify({
          result: 'success', term: glHit, cached: true
        })).setMimeType(ContentService.MimeType.JSON);
      }
      // (2) LLM lock-free (cheap flash tier) — ผ่าน IntelSphere path เดียวกับ askAIExpert
      var glParsed;
      try {
        var glRaw = executeChatbotQuery(buildGlossaryTermPrompt(data.word, data.sentence), GLOSSARY_MODEL, 1);
        glParsed = parseGlossaryJson(glRaw.content);
      } catch (glErr) {
        return ContentService.createTextOutput(JSON.stringify({
          result: 'error', message: (glErr && glErr.message) || 'AI ไม่พร้อมใช้งาน กรุณาลองใหม่'
        })).setMimeType(ContentService.MimeType.JSON);
      }
      if (!glParsed || !glParsed.en) {
        return ContentService.createTextOutput(JSON.stringify({
          result: 'error', message: 'ไม่สามารถแปลคำนี้ได้ กรุณาลองใหม่'
        })).setMimeType(ContentService.MimeType.JSON);
      }
      // (2b) dedup post-LLM บน canonical en (เผื่อโมเดล normalize ไปตรงแถวที่มีอยู่แล้ว) — ยังไม่ต้อง lock
      var glCanonHit = glossaryLookupCached(data.subject, normalizeGlossaryTerm(glParsed.en));
      if (glCanonHit) {
        return ContentService.createTextOutput(JSON.stringify({
          result: 'success', term: glCanonHit, cached: true
        })).setMimeType(ContentService.MimeType.JSON);
      }
      // (3) เขียน 1 แถวใต้ localized-15s lock (ล็อกเฉพาะการเขียน; re-check dedup race-safe ข้างใน)
      var glLock = LockService.getScriptLock();
      if (!glLock.tryLock(15000)) {
        return ContentService.createTextOutput(JSON.stringify({
          result: 'error', message: 'เซิร์ฟเวอร์ไม่ตอบสนองเนื่องจากโหลดสูง (Lock Timeout)'
        })).setMimeType(ContentService.MimeType.JSON);
      }
      try {
        var glWrite = writeGlossaryRowLocked(data.subject, glParsed, data.questionId);
        return ContentService.createTextOutput(JSON.stringify({
          result: 'success', term: glWrite.term, cached: glWrite.cached
        })).setMimeType(ContentService.MimeType.JSON);
      } finally {
        glLock.releaseLock();
      }
    }

    // ----------------------------------------------------
    // §3.6 generateHighYield — lazy-generate-then-cache miss-path (public, self-populating). โครงเดียวกับ askGlossaryTerm:
    // rate-limit → dedup(cache) ก่อนยิง LLM → LLM ทำ "นอก lock" → เขียน 1 แถวใต้ localized-15s lock. ***ห้ามยิง LLM ใต้ lock***
    // ไม่อยู่ใน admin tier — guest กดสร้างชีทสรุปได้ (เป็น UX หลักของ feature). subject resolve จาก categoryId ฝั่ง server
    // ----------------------------------------------------
    if (action === 'generateHighYield') {
      var hyToken = data.sessionToken || 'anon';
      if (!checkActionRateLimit('rl_highyield_', hyToken, HIGHYIELD_GEN_RATE_LIMIT)) {
        return ContentService.createTextOutput(JSON.stringify({
          result: 'error', message: 'สร้างชีทสรุปบ่อยเกินไป (สูงสุด ' + HIGHYIELD_GEN_RATE_LIMIT + ' ครั้ง/ชั่วโมง) กรุณาลองใหม่ภายหลัง'
        })).setMimeType(ContentService.MimeType.JSON);
      }
      var hyCat = String(data.category || '').trim();
      if (!hyCat) {
        return ContentService.createTextOutput(JSON.stringify({
          result: 'error', message: 'ไม่พบหัวข้อ (category)'
        })).setMimeType(ContentService.MimeType.JSON);
      }
      // (1) dedup ก่อนยิง LLM (อ่านผ่าน getHighYield cache) → ไม่เปลืองโทเคนถ้ามีแล้ว
      var hyHit = highYieldLookupCached(hyCat);
      if (hyHit) {
        return ContentService.createTextOutput(JSON.stringify({
          result: 'success', highyield: hyHit, cached: true
        })).setMimeType(ContentService.MimeType.JSON);
      }
      // (2) รวมข้อสอบของหมวดนี้ (อ่านอย่างเดียว, lock-free) — subject resolve จาก categoryId ฝั่ง server (fallback = data.subject)
      var hyText = aggregateCategoryText(hyCat, data.subject);
      if (!hyText) {
        return ContentService.createTextOutput(JSON.stringify({
          result: 'error', message: 'ไม่พบข้อสอบในหัวข้อนี้ จึงยังสร้างชีทสรุปไม่ได้'
        })).setMimeType(ContentService.MimeType.JSON);
      }
      // (3) LLM lock-free (cheap flash tier, max_tokens สูงกว่า default กัน JSON ขาด) — ผ่าน IntelSphere path เดียวกับ askAIExpert
      var hyParsed;
      try {
        var hyRaw = executeChatbotQuery(buildHighYieldPrompt(hyText), HIGHYIELD_MODEL, 1, HIGHYIELD_MAX_TOKENS);
        hyParsed = parseHighYieldJson(hyRaw.content);
      } catch (hyErr) {
        return ContentService.createTextOutput(JSON.stringify({
          result: 'error', message: (hyErr && hyErr.message) || 'AI ไม่พร้อมใช้งาน กรุณาลองใหม่'
        })).setMimeType(ContentService.MimeType.JSON);
      }
      if (!hyParsed || !hyParsed.summary_md) {
        return ContentService.createTextOutput(JSON.stringify({
          result: 'error', message: 'สร้างชีทสรุปไม่สำเร็จ (คำตอบ AI ไม่สมบูรณ์) กรุณาลองใหม่'
        })).setMimeType(ContentService.MimeType.JSON);
      }
      // (4) เขียน 1 แถวใต้ localized-15s lock (ล็อกเฉพาะการเขียน; re-check dedup race-safe ข้างใน)
      var hyLock = LockService.getScriptLock();
      if (!hyLock.tryLock(15000)) {
        return ContentService.createTextOutput(JSON.stringify({
          result: 'error', message: 'เซิร์ฟเวอร์ไม่ตอบสนองเนื่องจากโหลดสูง (Lock Timeout)'
        })).setMimeType(ContentService.MimeType.JSON);
      }
      try {
        var hyWrite = writeHighYieldRowLocked(hyCat, hyParsed, getVersionCached());
        return ContentService.createTextOutput(JSON.stringify({
          result: 'success', highyield: hyWrite.row, cached: hyWrite.cached
        })).setMimeType(ContentService.MimeType.JSON);
      } finally {
        hyLock.releaseLock();
      }
    }

    // ----------------------------------------------------
    // §3.4 voteHighYieldMnemonic — 👍/🚩 mnemonic (public). localized-15s tier (standalone block, mirror askGlossaryTerm write-lock).
    // อ่านคอลัมน์ H (Mnemonic_Votes JSON) "สดจากชีตใต้ lock" → mutate {idx:net+delta} → เขียนกลับ → invalidate cache.
    // ***ห้ามอ่านจาก cache*** (สอง user โหวต mnemonic คนละตัวของหมวดเดียวกันพร้อมกันจะทับกัน). re-vote guard เป็น localStorage ฝั่ง client
    // ----------------------------------------------------
    if (action === 'voteHighYieldMnemonic') {
      var mvToken = data.sessionToken || 'anon';
      if (!checkActionRateLimit('rl_hymvote_', mvToken, HIGHYIELD_VOTE_RATE_LIMIT)) {
        return ContentService.createTextOutput(JSON.stringify({
          result: 'error', message: 'โหวตบ่อยเกินไป กรุณาลองใหม่ภายหลัง'
        })).setMimeType(ContentService.MimeType.JSON);
      }
      var mvCat = String(data.category || '').trim();
      var mvIdx = parseInt(data.mnemonicIdx, 10);
      var mvDelta = (parseInt(data.delta, 10) < 0) ? -1 : 1; // clamp เป็น ±1
      if (!mvCat || isNaN(mvIdx) || mvIdx < 0) {
        return ContentService.createTextOutput(JSON.stringify({
          result: 'error', message: 'ข้อมูลโหวตไม่ถูกต้อง'
        })).setMimeType(ContentService.MimeType.JSON);
      }
      var mvLock = LockService.getScriptLock();
      if (!mvLock.tryLock(15000)) {
        return ContentService.createTextOutput(JSON.stringify({
          result: 'error', message: 'เซิร์ฟเวอร์ไม่ตอบสนองเนื่องจากโหลดสูง (Lock Timeout)'
        })).setMimeType(ContentService.MimeType.JSON);
      }
      try {
        var mvRes = voteHighYieldMnemonicLocked(mvCat, mvIdx, mvDelta); // อ่านสด+เขียนใต้ lock
        return ContentService.createTextOutput(JSON.stringify(mvRes)).setMimeType(ContentService.MimeType.JSON);
      } finally {
        mvLock.releaseLock();
      }
    }

    // ----------------------------------------------------
    // LOCALIZED LOCK GROUP (Locks briefly for writes, tryLock 15s)
    // ----------------------------------------------------
    var localizedActions = ['submitVote', 'submitReport', 'voteOnReport', 'deleteSession'];
    if (localizedActions.indexOf(action) > -1) {
      var lock = LockService.getScriptLock();
      var acquired = lock.tryLock(15000);
      if (!acquired) {
        return ContentService.createTextOutput(JSON.stringify({
          'result': 'error',
          'message': 'เซิร์ฟเวอร์ไม่ตอบสนองเนื่องจากโหลดสูง (Lock Timeout)'
        })).setMimeType(ContentService.MimeType.JSON);
      }
      var lockReleased = false; // T0.2: ให้ปลด Lock เองก่อนงานหนัก (reconciliation/Gemini) โดย finally ไม่ปลดซ้ำ
      try {
        if (action === 'deleteSession') {
          var token = data.sessionToken;
          if (token) {
            var ss = SpreadsheetApp.openById(SHEET_ID);
            var sheet = ss.getSheetByName("Sessions");
            if (sheet) {
              var rows = sheet.getDataRange().getValues();
              for (var i = rows.length - 1; i >= 1; i--) {
                if (rows[i][0] === token) {
                  sheet.deleteRow(i + 1);
                  break;
                }
              }
            }
          }
          return ContentService.createTextOutput(JSON.stringify({ 'result': 'success' })).setMimeType(ContentService.MimeType.JSON);
        }

        if (action === 'submitVote') {
          var voteSheet = doc.getSheetByName("Votes") || doc.insertSheet("Votes");
          var voteData = voteSheet.getDataRange().getValues();
          var suggestedCategory = data.suggestedCategory || [];
          var delta = data.delta || 1;
          var timestamp = new Date();

          suggestedCategory.forEach(function (category) {
            var foundRowIndex = -1;
            for (var j = 1; j < voteData.length; j++) {
              if (voteData[j][0] == data.questionId && voteData[j][2] == category) {
                foundRowIndex = j + 1;
                break;
              }
            }

            if (foundRowIndex !== -1) {
              var currentVote = parseInt(voteSheet.getRange(foundRowIndex, 4).getValue()) || 0;
              var newVote = currentVote + delta;

              if (newVote < 0) {
                voteSheet.deleteRow(foundRowIndex);
              } else {
                voteSheet.getRange(foundRowIndex, 4).setValue(newVote);
                voteSheet.getRange(foundRowIndex, 5).setValue(timestamp);
              }
            } else if (delta > 0) {
              voteSheet.appendRow([data.questionId, data.questionText, category, 1, timestamp, "Pending"]);
            }
          });
          updateVotesVersion();
          // T0.2: ปลด Lock ก่อนรัน processVotes() (full-sheet reconciliation + sort + updateVersion)
          // เพื่อไม่ให้ค้าง Lock ระหว่างงานหนัก — processVotes ไม่มีการเรียก UrlFetchApp จึงปลอดภัย
          lock.releaseLock();
          lockReleased = true;
          processVotes();
          return ContentService.createTextOutput(JSON.stringify({ 'result': 'success' })).setMimeType(ContentService.MimeType.JSON);
        }

        if (action === 'submitReport') {
          var sheet = doc.getSheetByName("Report") || doc.insertSheet("Report");
          if (sheet.getLastRow() == 0) {
            sheet.appendRow(["From", "Category", "QuestionID", "Question", "Image", "Choices", "SuggestedAnswer", "ReportDetail", "Time", "Status", "AdminNote", "Done", "SuggestedExplain", "VoteCount"]);
          } else {
            var headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
            if (headerRow.indexOf("SuggestedExplain") === -1) {
              sheet.getRange(1, 13).setValue("SuggestedExplain");
              sheet.getRange(1, 14).setValue("VoteCount");
            }
          }

          var qImg = (data.questionImages && data.questionImages.indexOf("http") === 0) ? data.questionImages.split("///")[0] : "";
          var ansSug = data.suggestedChoice || "";

          sheet.appendRow([
            data.from || "User",
            data.category || "",
            data.questionId || "",
            data.question || "",
            qImg,
            data.allChoices || "",
            ansSug,
            data.report || "",
            new Date().toISOString(),
            "Pending",
            "",
            "FALSE",
            data.suggestedExplain || "",
            1
          ]);

          updateVotesVersion();
          return ContentService.createTextOutput(JSON.stringify({ 'result': 'success' })).setMimeType(ContentService.MimeType.JSON);
        }

        if (action === 'voteOnReport') {
          var reportSheet = doc.getSheetByName("Report");
          if (!reportSheet) {
            return ContentService.createTextOutput(JSON.stringify({ result: 'error', message: 'Report sheet not found' })).setMimeType(ContentService.MimeType.JSON);
          }
          var rv = reportSheet.getDataRange().getValues();
          var targetTs = String(data.reportTimestamp || "").trim();
          var delta = parseInt(data.delta) || 1;
          var foundIdx = -1;
          for (var i = 1; i < rv.length; i++) {
            var sTime = rv[i][8] instanceof Date ? rv[i][8].toISOString() : String(rv[i][8]);
            if (sTime.trim() === targetTs) { foundIdx = i; break; }
          }
          if (foundIdx === -1) {
            return ContentService.createTextOutput(JSON.stringify({ result: 'error', message: 'Report not found' })).setMimeType(ContentService.MimeType.JSON);
          }

          // เขียนคะแนนโหวตใหม่ (งานเบาภายใต้ Lock)
          var newVotes = Math.max(0, (parseInt(rv[foundIdx][13]) || 0) + delta);
          reportSheet.getRange(foundIdx + 1, 14).setValue(newVotes);
          updateVotesVersion();

          // T0.2: threshold check เฉพาะแถวนี้ — รัน processReports (ซึ่งอาจเรียก Gemini/UrlFetchApp) เฉพาะเมื่อ
          // รายงานนี้ยัง Pending และแตะเกณฑ์แล้วเท่านั้น และต้องทำ "นอก Lock" เสมอ (ห้ามเรียก UrlFetchApp ใต้ LockService)
          var reportStatus = String(rv[foundIdx][9]).trim();
          var shouldProcess = (reportStatus === "Pending" && newVotes >= REPORT_VOTE_THRESHOLD);

          lock.releaseLock();
          lockReleased = true;
          if (shouldProcess) {
            processReports(doc);
          }
          return ContentService.createTextOutput(JSON.stringify({ result: 'success', newVoteCount: newVotes })).setMimeType(ContentService.MimeType.JSON);
        }
      } finally {
        if (!lockReleased) lock.releaseLock(); // อาจถูกปลดไปแล้วใน submitVote/voteOnReport
      }
    }

    // ----------------------------------------------------
    // ADMIN LOCK GROUP (Admin and write operations, tryLock 25s)
    // ----------------------------------------------------
    var adminLock = LockService.getScriptLock();
    var adminAcquired = adminLock.tryLock(25000);
    if (!adminAcquired) {
      return ContentService.createTextOutput(JSON.stringify({
        'result': 'error',
        'message': 'ระบบหลังบ้านทำงานหนักเนื่องจากมีการเขียนซ้อนกัน (Admin Lock Timeout)'
      })).setMimeType(ContentService.MimeType.JSON);
    }
    try {
      if (action === 'checkGoogleAuth') {
        var tokenPayload = verifyGoogleToken(data.idToken);
        if (!tokenPayload) {
          return ContentService.createTextOutput(JSON.stringify({
            'result': 'error',
            'message': 'Token ไม่ถูกต้องหรือหมดอายุการใช้งาน'
          })).setMimeType(ContentService.MimeType.JSON);
        }

        var email = tokenPayload.email;
        var hd = tokenPayload.hd;

        if (hd !== "kkumail.com" && hd !== "kku.ac.th") {
          return ContentService.createTextOutput(JSON.stringify({
            'result': 'error',
            'message': 'ต้องใช้บัญชี @kkumail.com หรือ @kku.ac.th ของทางมหาวิทยาลัยเท่านั้น'
          })).setMimeType(ContentService.MimeType.JSON);
        }

        var adminUser = findAdminByEmail(email);
        if (adminUser) {
          var sessionToken = createSession(email, adminUser);
          writeAdminLog(adminUser.displayName, adminUser.role, "AUTH", "LOGIN_SSO", "Session", "Google SSO Login Success", "", "", "");
          return ContentService.createTextOutput(JSON.stringify({
            'result': 'success',
            'user': adminUser,
            'sessionToken': sessionToken
          })).setMimeType(ContentService.MimeType.JSON);
        } else {
          writeAdminLog(email, "GUEST", "AUTH", "LOGIN_SSO_FAIL", "Session", "Google login blocked: Email not in whitelist", "", "", "");
          return ContentService.createTextOutput(JSON.stringify({
            'result': 'error',
            'message': 'บัญชีผู้ใช้นี้ไม่มีอยู่ในสิทธิ์การแก้ไขระบบ กรุณาติดต่อผู้ดูแลเพื่อเพิ่มรายชื่ออีเมลของคุณ'
          })).setMimeType(ContentService.MimeType.JSON);
        }
      }

      if (action === 'registerAdmin') {
        var sheet = doc.getSheetByName("Admins");
        var users = sheet.getDataRange().getValues();

        for (var i = 1; i < users.length; i++) {
          if (users[i][0] == data.userData.Username) {
            return ContentService.createTextOutput(JSON.stringify({ 'result': 'error', 'message': 'Username นี้ถูกใช้ไปแล้ว' })).setMimeType(ContentService.MimeType.JSON);
          }
          if (users[i][8] == data.userData.StudentID) {
            return ContentService.createTextOutput(JSON.stringify({ 'result': 'error', 'message': 'รหัสนักศึกษานี้ลงทะเบียนแล้ว' })).setMimeType(ContentService.MimeType.JSON);
          }
        }

        var avatarUrl = "https://api.dicebear.com/7.x/avataaars/svg?seed=" + data.userData.Username;
        if (data.userData.AvatarBase64) {
          try {
            avatarUrl = uploadToDrive(data.userData.AvatarBase64, data.userData.Username + "_avatar.png", "image/png");
          } catch (err) {
            return ContentService.createTextOutput(JSON.stringify({ 'result': 'error', 'message': 'Upload รูปไม่สำเร็จ: ' + err.message })).setMimeType(ContentService.MimeType.JSON);
          }
        }

        sheet.appendRow([
          data.userData.Username,
          data.userData.Password,
          data.userData.DisplayName,
          avatarUrl,
          data.userData.Role || "Admin",
          data.userData.KKUMail,
          data.userData.Prefix,
          data.userData.FullName,
          data.userData.StudentID,
          data.userData.Year,
          data.userData.Contact
        ]);

        updateVersion();
        writeAdminLog("System", "SYSTEM", "AUTH", "REGISTER", data.userData.Username, "New Admin Registered", "", "", "");

        return ContentService.createTextOutput(JSON.stringify({ 'result': 'success' })).setMimeType(ContentService.MimeType.JSON);
      }

      if (action === 'resetPassword') {
        var sheet = doc.getSheetByName("Admins");
        var users = sheet.getDataRange().getValues();
        var found = false;

        for (var i = 1; i < users.length; i++) {
          if (users[i][0] == data.verifyData.Username &&
            users[i][5] == data.verifyData.KKUMail &&
            users[i][8] == data.verifyData.StudentID &&
            users[i][7] == data.verifyData.FullName) {

            sheet.getRange(i + 1, 2).setValue(data.newPassword);
            updateVersion();
            writeAdminLog(data.verifyData.Username, "USER", "AUTH", "RESET_PWD", "Self", "Password Changed via Verification", "", "", "");
            found = true;
            break;
          }
        }

        if (found) {
          return ContentService.createTextOutput(JSON.stringify({ 'result': 'success' })).setMimeType(ContentService.MimeType.JSON);
        } else {
          return ContentService.createTextOutput(JSON.stringify({ 'result': 'error', 'message': 'ข้อมูลยืนยันตัวตนไม่ถูกต้อง' })).setMimeType(ContentService.MimeType.JSON);
        }
      }

      if (action === 'checkAuth') {
        var userObj = verifyAdmin(data.username, data.password);
        if (userObj) {
          writeAdminLog(userObj.username, userObj.role, "AUTH", "LOGIN", "Session", "Login Success", "", "", data.metadata || "");
          return ContentService.createTextOutput(JSON.stringify({
            'result': 'success',
            'user': userObj
          })).setMimeType(ContentService.MimeType.JSON);
        } else {
          writeAdminLog(data.username || "Unknown", "GUEST", "AUTH", "LOGIN_FAIL", "Session", "Login Failed", "", "", data.metadata || "");
          return ContentService.createTextOutput(JSON.stringify({
            'result': 'error',
            'message': 'Username หรือ Password ไม่ถูกต้อง'
          })).setMimeType(ContentService.MimeType.JSON);
        }
      }

      if (action === 'updateAdminProfile') {
        try {
          var ss = SpreadsheetApp.openById(SHEET_ID);
          var sheet = ss.getSheetByName("Admins");
          if (!sheet) throw new Error("ไม่พบแผ่นงาน 'Admins'");

          var values = sheet.getDataRange().getValues();
          var targetUsername = data.targetUsername ? data.targetUsername.toString().trim() : "";
          var foundRow = -1;
          var oldProfileData = {};

          for (var i = 1; i < values.length; i++) {
            if (values[i][0].toString().trim().toLowerCase() === targetUsername.toLowerCase()) {
              foundRow = i + 1;
              oldProfileData = {
                displayName: values[i][2],
                avatarUrl: values[i][3],
                prefix: values[i][6],
                fullName: values[i][7],
                year: values[i][9],
                contact: values[i][10]
              };
              break;
            }
          }

          if (foundRow !== -1) {
            var u = data.updateData;
            if (u.displayName !== undefined) sheet.getRange(foundRow, 3).setValue(u.displayName);
            if (u.avatarUrl !== undefined) sheet.getRange(foundRow, 4).setValue(u.avatarUrl);
            if (u.prefix !== undefined) sheet.getRange(foundRow, 7).setValue(u.prefix);
            if (u.fullName !== undefined) sheet.getRange(foundRow, 8).setValue(u.fullName);
            if (u.year !== undefined) sheet.getRange(foundRow, 10).setValue(u.year);
            if (u.contact !== undefined) sheet.getRange(foundRow, 11).setValue(u.contact);

            updateVersion();
            writeAdminLog(data.username || targetUsername, "ADMIN", "PROFILE", "UPDATE", targetUsername, "Updated profile details", oldProfileData, u, "");

            return ContentService.createTextOutput(JSON.stringify({
              'result': 'success',
              'message': 'Profile updated successfully'
            })).setMimeType(ContentService.MimeType.JSON);
          } else {
            throw new Error("ไม่พบชื่อผู้ใช้: " + targetUsername);
          }
        } catch (error) {
          return ContentService.createTextOutput(JSON.stringify({
            'result': 'error',
            'message': error.toString()
          })).setMimeType(ContentService.MimeType.JSON);
        }
      }

      if (action === 'uploadImage') {
        var userObj = null;
        if (data.sessionToken) {
          userObj = verifySessionToken(data.sessionToken);
        } else if (data.googleIdToken) {
          var payload = verifyGoogleToken(data.googleIdToken);
          if (payload) userObj = findAdminByEmail(payload.email);
        } else {
          userObj = verifyAdmin(data.username, data.adminPass);
        }

        if (!userObj) {
          return ContentService.createTextOutput(JSON.stringify({
            'result': 'error',
            'message': 'token_expired'
          })).setMimeType(ContentService.MimeType.JSON);
        }

        try {
          var subject = data.data.subject;
          var year = data.data.year;
          var fileUrl = uploadQuestionImageToDrive(data.data.base64, data.data.questionId, data.data.type, subject, year);
          writeAdminLog(userObj.username, userObj.role, "IMAGE", "UPLOAD", data.data.questionId, "Uploaded new " + data.data.type + " image", "", fileUrl, "");

          return ContentService.createTextOutput(JSON.stringify({
            'result': 'success',
            'url': fileUrl
          })).setMimeType(ContentService.MimeType.JSON);

        } catch (err) {
          return ContentService.createTextOutput(JSON.stringify({
            'result': 'error',
            'message': 'Drive Upload Error: ' + err.message
          })).setMimeType(ContentService.MimeType.JSON);
        }
      }

      // T2.5: อัปโหลดหลายรูปในการเรียกครั้งเดียว (สูงสุด 10 รูป) — auth และพารามิเตอร์ต่อรายการเหมือน uploadImage
      // แต่ละรายการอยู่ใน data.images[] = { base64, questionId, type, subject, year }
      // คืน urls[] เรียงตามลำดับ input; รายการที่ล้มเหลวจะเป็น { error: "..." } (ไม่ทำให้ทั้ง batch ล้ม)
      if (action === 'uploadImagesBatch') {
        var userObj = null;
        if (data.sessionToken) {
          userObj = verifySessionToken(data.sessionToken);
        } else if (data.googleIdToken) {
          var payload = verifyGoogleToken(data.googleIdToken);
          if (payload) userObj = findAdminByEmail(payload.email);
        } else {
          userObj = verifyAdmin(data.username, data.adminPass);
        }

        if (!userObj) {
          return ContentService.createTextOutput(JSON.stringify({
            'result': 'error',
            'message': 'token_expired'
          })).setMimeType(ContentService.MimeType.JSON);
        }

        var images = data.images || [];
        if (!Array.isArray(images) || images.length === 0) {
          return ContentService.createTextOutput(JSON.stringify({
            'result': 'error',
            'message': 'ไม่พบรายการรูปภาพ (images array is empty)'
          })).setMimeType(ContentService.MimeType.JSON);
        }
        if (images.length > 10) {
          return ContentService.createTextOutput(JSON.stringify({
            'result': 'error',
            'message': 'อัปโหลดได้สูงสุด 10 รูปต่อครั้ง (batch size exceeds 10)'
          })).setMimeType(ContentService.MimeType.JSON);
        }

        var urls = [];
        var successCount = 0;
        for (var bi = 0; bi < images.length; bi++) {
          var item = images[bi] || {};
          try {
            if (!item.base64) { urls.push({ error: 'missing base64' }); continue; }
            var fileUrl = uploadQuestionImageToDrive(item.base64, item.questionId, item.type, item.subject, item.year);
            urls.push(fileUrl);
            successCount++;
          } catch (err) {
            urls.push({ error: err.message });
          }
        }

        writeAdminLog(userObj.username, userObj.role, "IMAGE", "UPLOAD_BATCH",
          (images[0] && images[0].questionId) || "",
          "Batch uploaded " + successCount + "/" + images.length + " images", "", "", "");

        return ContentService.createTextOutput(JSON.stringify({
          'result': 'success',
          'urls': urls
        })).setMimeType(ContentService.MimeType.JSON);
      }

      if (action === 'deleteImage') {
        var userObj = verifyUser(data);
        if (!userObj) {
          return ContentService.createTextOutput(JSON.stringify({ 'result': 'error', 'message': 'Session หมดอายุ หรือไม่ได้รับอนุญาตให้เข้าถึง' })).setMimeType(ContentService.MimeType.JSON);
        }

        try {
          var fileUrl = data.data.url;
          var currentQid = String(data.data.currentQid || "").trim();

          var match = fileUrl.match(/id=([^&]+)/) || fileUrl.match(/\/d\/([^\/]+)/);
          if (!match) throw new Error("ไม่สามารถระบุ ID ของไฟล์จาก URL นี้ได้");
          var fileId = match[1];

          var qSheet = doc.getSheetByName("Questions");
          var qData = qSheet.getDataRange().getValues();
          var otherUsages = [];

          for (var i = 1; i < qData.length; i++) {
            var qid = String(qData[i][0]).trim();
            if (qid === currentQid) continue;

            var imgCol = String(qData[i][2]);
            var choiceCol = String(qData[i][3]);

            if (imgCol.indexOf(fileId) !== -1 || choiceCol.indexOf(fileId) !== -1) {
              otherUsages.push({ qid: qid, type: imgCol.indexOf(fileId) !== -1 ? 'Main' : 'Choice' });
            }
          }

          if (otherUsages.length > 0) {
            var nextOwner = otherUsages[0];
            var file = DriveApp.getFileById(fileId);
            var timestamp = new Date().getTime();
            var originalName = file.getName();
            var ext = originalName.substring(originalName.lastIndexOf('.')) || ".png";
            var newName = "Q_" + nextOwner.qid + "_" + nextOwner.type + "_" + timestamp + ext;
            file.setName(newName);

            writeAdminLog(userObj.username, userObj.role, "IMAGE", "TRANSFER", fileId, "ภาพยังถูกใช้โดยข้อ " + nextOwner.qid + " จึงแค่เปลี่ยนชื่อไฟล์", "", "", "");

            return ContentService.createTextOutput(JSON.stringify({
              'result': 'success', 'message': 'ปลดลิงก์สำเร็จ (ไฟล์ยังคงอยู่เพราะข้อ ' + nextOwner.qid + ' ใช้งานอยู่)'
            })).setMimeType(ContentService.MimeType.JSON);

          } else {
            var resultMessage = deleteImageToRecycleBin(fileUrl, userObj.username);
            writeAdminLog(userObj.username, userObj.role, "IMAGE", "DELETE", fileId, "ย้ายรูปลง Recycle Bin", "", "TRASHED", "");

            return ContentService.createTextOutput(JSON.stringify({
              'result': 'success', 'message': resultMessage
            })).setMimeType(ContentService.MimeType.JSON);
          }

        } catch (err) {
          return ContentService.createTextOutput(JSON.stringify({ 'result': 'error', 'message': 'ลบรูปภาพขัดข้อง: ' + err.message })).setMimeType(ContentService.MimeType.JSON);
        }
      }

      if (action === 'restoreImage') {
        var userObj = verifyUser(data.username, data.adminPass);
        if (!userObj) {
          return ContentService.createTextOutput(JSON.stringify({
            'result': 'error',
            'message': 'Session หมดอายุ หรือสิทธิ์ไม่ถูกต้อง กรุณาล็อกอินใหม่'
          })).setMimeType(ContentService.MimeType.JSON);
        }

        try {
          var fileUrl = data.data.url;
          var resultMessage = restoreImageFromRecycleBin(fileUrl);
          writeAdminLog(userObj.username, userObj.role, "IMAGE", "RESTORE", fileUrl, "Restored image from Recycle Bin", "TRASHED", "RESTORED", "");

          return ContentService.createTextOutput(JSON.stringify({
            'result': 'success',
            'url': fileUrl,
            'message': resultMessage
          })).setMimeType(ContentService.MimeType.JSON);

        } catch (err) {
          return ContentService.createTextOutput(JSON.stringify({
            'result': 'error',
            'message': 'ไม่สามารถกู้คืนรูปภาพได้ (อาจจะไม่มีอยู่ในถังขยะแล้ว): ' + err.message
          })).setMimeType(ContentService.MimeType.JSON);
        }
      }

      var adminActions = ['editQuestion', 'deleteQuestion', 'addCategory', 'adminImport', 'updateReportStatus', 'deleteCategory', 'updateCategory', 'deleteGroup', 'updateAccordionGroup', 'addSubject', 'updateSubject', 'deleteSubject', 'addAnnouncement', 'editAnnouncement', 'deleteAnnouncement', 'runRelationsBatchManual', 'runGlossaryBatchManual', 'runHighYieldBatchManual', 'runKeywordIndexBatchManual', 'bulkAddQuestionCategories'];
      if (adminActions.indexOf(action) > -1) {
        var userObj = null;
        if (data.sessionToken) {
          userObj = verifySessionToken(data.sessionToken);
        } else if (data.googleIdToken) {
          var payload = verifyGoogleToken(data.googleIdToken);
          if (payload) userObj = findAdminByEmail(payload.email);
        } else {
          userObj = verifyAdmin(data.username, data.adminPass);
        }

        if (!userObj) {
          return ContentService.createTextOutput(JSON.stringify({
            'result': 'error',
            'message': 'token_expired'
          })).setMimeType(ContentService.MimeType.JSON);
        }

        var user = userObj.displayName || "Unknown Admin";
        var userRole = userObj.role || "Admin";
        var metadata = data.metadata || "";

        var sheet;

        // Feature 4: สั่งสร้าง relations ของวิชาเดียวแบบ manual เพื่อทดสอบ/populate โดยไม่ต้องรอ cron
        // อยู่ใน admin tier (ต้องมี session ที่ถูกต้อง); เป็น write จึงทำใต้ lock. ใช้กับวิชาเล็กเท่านั้น
        // (วิชาใหญ่ เช่น CVS 2516 ข้อ ควรใช้ nightly runQuestionRelationsBatch เพื่อเลี่ยงถือ lock นาน)
        if (action === 'runRelationsBatchManual') {
          var relSubj = data.subject;
          if (!relSubj) {
            return ContentService.createTextOutput(JSON.stringify({
              result: 'error', message: 'ต้องระบุ subject'
            })).setMimeType(ContentService.MimeType.JSON);
          }
          var relCount = generateQuestionRelationsForSubject(relSubj);
          writeAdminLog(user, userRole, "RELATIONS", "GENERATE", relSubj, "Generated question relations", "", String(relCount), metadata);
          return ContentService.createTextOutput(JSON.stringify({
            result: 'success', subject: relSubj, relationRows: relCount
          })).setMimeType(ContentService.MimeType.JSON);
        }

        // Feature 2: สั่งสกัด glossary ของวิชาเดียวแบบ manual (ทดสอบ/populate) โดยไม่ต้องรอ cron
        // ⚠️ ยิง LLM "ใต้ admin lock 25s" = admin wall ที่แผนยอมรับ — ใช้เฉพาะวิชาเล็ก/ทดสอบเท่านั้น
        // (nightly runGlossaryBatch วิ่งผ่าน time-driven trigger ไม่ผ่าน doPost จึง "ไม่ถือ lock";
        //  trigger เว้นไว้ไม่ติดตั้ง — tap/select path (askGlossaryTerm) เติม glossary เองแบบ lock-free)
        if (action === 'runGlossaryBatchManual') {
          var glSubj = data.subject;
          if (!glSubj) {
            return ContentService.createTextOutput(JSON.stringify({
              result: 'error', message: 'ต้องระบุ subject'
            })).setMimeType(ContentService.MimeType.JSON);
          }
          var glWritten = generateGlossaryForSubject(glSubj);
          writeAdminLog(user, userRole, "GLOSSARY", "GENERATE", glSubj, "Generated glossary terms", "", String(glWritten), metadata);
          return ContentService.createTextOutput(JSON.stringify({
            result: 'success', subject: glSubj, termsWritten: glWritten
          })).setMimeType(ContentService.MimeType.JSON);
        }

        // Feature 3: สั่งสร้างชีทสรุป high-yield ของ "ทุกหมวดในวิชาเดียว" แบบ manual (ทดสอบ/populate) โดยไม่ต้องรอ cron
        // ⚠️ ยิง LLM (หลาย call ต่อหมวด) "ใต้ admin lock 25s" = admin wall ที่แผนยอมรับ — ใช้เฉพาะวิชาเล็ก/ทดสอบเท่านั้น
        // (วิชาใหญ่ควรใช้ nightly runHighYieldBatch ผ่าน time-driven trigger ซึ่งไม่ผ่าน doPost จึงไม่ถือ lock; trigger เว้นไว้ไม่ติดตั้ง)
        if (action === 'runHighYieldBatchManual') {
          var hySubj = data.subject;
          if (!hySubj) {
            return ContentService.createTextOutput(JSON.stringify({
              result: 'error', message: 'ต้องระบุ subject'
            })).setMimeType(ContentService.MimeType.JSON);
          }
          var hyGen = generateHighYieldForSubject(hySubj);
          writeAdminLog(user, userRole, "HIGHYIELD", "GENERATE", hySubj, "Generated high-yield sheets", "", String(hyGen), metadata);
          return ContentService.createTextOutput(JSON.stringify({
            result: 'success', subject: hySubj, categoriesWritten: hyGen
          })).setMimeType(ContentService.MimeType.JSON);
        }

        // Feature 6: สั่งสร้าง keyword index ของ "ทุกหมวดในวิชาเดียว" แบบ manual (§6.2 idle-day/admin เท่านั้น) โดยไม่ต้องรอ cron
        // token-free 100% (ไม่ยิง LLM) → เร็ว ไม่ต้อง checkpoint. แต่วนทุกหมวด × ทุก term × ทุกข้อ "ใต้ admin lock 25s"
        // ⚠️ วิชาใหญ่ (เช่น CVS 2516 ข้อ หลายหมวด) อาจถือ lock นาน — ใช้เฉพาะวิชาเล็ก/ทดสอบ; ไม่มี public gen endpoint / ไม่ติดตั้ง trigger
        if (action === 'runKeywordIndexBatchManual') {
          var kwSubj = data.subject;
          if (!kwSubj) {
            return ContentService.createTextOutput(JSON.stringify({
              result: 'error', message: 'ต้องระบุ subject'
            })).setMimeType(ContentService.MimeType.JSON);
          }
          var kwGen = generateKeywordIndexForSubject(kwSubj);
          writeAdminLog(user, userRole, "KEYWORDINDEX", "GENERATE", kwSubj, "Generated keyword index", "", String(kwGen.rows), metadata);
          return ContentService.createTextOutput(JSON.stringify({
            result: 'success', subject: kwSubj, categoriesProcessed: kwGen.categories, rowsWritten: kwGen.rows
          })).setMimeType(ContentService.MimeType.JSON);
        }

        if (action === 'addAnnouncement') {
          sheet = doc.getSheetByName("Announcements");
          sheet.appendRow([
            data.data.Id,
            data.data.Text,
            data.data.Type,
            data.data.Active,
            data.data.Order
          ]);
          updateVersion();
          writeAdminLog(user, userRole, "ANNOUNCEMENT", "ADD", data.data.Id, "Added Announcement", "", data.data.Text, metadata);
          return ContentService.createTextOutput(JSON.stringify({ 'result': 'success' })).setMimeType(ContentService.MimeType.JSON);
        }

        if (action === 'editAnnouncement') {
          sheet = doc.getSheetByName("Announcements");
          var rows = sheet.getDataRange().getValues();
          for (var i = 1; i < rows.length; i++) {
            if (rows[i][0] == data.data.Id) {
              var oldText = rows[i][1];
              sheet.getRange(i + 1, 2, 1, 4).setValues([[
                data.data.Text,
                data.data.Type,
                data.data.Active,
                data.data.Order
              ]]);
              updateVersion();
              writeAdminLog(user, userRole, "ANNOUNCEMENT", "EDIT", data.data.Id, "Updated Announcement", oldText, data.data.Text, metadata);
              return ContentService.createTextOutput(JSON.stringify({ 'result': 'success' })).setMimeType(ContentService.MimeType.JSON);
            }
          }
        }

        if (action === 'deleteAnnouncement') {
          sheet = doc.getSheetByName("Announcements");
          var rows = sheet.getDataRange().getValues();
          for (var i = 1; i < rows.length; i++) {
            if (rows[i][0] == data.data.Id) {
              var oldText = rows[i][1];
              sheet.deleteRow(i + 1);
              updateVersion();
              writeAdminLog(user, userRole, "ANNOUNCEMENT", "DELETE", data.data.Id, "Deleted Announcement", oldText, "DELETED", metadata);
              return ContentService.createTextOutput(JSON.stringify({ 'result': 'success' })).setMimeType(ContentService.MimeType.JSON);
            }
          }
        }

        if (action === 'editQuestion') {
          sheet = doc.getSheetByName("Questions");
          var rows = sheet.getDataRange().getValues();
          var headers = rows[0];

          for (var i = 1; i < rows.length; i++) {
            if (rows[i][0] == data.data.id) {
              var oldRowData = {};
              for (var k = 0; k < headers.length; k++) {
                oldRowData[headers[k]] = rows[i][k];
              }

              var catToSave = Array.isArray(data.data.category) ? JSON.stringify(data.data.category) : data.data.category;
              sheet.getRange(i + 1, 2, 1, 6).setValues([
                [data.data.problem, data.data.img, data.data.choices, data.data.answer, data.data.explain, catToSave]
              ]);

              try {
                const catsForSplit = Array.isArray(data.data.category) ? data.data.category : JSON.parse(catToSave);
                autoCreateSplitCategories(data.data.id, catsForSplit);
              } catch (e) { console.log("Split error in editQuestion: " + e); }

              updateVersion();
              writeAdminLog(user, userRole, "QUESTION", "EDIT", data.data.id, "Question Updated", oldRowData, data.data, metadata);

              return ContentService.createTextOutput(JSON.stringify({
                'result': 'success'
              })).setMimeType(ContentService.MimeType.JSON);
            }
          }
        }

        // เพิ่ม category (เลคเชอร์) หลายข้อในครั้งเดียว — ใช้โดย AI batch categorizer ใน DATABASE admin panel
        // data.data.updates = [{id: questionId, categoryId: catId}] — append เท่านั้น ไม่ replace (semantics เดียวกับ vote-confirm)
        if (action === 'bulkAddQuestionCategories') {
          var updates = (data.data && data.data.updates) || [];
          if (!updates.length) {
            return ContentService.createTextOutput(JSON.stringify({ 'result': 'error', 'message': 'ไม่มีรายการ updates' })).setMimeType(ContentService.MimeType.JSON);
          }
          if (updates.length > 100) {
            return ContentService.createTextOutput(JSON.stringify({ 'result': 'error', 'message': 'เกิน 100 ข้อต่อรอบ — แบ่ง chunk จาก frontend' })).setMimeType(ContentService.MimeType.JSON);
          }

          sheet = doc.getSheetByName("Questions");
          var qData = sheet.getDataRange().getValues();
          var qIdMap = {};
          for (var i = 1; i < qData.length; i++) qIdMap[qData[i][0]] = i + 1;

          var applied = 0, skipped = 0;
          for (var u = 0; u < updates.length; u++) {
            var upd = updates[u];
            var rowIdx = qIdMap[upd.id];
            if (!rowIdx || !upd.categoryId) { skipped++; continue; }

            var rawCat = String(qData[rowIdx - 1][6] || '');
            var cats = [];
            try { cats = rawCat ? JSON.parse(rawCat.replace(/'/g, '"')) : []; }
            catch (e) { cats = rawCat ? [rawCat] : []; }

            if (cats.indexOf(upd.categoryId) !== -1) { skipped++; continue; }
            cats.push(upd.categoryId);
            sheet.getRange(rowIdx, 7).setValue(JSON.stringify(cats));

            try { autoCreateSplitCategories(upd.id, cats, true); } // skipSort=true — sort ทีเดียวตอนจบ
            catch (e) { console.log("Split error in bulkAddQuestionCategories: " + e); }
            applied++;
          }

          if (applied > 0) {
            try { sortCategorySheet(); } catch (e) { console.log("Sort error in bulkAddQuestionCategories: " + e); }
            updateVersion();
          }
          writeAdminLog(user, userRole, "QUESTION", "BULK_CATEGORIZE", updates.length + " items", "AI batch categorize", "", { applied: applied, skipped: skipped }, metadata);

          return ContentService.createTextOutput(JSON.stringify({
            'result': 'success', 'applied': applied, 'skipped': skipped
          })).setMimeType(ContentService.MimeType.JSON);
        }

        if (action === 'deleteQuestion') {
          sheet = doc.getSheetByName("Questions");
          var rows = sheet.getDataRange().getValues();
          var headers = rows[0];

          for (var i = 1; i < rows.length; i++) {
            if (rows[i][0] == data.data.id) {
              var oldRowData = {};
              for (var k = 0; k < headers.length; k++) {
                oldRowData[headers[k]] = rows[i][k];
              }

              sheet.deleteRow(i + 1);
              updateVersion();
              writeAdminLog(user, userRole, "QUESTION", "DELETE", data.data.id, "Question Deleted", oldRowData, "DELETED", metadata);

              return ContentService.createTextOutput(JSON.stringify({
                'result': 'success'
              })).setMimeType(ContentService.MimeType.JSON);
            }
          }
        }

        if (action === 'adminImport') {
          var sheetMap = {
            'struct': 'Structure',
            'category': 'Category',
            'ques': 'Questions'
          };

          var payload = data.data;
          var realSheetName = sheetMap[payload.sheetName] || payload.sheetName;
          var targetSheet = doc.getSheetByName(realSheetName);

          if (!targetSheet) {
            return ContentService.createTextOutput(JSON.stringify({
              'result': 'error',
              'message': 'ไม่พบแผ่นงาน: ' + realSheetName
            })).setMimeType(ContentService.MimeType.JSON);
          }

          var importData = payload.data;
          if (!importData || importData.length === 0) {
            return ContentService.createTextOutput(JSON.stringify({
              'result': 'error',
              'message': 'ข้อมูลว่างเปล่า'
            })).setMimeType(ContentService.MimeType.JSON);
          }

          try {
            var lastRow = targetSheet.getLastRow();

            if (realSheetName === 'Questions') {
              var appended = 0, updated = 0;
              var toAppend = [];

              // T2.4: หลีกเลี่ยง setValues ต่อแถวสำหรับข้อที่อัปเดต — อ่าน block เดียว แก้ใน memory แล้วเขียนกลับครั้งเดียว
              // เงื่อนไข: ทุกแถวนำเข้าต้องกว้างเท่ากัน (uniform) จึงเขียนเป็นบล็อกสี่เหลี่ยมได้อย่างปลอดภัย
              var width = importData[0].length;
              var uniform = importData.every(function (r) { return r.length === width; });

              if (uniform) {
                // อ่านบล็อกที่มีอยู่ครั้งเดียว (กว้าง = width) เพื่อคงคอลัมน์ส่วนเกิน (ถ้ามี) ไว้ไม่ถูกแตะ
                var block = lastRow > 1 ? targetSheet.getRange(2, 1, lastRow - 1, width).getValues() : [];
                var idToIdx = {};
                for (var bi = 0; bi < block.length; bi++) idToIdx[String(block[bi][0]).trim()] = bi;

                importData.forEach(function (row) {
                  var qId = String(row[0]).trim();
                  if (idToIdx.hasOwnProperty(qId)) {
                    block[idToIdx[qId]] = row; // อัปเดตใน memory
                    updated++;
                  } else {
                    idToIdx[qId] = block.length; // กันซ้ำภายใน payload เดียวกัน
                    block.push(row);
                    toAppend.push(row);
                  }
                });

                if (block.length > 0) {
                  // เขียนทั้งบล็อก (updates + appends) ในครั้งเดียว
                  targetSheet.getRange(2, 1, block.length, width).setValues(block);
                }
                appended = toAppend.length;
              } else {
                // Fallback (แถวกว้างไม่เท่ากัน): upsert ต่อแถวแบบเดิม เพื่อความปลอดภัยของข้อมูล
                var existing = lastRow > 1 ? targetSheet.getRange(2, 1, lastRow - 1, 1).getValues() : [];
                var existingIds = existing.map(function (r) { return String(r[0]).trim(); });
                importData.forEach(function (row) {
                  var qId = String(row[0]).trim();
                  var idx = existingIds.indexOf(qId);
                  if (idx >= 0) {
                    targetSheet.getRange(idx + 2, 1, 1, row.length).setValues([row]);
                    updated++;
                  } else {
                    toAppend.push(row);
                    existingIds.push(qId);
                  }
                });
                if (toAppend.length > 0) {
                  var newLastRow = targetSheet.getLastRow();
                  targetSheet.getRange(newLastRow + 1, 1, toAppend.length, toAppend[0].length).setValues(toAppend);
                }
                appended = toAppend.length;
              }

              updateVersion();
              writeAdminLog(user, userRole, "DATA", "IMPORT", realSheetName,
                "Upserted Questions: " + appended + " added, " + updated + " updated", "",
                "Added " + appended + ", Updated " + updated, metadata);

              return ContentService.createTextOutput(JSON.stringify({
                'result': 'success',
                'count': appended + updated,
                'added': appended,
                'updated': updated,
                'message': 'นำเข้าสำเร็จ: เพิ่ม ' + appended + ' แถว, อัปเดต ' + updated + ' แถว'
              })).setMimeType(ContentService.MimeType.JSON);

            } else {
              var existingKeys = new Set();
              if (lastRow > 0) {
                var fullData = targetSheet.getDataRange().getValues();
                for (var i = 0; i < fullData.length; i++) {
                  if (realSheetName === 'Structure') {
                    existingKeys.add(fullData[i][1] + "|" + fullData[i][3]);
                  } else {
                    existingKeys.add(String(fullData[i][0]));
                  }
                }
              }

              var finalData = importData.filter(function (row) {
                var key = realSheetName === 'Structure' ? row[1] + "|" + row[3] : String(row[0]);
                return !existingKeys.has(key);
              });

              if (finalData.length > 0) {
                targetSheet.getRange(lastRow + 1, 1, finalData.length, finalData[0].length).setValues(finalData);
                updateVersion();
                writeAdminLog(user, userRole, "DATA", "IMPORT", realSheetName, "Imported " + finalData.length + " new rows (Skipped " + (importData.length - finalData.length) + " duplicates)", "", "Added " + finalData.length + " rows", metadata);

                return ContentService.createTextOutput(JSON.stringify({
                  'result': 'success',
                  'count': finalData.length,
                  'skipped': importData.length - finalData.length,
                  'message': 'นำเข้าสำเร็จ ' + finalData.length + ' แถว (ข้ามข้อมูลซ้ำ ' + (importData.length - finalData.length) + ' แถว)'
                })).setMimeType(ContentService.MimeType.JSON);
              } else {
                return ContentService.createTextOutput(JSON.stringify({
                  'result': 'success',
                  'count': 0,
                  'skipped': importData.length,
                  'message': 'ไม่มีข้อมูลใหม่ให้นำเข้า (ข้อมูลทั้งหมดมีอยู่แล้วในระบบ)'
                })).setMimeType(ContentService.MimeType.JSON);
              }
            }
          } catch (err) {
            return ContentService.createTextOutput(JSON.stringify({
              'result': 'error',
              'message': 'GS Error: ' + err.toString()
            })).setMimeType(ContentService.MimeType.JSON);
          }
        }

        if (action === 'updateReportStatus') {
          sheet = doc.getSheetByName("Report");
          var rows = sheet.getDataRange().getValues();
          for (var i = 1; i < rows.length; i++) {
            var sTime = rows[i][8] instanceof Date ? rows[i][8].toISOString() : String(rows[i][8]);

            if (sTime === String(data.data.timestamp)) {
              var oldStatus = rows[i][9];

              sheet.getRange(i + 1, 10, 1, 3).setValues([
                [data.data.status, data.data.adminNote, data.data.done]
              ]);

              updateVersion();
              writeAdminLog(user, userRole, "REPORT", "UPDATE", "Report_Row_" + (i + 1), "Updated Report Status", oldStatus, data.data.status, metadata);

              return ContentService.createTextOutput(JSON.stringify({ 'result': 'success' })).setMimeType(ContentService.MimeType.JSON);
            }
          }
        }

        if (action === 'deleteCategory') {
          var sheet = doc.getSheetByName("Category");
          var rows = sheet.getDataRange().getValues();
          for (var i = 1; i < rows.length; i++) {
            if (rows[i][0] == data.data.CategoryID) {
              var catNameOld = rows[i][3];
              sheet.deleteRow(i + 1);
              updateVersion();
              writeAdminLog(user, userRole, "CATEGORY", "DELETE", data.data.CategoryID, "Category Deleted", catNameOld, "DELETED", metadata);
              return ContentService.createTextOutput(JSON.stringify({
                'result': 'success'
              })).setMimeType(ContentService.MimeType.JSON);
            }
          }
        }

        if (action === 'updateCategory') {
          var sheet = doc.getSheetByName("Category");
          var rows = sheet.getDataRange().getValues();
          for (var i = 1; i < rows.length; i++) {
            if (rows[i][0] == data.data.CategoryID) {
              var oldName = rows[i][3];
              sheet.getRange(i + 1, 4).setValue(data.data.CategoryName);
              updateVersion();
              writeAdminLog(user, userRole, "CATEGORY", "EDIT", data.data.CategoryID, "Renamed Category", oldName, data.data.CategoryName, metadata);
              return ContentService.createTextOutput(JSON.stringify({
                'result': 'success'
              })).setMimeType(ContentService.MimeType.JSON);
            }
          }
        }

        if (action === 'deleteGroup') {
          var sheet = doc.getSheetByName("Category");
          var rows = sheet.getDataRange().getValues();
          var deletedCount = 0;
          for (var i = rows.length - 1; i >= 1; i--) {
            if (rows[i][1] == data.data.SubjectRef && rows[i][2] == data.data.AccordionGroup) {
              sheet.deleteRow(i + 1);
              deletedCount++;
            }
          }

          var structSheet = doc.getSheetByName("Structure");
          var sRows = structSheet.getDataRange().getValues();
          for (var i = sRows.length - 1; i >= 1; i--) {
            if (sRows[i][1] == data.data.SubjectRef && sRows[i][3] == data.data.AccordionGroup) {
              structSheet.deleteRow(i + 1);
            }
          }

          updateVersion();
          writeAdminLog(user, userRole, "GROUP", "DELETE", data.data.SubjectRef + "_" + data.data.AccordionGroup, "Deleted Group & " + deletedCount + " categories", "", "DELETED", metadata);
          return ContentService.createTextOutput(JSON.stringify({
            'result': 'success'
          })).setMimeType(ContentService.MimeType.JSON);
        }

        if (action === 'updateAccordionGroup') {
          var oldGroup = data.data.OldAccordionGroup;
          var newGroup = data.data.NewAccordionGroup;
          var subjectId = data.data.SubjectRef;
          var updatedCount = 0;

          var catSheet = doc.getSheetByName("Category");
          var cRows = catSheet.getDataRange().getValues();
          for (var i = 1; i < cRows.length; i++) {
            if (cRows[i][1] == subjectId && cRows[i][2] == oldGroup) {
              catSheet.getRange(i + 1, 3).setValue(newGroup);
              updatedCount++;
            }
          }

          var structSheet = doc.getSheetByName("Structure");
          var sRows = structSheet.getDataRange().getValues();
          for (var i = 1; i < sRows.length; i++) {
            if (sRows[i][1] == subjectId && sRows[i][3] == oldGroup) {
              structSheet.getRange(i + 1, 4).setValue(newGroup);
            }
          }

          updateVersion();
          writeAdminLog(user, userRole, "GROUP", "EDIT", subjectId + "_" + oldGroup, "Renamed Group", oldGroup, newGroup, metadata);
          return ContentService.createTextOutput(JSON.stringify({
            'result': 'success',
            'message': 'Updated ' + updatedCount + ' categories'
          })).setMimeType(ContentService.MimeType.JSON);
        }

        if (action === 'addSubject') {
          sheet = doc.getSheetByName("Structure");
          sheet.appendRow([data.data.Year, data.data.SubjectID, data.data.SubjectName, "GENERAL"]);
          updateVersion();
          writeAdminLog(user, userRole, "SUBJECT", "ADD", data.data.SubjectID, "Added Subject", "", data.data.SubjectName, metadata);
          return ContentService.createTextOutput(JSON.stringify({ 'result': 'success' })).setMimeType(ContentService.MimeType.JSON);
        }

        if (action === 'updateSubject') {
          sheet = doc.getSheetByName("Structure");
          var rows = sheet.getDataRange().getValues();
          for (var i = 1; i < rows.length; i++) {
            if (rows[i][1] == data.data.SubjectID) {
              var oldName = rows[i][2];
              sheet.getRange(i + 1, 1).setValue(data.data.Year);
              sheet.getRange(i + 1, 3).setValue(data.data.SubjectName);
              updateVersion();
              writeAdminLog(user, userRole, "SUBJECT", "EDIT", data.data.SubjectID, "Updated Subject Info", oldName, data.data.SubjectName, metadata);
              return ContentService.createTextOutput(JSON.stringify({ 'result': 'success' })).setMimeType(ContentService.MimeType.JSON);
            }
          }
        }

        if (action === 'deleteSubject') {
          var subjectId = data.data.SubjectID;
          var structSheet = doc.getSheetByName("Structure");
          var sRows = structSheet.getDataRange().getValues();
          for (var i = sRows.length - 1; i >= 1; i--) {
            if (sRows[i][1] == subjectId) structSheet.deleteRow(i + 1);
          }
          var catSheet = doc.getSheetByName("Category");
          var cRows = catSheet.getDataRange().getValues();
          for (var i = cRows.length - 1; i >= 1; i--) {
            if (cRows[i][1] == subjectId) catSheet.deleteRow(i + 1);
          }
          updateVersion();
          writeAdminLog(user, userRole, "SUBJECT", "DELETE", subjectId, "Deleted Subject & Related Data", "", "DELETED", metadata);
          return ContentService.createTextOutput(JSON.stringify({ 'result': 'success' })).setMimeType(ContentService.MimeType.JSON);
        }

        if (action === 'addCategory') {
          sheet = doc.getSheetByName("Category");
          sheet.appendRow([data.data.CategoryID, data.data.SubjectRef, data.data.AccordionGroup, data.data.CategoryName]);

          var structSheet = doc.getSheetByName("Structure");
          var sRows = structSheet.getDataRange().getValues();
          var groupExists = false;
          for (var i = 1; i < sRows.length; i++) {
            if (sRows[i][1] == data.data.SubjectRef && sRows[i][3] == data.data.AccordionGroup) {
              groupExists = true;
              break;
            }
          }

          if (!groupExists) {
            var year = "";
            var subjectName = "";
            for (var i = 1; i < sRows.length; i++) {
              if (sRows[i][1] == data.data.SubjectRef) {
                year = sRows[i][0];
                subjectName = sRows[i][2];
                break;
              }
            }
            if (subjectName) {
              structSheet.appendRow([year, data.data.SubjectRef, subjectName, data.data.AccordionGroup]);
            }
          }

          updateVersion();
          sortCategorySheet();
          writeAdminLog(user, userRole, "CATEGORY", "ADD", data.data.CategoryID, "Added Category", "", data.data.CategoryName, metadata);
          return ContentService.createTextOutput(JSON.stringify({ 'result': 'success' })).setMimeType(ContentService.MimeType.JSON);
        }
      }
      return ContentService.createTextOutput(JSON.stringify({
        'result': 'error',
        'message': 'Action "' + action + '" not found or logic failed'
      })).setMimeType(ContentService.MimeType.JSON);
    } finally {
      adminLock.releaseLock();
    }

  } catch (e) {
    return ContentService.createTextOutput(JSON.stringify({
      'result': 'error',
      'message': e.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

