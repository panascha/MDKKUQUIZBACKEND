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
      // verifyAnySession: คืน user ทั้ง Admin และ Student (แค่ยืนยันตัวตน — สิทธิ์แอดมินตรวจราย action อยู่แล้ว)
      var userObj = verifyAnySession(data.sessionToken);
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

    // ----------------------------------------------------
    // PROGRESS SYNC (cross-device continue) — ต้องมี session (Admin หรือ Student ก็ได้)
    // getProgress = อ่านล้วน lock-free; saveProgress = auth/ตรวจนอก lock แล้วเขียนใต้ localized-15s (แบบ ingestKB)
    // ----------------------------------------------------
    if (action === 'getProgress') {
      var gpUser = verifyAnySession(data.sessionToken);
      if (!gpUser) {
        return ContentService.createTextOutput(JSON.stringify({
          result: 'error', message: 'session_expired'
        })).setMimeType(ContentService.MimeType.JSON);
      }
      var gpSubject = String(data.subject || '').trim();
      if (!gpSubject) {
        return ContentService.createTextOutput(JSON.stringify({
          result: 'error', message: 'missing subject'
        })).setMimeType(ContentService.MimeType.JSON);
      }
      var gpSheet = doc.getSheetByName("Progress");
      if (gpSheet && gpSheet.getLastRow() > 1 && gpSheet.getLastColumn() >= 4) {
        // อ่านเฉพาะ 3 คอลัมน์ key ก่อน (ห้าม getDataRange ทั้งชีต — คอลัมน์ blob ใหญ่ อ่านทุกแถวทุก request ไม่ไหว)
        var gpKeys = gpSheet.getRange(1, 1, gpSheet.getLastRow(), 3).getValues();
        for (var gpI = 1; gpI < gpKeys.length; gpI++) {
          if (gpKeys[gpI][0] === gpUser.email && gpKeys[gpI][1] === gpSubject) {
            var gpCells = gpSheet.getRange(gpI + 1, 4, 1, gpSheet.getLastColumn() - 3).getValues()[0];
            var gpBlob = gpCells.filter(function (c) { return c !== "" && c != null; }).join("");
            var gpState;
            try {
              gpState = JSON.parse(Utilities.ungzip(
                Utilities.newBlob(Utilities.base64Decode(gpBlob), 'application/x-gzip')
              ).getDataAsString());
            } catch (gpErr) {
              return ContentService.createTextOutput(JSON.stringify({
                result: 'error', message: 'corrupt progress blob'
              })).setMimeType(ContentService.MimeType.JSON);
            }
            return ContentService.createTextOutput(JSON.stringify({
              result: 'success', timestamp: Number(gpKeys[gpI][2]) || 0, state: gpState
            })).setMimeType(ContentService.MimeType.JSON);
          }
        }
      }
      return ContentService.createTextOutput(JSON.stringify({ result: 'empty' })).setMimeType(ContentService.MimeType.JSON);
    }

    if (action === 'saveProgress') {
      // rate-limit ก่อน auth (กัน flood ด้วย garbage token) — client debounce ~2 นาที → 60/ชม. เหลือเฟือ
      if (!checkActionRateLimit('rl_prog_', data.sessionToken || 'anon', 60)) {
        return ContentService.createTextOutput(JSON.stringify({
          result: 'error', message: 'ซิงค์บ่อยเกินไป กรุณาลองใหม่ภายหลัง'
        })).setMimeType(ContentService.MimeType.JSON);
      }
      var spUser = verifyAnySession(data.sessionToken);
      if (!spUser) {
        return ContentService.createTextOutput(JSON.stringify({
          result: 'error', message: 'session_expired'
        })).setMimeType(ContentService.MimeType.JSON);
      }
      var spSubject = String(data.subject || '').trim();
      var spState = data.state;
      var spTs = spState && Number(spState.timestamp);
      if (!spSubject || !spState || !spTs) {
        return ContentService.createTextOutput(JSON.stringify({
          result: 'error', message: 'missing subject/state/timestamp'
        })).setMimeType(ContentService.MimeType.JSON);
      }
      // gzip นอก lock (งาน CPU ไม่ต้องถือ lock)
      var spB64 = Utilities.base64Encode(
        Utilities.gzip(Utilities.newBlob(JSON.stringify(spState), 'application/octet-stream')).getBytes()
      );
      var spChunks = [];
      for (var spOff = 0; spOff < spB64.length; spOff += 45000) {
        spChunks.push(spB64.substring(spOff, spOff + 45000));
      }
      var spLock = LockService.getScriptLock();
      if (!spLock.tryLock(15000)) {
        return ContentService.createTextOutput(JSON.stringify({
          result: 'error', message: 'เซิร์ฟเวอร์ไม่ตอบสนองเนื่องจากโหลดสูง (Lock Timeout)'
        })).setMimeType(ContentService.MimeType.JSON);
      }
      try {
        var spSheet = getOrCreateProgressSheet(doc);
        // อ่านเฉพาะ 3 คอลัมน์ key (เหตุผลเดียวกับ getProgress — ห้ามอ่านคอลัมน์ blob ทั้งชีต)
        var spKeys = spSheet.getRange(1, 1, spSheet.getLastRow(), 3).getValues();
        var spRow = -1;
        for (var spI = 1; spI < spKeys.length; spI++) {
          if (spKeys[spI][0] === spUser.email && spKeys[spI][1] === spSubject) { spRow = spI + 1; break; }
        }
        if (spRow !== -1) {
          var spStored = Number(spKeys[spRow - 1][2]) || 0;
          // Overwrite guard: รับเฉพาะ state ที่ใหม่กว่า — เครื่องเก่าค้างหน้าจอจะทับของใหม่ไม่ได้
          if (spStored >= spTs) {
            return ContentService.createTextOutput(JSON.stringify({
              result: 'stale', cloudTimestamp: spStored
            })).setMimeType(ContentService.MimeType.JSON);
          }
          // เคลียร์ chunk เก่าทั้งแถว (เผื่อ blob ใหม่สั้นกว่า) แล้วเขียนทับ
          var spLastCol = spSheet.getLastColumn();
          if (spLastCol > 2) spSheet.getRange(spRow, 3, 1, spLastCol - 2).clearContent();
          spSheet.getRange(spRow, 3, 1, 1 + spChunks.length).setValues([[spTs].concat(spChunks)]);
        } else {
          spSheet.appendRow([spUser.email, spSubject, spTs].concat(spChunks));
        }
        return ContentService.createTextOutput(JSON.stringify({ result: 'success', timestamp: spTs })).setMimeType(ContentService.MimeType.JSON);
      } finally {
        spLock.releaseLock();
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

    // AI Search Overview — สรุปภาพรวมคำค้นหา (REAL js/search.js) — lock-free
    // login-gate ก่อน rate-limit: บังคับ session จริงเพื่อให้ bucket แยกรายคน
    // (ถ้า fallback เป็น 'anon' guest ทั้งเว็บจะกองใน bucket เดียว 15/ชม. — bug เดียวกับที่ discussion.js กันไว้)
    if (action === 'getSearchAIOverview') {
      var soUser = verifyAnySession(data.sessionToken);
      if (!soUser) {
        return ContentService.createTextOutput(JSON.stringify({
          result: 'error', message: 'session_expired'
        })).setMimeType(ContentService.MimeType.JSON);
      }
      if (!checkActionRateLimit('rl_ai_overview_', soUser.email, 15)) {
        return ContentService.createTextOutput(JSON.stringify({
          result: 'error', message: 'คุณขอสรุปภาพรวมบ่อยเกินไป (สูงสุด 15 ครั้ง/ชั่วโมง) กรุณาลองใหม่ภายหลัง'
        })).setMimeType(ContentService.MimeType.JSON);
      }
      var soKeyword = String(data.keyword || '').trim().slice(0, 100);
      var soSnippets = Array.isArray(data.snippets) ? data.snippets.slice(0, 5) : [];
      if (!soKeyword || soSnippets.length < 2) {
        return ContentService.createTextOutput(JSON.stringify({
          result: 'error', message: 'ข้อมูลไม่พอสำหรับสรุปภาพรวม'
        })).setMimeType(ContentService.MimeType.JSON);
      }
      var soKeyInfo = getAvailableAIKey("Gemini", "gemini-3.5-flash-lite");
      var soRes = callGeminiSearchOverview(soKeyword, data.stats || {}, soSnippets, soKeyInfo);
      if (!soRes.ok) {
        return ContentService.createTextOutput(JSON.stringify({
          result: 'error', message: soRes.error
        })).setMimeType(ContentService.MimeType.JSON);
      }
      return ContentService.createTextOutput(JSON.stringify({
        result: 'success', overview: soRes.data, model: soRes.model
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // บริจาค/ปลุก Gemini key ลง AI_Config pool (converter ใช้) — localized lock (sheet write)
    // Rate-limit by donor (session/clientId/username) ก่อน validate — กัน key-testing oracle + garbage spray
    if (action === 'seedGeminiKey') {
      var gdRlKey = data.sessionToken || data.clientId || data.username || 'anon';
      if (!checkActionRateLimit('rl_gemdon_', gdRlKey, 3)) {
        return ContentService.createTextOutput(JSON.stringify({
          result: 'error', message: 'คุณส่งคำขอบริจาคบ่อยเกินไป (สูงสุด 3 ครั้ง/ชั่วโมง) กรุณาลองใหม่ภายหลัง'
        })).setMimeType(ContentService.MimeType.JSON);
      }
      var gdLock = LockService.getScriptLock();
      if (!gdLock.tryLock(15000)) {
        return ContentService.createTextOutput(JSON.stringify({
          result: 'error', message: 'ระบบกำลังไม่ว่าง กรุณาลองใหม่อีกครั้ง'
        })).setMimeType(ContentService.MimeType.JSON);
      }
      try {
        return seedGeminiKey(data.apiKey, data.donorName);
      } finally {
        gdLock.releaseLock();
      }
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

    // agentPoolStatus — read-only per-key quota snapshot สำหรับ dashboard ของ router (owner-only, auth เดียวกับ agentQuery)
    if (action === 'agentPoolStatus') {
      var poolStatusAuthed = verifyAgentQueryOwnerSecret(data.ownerSecret) || !!verifySessionToken(data.sessionToken);
      if (!poolStatusAuthed) {
        return ContentService.createTextOutput(JSON.stringify({
          result: 'error', message: 'session_expired'
        })).setMimeType(ContentService.MimeType.JSON);
      }
      try {
        return ContentService.createTextOutput(JSON.stringify({
          result: 'success', keys: getAgentPoolStatus()
        })).setMimeType(ContentService.MimeType.JSON);
      } catch (poolStatusErr) {
        return ContentService.createTextOutput(JSON.stringify({
          result: 'error', message: poolStatusErr.message
        })).setMimeType(ContentService.MimeType.JSON);
      }
    }

    // ownerEnableGeminiModel — owner-triggered re-probe of one AI_Models row (Q6 gate), mirrors the
    // existing runXBatchManual admin-triggered-batch pattern (e.g. runGlossaryBatchManual below) but
    // owner-secret-gated like agentQuery/agentPoolStatus (no admin-session fallback — spends real
    // Gemini quota + mutates the shared sheet). Wraps the editor-only enableGeminiModel().
    if (action === 'ownerEnableGeminiModel') {
      var oemAuthed = verifyAgentQueryOwnerSecret(data.ownerSecret);
      if (!oemAuthed) {
        return ContentService.createTextOutput(JSON.stringify({
          result: 'error', message: 'session_expired'
        })).setMimeType(ContentService.MimeType.JSON);
      }
      return enableGeminiModel(data.model);
    }

    // setModelRpd — admin panel เขียน RPD_Limit/Priority ของโมเดลใน AI_Models (P2-Q1/Q5/Q7)
    // auth: mirror getFeedback (sessionToken admin หรือ username+adminPass). lock-free: single-cell write ความถี่ต่ำ
    // (สอดคล้อง Q5 — AI_Config/AI_Models write ไม่ใช้ LockService; off-by-one ยอมรับได้)
    if (action === 'setModelRpd') {
      var smrUser = null;
      if (data.sessionToken) smrUser = verifySessionToken(data.sessionToken);
      else if (data.username) smrUser = verifyAdmin(data.username, data.adminPass);
      if (!smrUser) {
        return ContentService.createTextOutput(JSON.stringify({
          result: 'error', message: 'session_expired'
        })).setMimeType(ContentService.MimeType.JSON);
      }
      return setModelRpd(data.model, data.rpd, data.priority);
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

        // imageUrls: URL รูปโจทย์แบบสาธารณะ (lh3.googleusercontent.com) — จำกัด 4 รูป กัน payload บวม
        // รับเฉพาะ https เท่านั้น: กัน data:/file:/http: ที่ทำให้ gateway ไปดึงของแปลกๆ แทนเรา
        var isImages = Array.isArray(data.imageUrls)
          ? data.imageUrls.filter(function (u) { return typeof u === 'string' && /^https:\/\//.test(u); }).slice(0, 4)
          : [];

        try {
          var aiResult = executeChatbotQuery(data.prompt, isModel, 1, null, isImages);
          return ContentService.createTextOutput(JSON.stringify({
            result: 'success', answer: aiResult.content, servedModel: aiResult.servedModel,
            switched: aiResult.switched, imagesSent: !!aiResult.imagesSent,
            finishReason: aiResult.finishReason || null
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
      // AI Expert/chatbot = ปริมาณสูง → ปักหมุด flash-lite (RPD 500) ไม่ปล่อยตาม Priority ทะเบียน
      // ถ้าโควต้าตัวนี้หมด getAvailableAIKey จะตกไป Priority ปกติเอง (ai-gemini.gs:306-313)
      var apiKeyInfo = getAvailableAIKey(provider, "gemini-3.5-flash-lite");

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

    // ----------------------------------------------------
    // convertPdfBatch — Student PDF Converter Phase 1 (Idea/active/student-pdf-converter-plan.md)
    // Gemini proxy แปลง PDF→คำถามผ่าน AI_Config Gemini pool — lock-free tier (ไม่มี sheet write นอกจาก quota column, แบบเดียวกับ askAIExpert)
    // AUTH GATE: ต้องมี session KKU (Admin/Student — verifyAnySession) หรือ username+adminPass เดิมของ DATABASE
    //   — กันคนนอกใช้เป็น open Gemini proxy เผาโควต้า pool (advisor must-fix)
    // Rate limit 40 POST/ชม. ต่อ user — ขึ้นจาก 20 เมื่อ 2026-08-09
    //   เดิมคิดบน "~4 batch/ไฟล์" แต่ฝั่ง client ซอยชุดตามจำนวนข้อแล้ว (~15 ข้อ/ชุด)
    //   ไฟล์ 90 ข้อ = 7 POST ดังนั้น 20/ชม. เหลือแค่ 2 ไฟล์/ชม. ซึ่งน้อยเกินใช้งานจริง
    // ----------------------------------------------------
    if (action === 'convertPdfBatch') {
      // rate-limit ก่อน auth (กัน flood ด้วย garbage token — mirror saveProgress)
      var pcRlKey = data.sessionToken || data.username || data.clientId || 'anon';
      if (!checkActionRateLimit('rl_pdfconv_', pcRlKey, 40)) {
        return ContentService.createTextOutput(JSON.stringify({
          result: 'error', message: 'แปลง PDF บ่อยเกินไป (จำกัดต่อชั่วโมง) กรุณาลองใหม่ภายหลัง'
        })).setMimeType(ContentService.MimeType.JSON);
      }
      var pcUser = null;
      if (data.sessionToken) pcUser = verifyAnySession(data.sessionToken);
      if (!pcUser && data.username) pcUser = verifyAdmin(data.username, data.adminPass);
      if (!pcUser) {
        return ContentService.createTextOutput(JSON.stringify({
          result: 'error', message: 'session_expired'
        })).setMimeType(ContentService.MimeType.JSON);
      }
      var pcPrompt = String(data.prompt || '').trim();
      var pcPdf = String(data.pdfB64 || '');
      var pcImages = Array.isArray(data.images) ? data.images : [];
      if (!pcPrompt || (!pcPdf && pcImages.length === 0)) {
        return ContentService.createTextOutput(JSON.stringify({
          result: 'error', message: 'ข้อมูลไม่ครบ (ต้องมี prompt และ PDF หรือรูปหน้ากระดาษ)'
        })).setMimeType(ContentService.MimeType.JSON);
      }
      if (pcImages.length > 20) {
        return ContentService.createTextOutput(JSON.stringify({
          result: 'error', message: 'ส่งได้สูงสุด 20 หน้าต่อชุด กรุณาลด batch size'
        })).setMimeType(ContentService.MimeType.JSON);
      }
      // Student PDF converter = งานหนัก/คุณภาพต้องมาก่อน → ปักหมุด flash tier
      // (ไม่งั้นได้ flash-lite ตาม Priority 1 ในทะเบียน); หมดโควต้าแล้วค่อยตกตาม Priority
      var pcKeyInfo = getAvailableAIKey("Gemini", "gemini-3.5-flash");
      if (!pcKeyInfo) {
        return ContentService.createTextOutput(JSON.stringify({
          result: 'error', message: 'ขณะนี้ไม่มี Gemini API Key ที่พร้อมใช้งาน (โควต้ารายวันเต็มทุก Key) กรุณาลองใหม่พรุ่งนี้'
        })).setMimeType(ContentService.MimeType.JSON);
      }
      try {
        var pcRes = callGeminiConverter(pcPrompt, pcKeyInfo, pcPdf, pcImages);
        return ContentService.createTextOutput(JSON.stringify({
          result: 'success',
          raw: pcRes.raw,
          finishReason: pcRes.finishReason,
          servedModel: pcRes.model,
          usage: pcRes.usage || null, // token counts — ใช้แยกว่า "ข้อหาย" เพราะโมเดลออกไม่ครบ หรือคำตอบถูกตัด
          quota: (pcKeyInfo.usage + 1) + "/" + pcKeyInfo.limit
        })).setMimeType(ContentService.MimeType.JSON);
      } catch (pcErr) {
        return ContentService.createTextOutput(JSON.stringify({
          result: 'error', message: pcErr.message
        })).setMimeType(ContentService.MimeType.JSON);
      }
    }

    // ----------------------------------------------------
    // verifyQuestionBatch — 2-model debate + arbiter consensus verification (admin-tier auth, lock-free LLM execution)
    // Auth: verifyAdmin OR verifySessionToken. NO lock during LLM calls (advisory rule).
    // Models hardcoded: DeepSeek-V4-Pro + Claude-Sonnet-4.5 via executeChatbotQuery, judge = Gemini-3.5-Flash via callGeminiAI.
    // ----------------------------------------------------
    if (action === 'verifyQuestionBatch') {
      var vqUser = null;
      if (data.sessionToken) vqUser = verifySessionToken(data.sessionToken);
      if (!vqUser && data.username) vqUser = verifyAdmin(data.username, data.adminPass);
      if (!vqUser) {
        return ContentService.createTextOutput(JSON.stringify({
          result: 'error', message: 'session_expired'
        })).setMimeType(ContentService.MimeType.JSON);
      }
      try {
        var vqRes = verifyQuestionBatch(data.questions, vqUser);
        return ContentService.createTextOutput(JSON.stringify(vqRes)).setMimeType(ContentService.MimeType.JSON);
      } catch (vqErr) {
        return ContentService.createTextOutput(JSON.stringify({
          result: 'error', message: vqErr.message || String(vqErr)
        })).setMimeType(ContentService.MimeType.JSON);
      }
    }

    // ----------------------------------------------------
    // logUserInteraction — ระบบ log ใหม่ (แทน batchLog/UserActivity เดิมที่ถูกถอดออก)
    // เขียนลงไฟล์ audit แยก (getAuditSheetId) → overflow ไม่แตะคลังข้อสอบ (SHEET_ID)
    // อยู่ lock-free tier: append-only ต่อแถว, return ก่อนขอ Lock ใด ๆ
    // การ์ด (ลำดับเดียวกับ convertPdfBatch/submitFeedback): rate limit → batch cap → payload cap → field caps
    // identity ดึงจาก sessionToken ฝั่ง server เท่านั้น (ไม่เชื่อ email ที่ client ส่งมา — กัน spoof)
    // fire-and-forget: error ใด ๆ ตอบ success(dropped) ไม่โยน error ให้นักเรียน
    // ----------------------------------------------------
    if (action === 'logUserInteraction') {
      // PRIVACY: เก็บสถิติไม่ระบุตัวตนเท่านั้น (intent tag + feature) — ไม่ derive/เขียน identity ใด ๆ
      // clientId ที่ส่งมาถูกใช้ "เฉพาะ" เป็น rate-limit key ชั่วคราวใน CacheService — ไม่เคยเขียนลง sheet
      // (1) rate limit — แยก 2 bucket:
      //   - ai_intent → 120/ชม. (client ยิง 1 คำขอ/คำถาม)
      //   - feature events → 30/ชม. (batched) กัน event ถี่ ๆ มาเบียดโควต้า
      var uiEventsRaw = Array.isArray(data.events) ? data.events : [];
      var uiHasIntent = false;
      for (var upi = 0; upi < uiEventsRaw.length; upi++) {
        if (uiEventsRaw[upi] && uiEventsRaw[upi].eventType === 'ai_intent') { uiHasIntent = true; break; }
      }
      var uiKey = data.clientId || 'anon';   // throttle key ชั่วคราวเท่านั้น — ไม่ persist
      var uiRlOk = uiHasIntent
        ? checkActionRateLimit('rl_uiprompt_', uiKey, 120)
        : checkActionRateLimit('rl_uilog_', uiKey, 30);
      if (!uiRlOk) {
        return ContentService.createTextOutput(JSON.stringify({ result: 'success', dropped: true })).setMimeType(ContentService.MimeType.JSON);
      }
      // (2) payload cap 512KB (contents = raw POST body, คำนวณไว้แล้วต้นฟังก์ชัน)
      if (contents && contents.length > 524288) {
        return ContentService.createTextOutput(JSON.stringify({ result: 'success', dropped: true })).setMimeType(ContentService.MimeType.JSON);
      }
      var events = uiEventsRaw; // parse แล้วด้านบน
      if (events.length === 0) {
        return ContentService.createTextOutput(JSON.stringify({ result: 'success' })).setMimeType(ContentService.MimeType.JSON);
      }
      // (3) batch cap 50 แถว/คำขอ
      if (events.length > 50) events = events.slice(0, 50);

      var uiAppId = String(data.appId || 'unknown').slice(0, 40);
      writeInteractionEvents_(uiAppId, events);
      return ContentService.createTextOutput(JSON.stringify({ result: 'success' })).setMimeType(ContentService.MimeType.JSON);
    }

    // ----------------------------------------------------
    // batchLog — DEPRECATED: ระบบ UserActivity เดิมถูกถอดออก (write-only, ไม่มีใครอ่าน + ไม่มี cap)
    // คง stub ที่ "รับแล้วทิ้ง" ไว้ เพื่อไม่ให้ PWA client เวอร์ชันเก่า (cache) ยิงมาแล้ว error
    // ----------------------------------------------------
    if (action === 'batchLog') {
      return ContentService.createTextOutput(JSON.stringify({ result: 'success', dropped: true })).setMimeType(ContentService.MimeType.JSON);
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
    // submitFeedback — รายงานบั๊ก/เสนอฟีเจอร์ของ "ตัวแอป" (คนละระบบกับ submitReport ที่รายงานข้อสอบ)
    // โครงเดียวกับ saveProgress/ingestKB: rate-limit + validate + Drive save ทำ "นอก lock"
    // แล้วเขียน 1 แถวใต้ localized-15s lock. anonymous ได้ (token ไม่ valid = ส่งแบบ anonymous ไม่ error)
    // แผน: Idea/active/user-feedback-reporting.md
    // ----------------------------------------------------
    if (action === 'submitFeedback') {
      // (1) rate limit 5/ชม. — key = sessionToken > clientId (localStorage UUID) > 'anon'
      if (!checkActionRateLimit('rl_appfb_', data.sessionToken || data.clientId || 'anon', 5)) {
        return ContentService.createTextOutput(JSON.stringify({
          result: 'error', message: 'ส่งฟีดแบ็กบ่อยเกินไป (สูงสุด 5 ครั้ง/ชั่วโมง) กรุณาลองใหม่ภายหลัง'
        })).setMimeType(ContentService.MimeType.JSON);
      }
      // (2) validate: payload รวม ≤1.5MB, type ใน whitelist, description ไม่ว่าง, รูป ≤2
      if (contents.length > 1572864) {
        return ContentService.createTextOutput(JSON.stringify({
          result: 'error', message: 'ขนาดข้อมูลใหญ่เกินไป (เกิน 1.5MB) กรุณาลดขนาด/จำนวนรูปภาพ'
        })).setMimeType(ContentService.MimeType.JSON);
      }
      var fbType = String(data.type || '').trim();
      if (['Bug', 'Feature', 'Other'].indexOf(fbType) < 0) {
        return ContentService.createTextOutput(JSON.stringify({
          result: 'error', message: 'ประเภทฟีดแบ็กไม่ถูกต้อง'
        })).setMimeType(ContentService.MimeType.JSON);
      }
      var fbDesc = String(data.description || '').trim();
      if (!fbDesc) {
        return ContentService.createTextOutput(JSON.stringify({
          result: 'error', message: 'กรุณากรอกรายละเอียด'
        })).setMimeType(ContentService.MimeType.JSON);
      }
      var fbImages = data.images || [];
      if (!Array.isArray(fbImages) || fbImages.length > 2) {
        return ContentService.createTextOutput(JSON.stringify({
          result: 'error', message: 'แนบรูปได้สูงสุด 2 รูปต่อรายงาน'
        })).setMimeType(ContentService.MimeType.JSON);
      }
      // (3) identity: token valid → email (Admin หรือ Student — verifyAnySession); ไม่ valid = anonymous ไม่ error
      var fbEmail = 'anonymous';
      if (data.sessionToken) {
        var fbUser = verifyAnySession(data.sessionToken);
        if (fbUser && fbUser.email) fbEmail = fbUser.email;
      }
      // (4) เซฟรูปลง Drive subfolder "Feedback" — นอก lock (Drive ops ไม่แตะชีต; T0.2 pattern)
      var fbUrls = [];
      for (var fbI = 0; fbI < fbImages.length; fbI++) {
        try {
          fbUrls.push(saveFeedbackImageToDrive(String(fbImages[fbI])));
        } catch (fbImgErr) {
          fbUrls.push('upload_failed');
        }
      }
      // (5) เขียน 1 แถวใต้ localized-15s lock
      var fbLock = LockService.getScriptLock();
      if (!fbLock.tryLock(15000)) {
        return ContentService.createTextOutput(JSON.stringify({
          result: 'error', message: 'เซิร์ฟเวอร์ไม่ตอบสนองเนื่องจากโหลดสูง (Lock Timeout)'
        })).setMimeType(ContentService.MimeType.JSON);
      }
      try {
        var fbSheet = setupFeedbackSheet(); // lazy-create ครั้งแรก
        fbSheet.appendRow([
          new Date(),
          fbType,
          fbDesc.slice(0, 5000),
          fbEmail,
          String(data.clientId || '').slice(0, 64),
          String(data.context || '').slice(0, 2000),
          fbUrls.join('///'),
          'New',
          ''
        ]);
        return ContentService.createTextOutput(JSON.stringify({ result: 'success' }))
          .setMimeType(ContentService.MimeType.JSON);
      } finally {
        fbLock.releaseLock();
      }
    }

    // ----------------------------------------------------
    // getSubjectPopularity — merge server-side subject selection counts into local
    // auth via verifyAnySession (Student/Admin), lock-free read
    // ----------------------------------------------------
    if (action === 'getSubjectPopularity') {
      var spUser = verifyAnySession(data.sessionToken);
      if (!spUser || !spUser.email) {
        return ContentService.createTextOutput(JSON.stringify({ result: 'error', message: 'login_required' }))
          .setMimeType(ContentService.MimeType.JSON);
      }
      var spSheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Subjects_Popularity');
      var spData = spSheet ? spSheet.getDataRange().getValues() : [];
      var spCounts = {};
      for (var si = 0; si < spData.length; si++) {
        if (spData[si][0] === spUser.email) {
          spCounts[String(spData[si][1])] = Math.max(spCounts[String(spData[si][1])] || 0, Number(spData[si][2]) || 0);
        }
      }
      return ContentService.createTextOutput(JSON.stringify({ result: 'success', counts: spCounts }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // ----------------------------------------------------
    // getFeedback — อ่าน Feedback ทั้งหมด (admin เท่านั้น — แถวมี PII จึง "ห้าม" ไปรวมใน getAllData ที่ไม่ auth)
    // pure read → lock-free. verifySessionToken = admin-only (token Student คืน null)
    // ----------------------------------------------------
    if (action === 'getFeedback') {
      // DATABASE login เป็น username+adminPass (ไม่มี sessionToken) — รองรับทั้งสองแบบเหมือน uploadImage
      var gfUser = null;
      if (data.sessionToken) {
        gfUser = verifySessionToken(data.sessionToken);
      } else if (data.username) {
        gfUser = verifyAdmin(data.username, data.adminPass);
      }
      if (!gfUser) {
        return ContentService.createTextOutput(JSON.stringify({
          result: 'error', message: 'session_expired'
        })).setMimeType(ContentService.MimeType.JSON);
      }
      return getFeedbackRows();
    }

    // ====================================================
    // REVIEWS + DONATIONS (2026-08-28) — reviews.gs / donations.gs
    // submit* = localized-15s (lock ภายในฟังก์ชัน แบบ saveProgress/submitFeedback)
    // update*/get*Admin = auth ที่ router แล้วเรียกฟังก์ชัน (getReviewsForAdmin/updateReviewStatus จัดการ lock/read เอง)
    // ====================================================

    // student write รีวิว (localized-15s, upsert 1 คน/วิชา) — auth ทำใน submitReview (verifyAnySession)
    if (action === 'submitReview') {
      if (!checkActionRateLimit('rl_review_', data.sessionToken || data.clientId || 'anon', 5)) {
        return ContentService.createTextOutput(JSON.stringify({
          result: 'error', message: 'ส่งรีวิวบ่อยเกินไป (สูงสุด 5 ครั้ง/ชั่วโมง) กรุณาลองใหม่ภายหลัง'
        })).setMimeType(ContentService.MimeType.JSON);
      }
      return submitReview(doc, data);
    }

    // donation write + slip OCR (localized-15s; OCR/Drive นอก lock ใน submitDonation) — login ไม่บังคับ
    if (action === 'submitDonation') {
      if (contents.length > 10485760) { // 10MB payload guard (รูปสลิป base64)
        return ContentService.createTextOutput(JSON.stringify({
          result: 'error', message: 'ขนาดข้อมูลใหญ่เกินไป กรุณาลดขนาดรูปสลิป'
        })).setMimeType(ContentService.MimeType.JSON);
      }
      if (!checkActionRateLimit('rl_donate_', data.sessionToken || data.clientId || 'anon', 5)) {
        return ContentService.createTextOutput(JSON.stringify({
          result: 'error', message: 'ส่งสลิปบ่อยเกินไป (สูงสุด 5 ครั้ง/ชั่วโมง) กรุณาลองใหม่ภายหลัง'
        })).setMimeType(ContentService.MimeType.JSON);
      }
      return submitDonation(doc, data);
    }

    // admin read: รีวิวทุกสถานะ (dual-auth: sessionToken admin หรือ username+adminPass — เหมือน getFeedback)
    if (action === 'getReviewsAdmin') {
      var graUser = data.sessionToken ? verifySessionToken(data.sessionToken) : (data.username ? verifyAdmin(data.username, data.adminPass) : null);
      if (!graUser) return ContentService.createTextOutput(JSON.stringify({ result: 'error', message: 'session_expired' })).setMimeType(ContentService.MimeType.JSON);
      return getReviewsForAdmin();
    }

    // admin read: donations ทุกแถว รวม SlipDriveUrl (owner-only PII → ห้ามไปรวมใน getAllData ที่ไม่ auth)
    if (action === 'getDonations') {
      var gdnUser = data.sessionToken ? verifySessionToken(data.sessionToken) : (data.username ? verifyAdmin(data.username, data.adminPass) : null);
      if (!gdnUser) return ContentService.createTextOutput(JSON.stringify({ result: 'error', message: 'session_expired' })).setMimeType(ContentService.MimeType.JSON);
      return getDonationsForAdmin();
    }

    // admin moderate รีวิว (25s lock ใน updateReviewStatus) — verifySessionToken = admin-only (Student token → null)
    if (action === 'updateReviewStatus') {
      var ursUser = data.sessionToken ? verifySessionToken(data.sessionToken) : (data.username ? verifyAdmin(data.username, data.adminPass) : null);
      if (!ursUser) return ContentService.createTextOutput(JSON.stringify({ result: 'error', message: 'session_expired' })).setMimeType(ContentService.MimeType.JSON);
      return updateReviewStatus(doc, data);
    }

    // admin override สถานะบริจาค (25s lock ใน updateDonationStatus)
    if (action === 'updateDonationStatus') {
      var udsUser = data.sessionToken ? verifySessionToken(data.sessionToken) : (data.username ? verifyAdmin(data.username, data.adminPass) : null);
      if (!udsUser) return ContentService.createTextOutput(JSON.stringify({ result: 'error', message: 'session_expired' })).setMimeType(ContentService.MimeType.JSON);
      return updateDonationStatus(doc, data);
    }

    // ----------------------------------------------------
    // getAdminSync — delta-sync แดชบอร์ดแอดมิน: NOT_MODIFIED หรือ {small slices + question delta}
    // pure read → lock-free. dual auth แบบ getFeedback (sessionToken admin หรือ username+adminPass)
    // ----------------------------------------------------
    if (action === 'getAdminSync') {
      var gasUser = null;
      if (data.sessionToken) {
        gasUser = verifySessionToken(data.sessionToken);
      } else if (data.username) {
        gasUser = verifyAdmin(data.username, data.adminPass);
      }
      if (!gasUser) {
        return ContentService.createTextOutput(JSON.stringify({
          result: 'error', message: 'session_expired'
        })).setMimeType(ContentService.MimeType.JSON);
      }
      return getAdminSyncData(data.clientVer, data.since);
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
    // §2.7 deleteGlossaryTerm — ลบศัพท์ที่ไม่ควรอยู่ในคลัง (เช่นคำทั่วไป "the", "woman") ออกให้ทุกคน
    // ต้องล็อกอิน KKU (Admin หรือ Student — verifyAnySession); localized-15s tier (ลบ 1 แถวใต้ lock)
    // rate-limit ก่อน auth (กัน flood ด้วย garbage token) — mirror saveProgress
    // ----------------------------------------------------
    if (action === 'deleteGlossaryTerm') {
      if (!checkActionRateLimit('rl_gldel_', data.sessionToken || 'anon', 30)) {
        return ContentService.createTextOutput(JSON.stringify({
          result: 'error', message: 'ลบศัพท์บ่อยเกินไป กรุณาลองใหม่ภายหลัง'
        })).setMimeType(ContentService.MimeType.JSON);
      }
      var gdUser = verifyAnySession(data.sessionToken);
      if (!gdUser) {
        return ContentService.createTextOutput(JSON.stringify({
          result: 'error', message: 'session_expired'
        })).setMimeType(ContentService.MimeType.JSON);
      }
      var gdKey = normalizeGlossaryTerm(data.term_en);
      if (!gdKey) {
        return ContentService.createTextOutput(JSON.stringify({
          result: 'error', message: 'ไม่พบคำที่ต้องการลบ'
        })).setMimeType(ContentService.MimeType.JSON);
      }
      var gdLock = LockService.getScriptLock();
      if (!gdLock.tryLock(15000)) {
        return ContentService.createTextOutput(JSON.stringify({
          result: 'error', message: 'เซิร์ฟเวอร์ไม่ตอบสนองเนื่องจากโหลดสูง (Lock Timeout)'
        })).setMimeType(ContentService.MimeType.JSON);
      }
      try {
        var gdRes = deleteGlossaryRowLocked(gdKey);
        if (gdRes.deleted) {
          // audit trail — นักศึกษาทุกคนลบได้ ต้องตามรอยได้ว่าใครลบคำไหน
          writeAdminLog(gdUser.email, gdUser.role, "GLOSSARY", "DELETE", gdRes.term_en, "Deleted glossary term", gdRes.term_th, "", "");
        }
        return ContentService.createTextOutput(JSON.stringify({
          result: 'success', deleted: gdRes.deleted
        })).setMimeType(ContentService.MimeType.JSON);
      } finally {
        gdLock.releaseLock();
      }
    }

    // ----------------------------------------------------
    // Feature 4: postComment — เขียน Discussion 1 แถว (login-only, localized-15s)
    // rate-limit rl_disc_ 10/hr ก่อน auth (mirror deleteGlossaryTerm); เพดาน 100 comment เช็คใต้ lock ใน helper
    // ----------------------------------------------------
    if (action === 'postComment') {
      if (!checkActionRateLimit('rl_disc_', data.sessionToken || 'anon', 10)) {
        return ContentService.createTextOutput(JSON.stringify({
          result: 'error', message: 'แสดงความคิดเห็นบ่อยเกินไป (สูงสุด 10 ครั้ง/ชั่วโมง) กรุณาลองใหม่ภายหลัง'
        })).setMimeType(ContentService.MimeType.JSON);
      }
      var pcUser2 = verifyAnySession(data.sessionToken);
      if (!pcUser2) {
        return ContentService.createTextOutput(JSON.stringify({
          result: 'error', message: 'ต้องเข้าสู่ระบบก่อนแสดงความคิดเห็น'
        })).setMimeType(ContentService.MimeType.JSON);
      }
      var pcQid = String(data.qid || '').trim();
      var pcText = String(data.text || '').trim();
      var pcNick = String(data.nickname || '').trim() || pcUser2.displayName || String(pcUser2.email).split('@')[0];
      if (pcNick.length > 40) pcNick = pcNick.slice(0, 40);
      if (!pcQid) {
        return ContentService.createTextOutput(JSON.stringify({ result: 'error', message: 'missing qid' })).setMimeType(ContentService.MimeType.JSON);
      }
      if (!pcText) {
        return ContentService.createTextOutput(JSON.stringify({ result: 'error', message: 'ข้อความว่างเปล่า' })).setMimeType(ContentService.MimeType.JSON);
      }
      if (pcText.length > DISCUSSION_MAX_CHARS) {
        return ContentService.createTextOutput(JSON.stringify({ result: 'error', message: 'ข้อความยาวเกิน ' + DISCUSSION_MAX_CHARS + ' ตัวอักษร' })).setMimeType(ContentService.MimeType.JSON);
      }
      var pcLock = LockService.getScriptLock();
      if (!pcLock.tryLock(15000)) {
        return ContentService.createTextOutput(JSON.stringify({
          result: 'error', message: 'เซิร์ฟเวอร์ไม่ตอบสนองเนื่องจากโหลดสูง (Lock Timeout)'
        })).setMimeType(ContentService.MimeType.JSON);
      }
      try {
        var pcRes = postDiscussionCommentLocked_(pcQid, pcUser2.email, pcNick, pcText);
        if (!pcRes.ok) {
          return ContentService.createTextOutput(JSON.stringify({ result: 'error', message: pcRes.message })).setMimeType(ContentService.MimeType.JSON);
        }
        return ContentService.createTextOutput(JSON.stringify({ result: 'success', comment: pcRes.comment })).setMimeType(ContentService.MimeType.JSON);
      } finally {
        pcLock.releaseLock();
      }
    }

    // ----------------------------------------------------
    // Feature 4: deleteComment — soft-delete Discussion 1 แถว (login-only, localized-15s)
    // สิทธิ์: เจ้าของ (email ตรง) หรือ admin (role !== 'Student'); ตรวจใน helper ใต้ lock. ไม่ rate-limit (decision #6)
    // ----------------------------------------------------
    if (action === 'deleteComment') {
      // dual-auth: Google sessionToken (REAL self-delete/admin) หรือ username+adminPass (DATABASE moderation)
      var dcViaPassword = false;
      var dcUser = verifyAnySession(data.sessionToken);
      if (!dcUser && data.username) {
        dcUser = verifyAdmin(data.username, data.adminPass);
        dcViaPassword = !!dcUser;
      }
      if (!dcUser) {
        return ContentService.createTextOutput(JSON.stringify({
          result: 'error', message: 'session_expired'
        })).setMimeType(ContentService.MimeType.JSON);
      }
      var dcQid = String(data.qid || '').trim();
      var dcTs = String(data.timestamp || '').trim();
      if (!dcQid || !dcTs) {
        return ContentService.createTextOutput(JSON.stringify({ result: 'error', message: 'ข้อมูลไม่ครบ' })).setMimeType(ContentService.MimeType.JSON);
      }
      // username+adminPass ผ่าน verifyAdmin แปลว่าเป็นแอดมินโดยนิยาม → admin เสมอ (ไม่พึ่ง role field ที่อาจว่าง)
      var dcIsAdmin = dcViaPassword || dcUser.role !== 'Student';
      var dcLock = LockService.getScriptLock();
      if (!dcLock.tryLock(15000)) {
        return ContentService.createTextOutput(JSON.stringify({
          result: 'error', message: 'เซิร์ฟเวอร์ไม่ตอบสนองเนื่องจากโหลดสูง (Lock Timeout)'
        })).setMimeType(ContentService.MimeType.JSON);
      }
      try {
        var dcRes = deleteDiscussionCommentLocked_(dcQid, dcTs, dcUser.email, dcIsAdmin);
        if (!dcRes.ok) {
          return ContentService.createTextOutput(JSON.stringify({ result: 'error', message: dcRes.message })).setMimeType(ContentService.MimeType.JSON);
        }
        return ContentService.createTextOutput(JSON.stringify({ result: 'success' })).setMimeType(ContentService.MimeType.JSON);
      } finally {
        dcLock.releaseLock();
      }
    }

    // ----------------------------------------------------
    // Feature 4: getDiscussionAdmin — อ่านทุกแถว Discussion (รวม deleted + email PII) สำหรับ moderation ฝั่ง DATABASE
    // pure read → lock-free. dual-auth แบบ getFeedback (sessionToken admin หรือ username+adminPass). ห้ามไปรวมใน getAllData ที่ไม่ auth
    // ----------------------------------------------------
    if (action === 'getDiscussionAdmin') {
      var gdaUser = null;
      if (data.sessionToken) {
        gdaUser = verifySessionToken(data.sessionToken);
      } else if (data.username) {
        gdaUser = verifyAdmin(data.username, data.adminPass);
      }
      if (!gdaUser) {
        return ContentService.createTextOutput(JSON.stringify({
          result: 'error', message: 'session_expired'
        })).setMimeType(ContentService.MimeType.JSON);
      }
      return ContentService.createTextOutput(JSON.stringify({
        result: 'success', discussion: readAllDiscussionForAdmin_()
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // ----------------------------------------------------
    // Feature 5 (Item 5): setCommentStatus — ปัก/ยกเลิกปักหมุด "เฉลยที่ดีที่สุด" (admin-only, localized-15s)
    // auth เหมือน getDiscussionAdmin/getFeedback: verifySessionToken (admin เท่านั้น — Student token คืน null)
    // หรือ username+adminPass ของ DATABASE. ***ห้ามใช้ verifyAnySession*** (นั่นรับ Student ด้วย)
    // ----------------------------------------------------
    if (action === 'setCommentStatus') {
      var scsUser = null;
      if (data.sessionToken) scsUser = verifySessionToken(data.sessionToken);
      else if (data.username) scsUser = verifyAdmin(data.username, data.adminPass);
      if (!scsUser) {
        return ContentService.createTextOutput(JSON.stringify({
          result: 'error', message: 'session_expired'
        })).setMimeType(ContentService.MimeType.JSON);
      }
      var scsQid = String(data.qid || '').trim();
      var scsTs = String(data.timestamp || '').trim();
      var scsStatus = String(data.newStatus || '').trim();
      if (!scsQid || !scsTs) {
        return ContentService.createTextOutput(JSON.stringify({ result: 'error', message: 'ข้อมูลไม่ครบ' })).setMimeType(ContentService.MimeType.JSON);
      }
      if (scsStatus !== 'pinned' && scsStatus !== 'visible') {
        return ContentService.createTextOutput(JSON.stringify({ result: 'error', message: 'สถานะไม่ถูกต้อง' })).setMimeType(ContentService.MimeType.JSON);
      }
      var scsLock = LockService.getScriptLock();
      if (!scsLock.tryLock(15000)) {
        return ContentService.createTextOutput(JSON.stringify({
          result: 'error', message: 'เซิร์ฟเวอร์ไม่ตอบสนองเนื่องจากโหลดสูง (Lock Timeout)'
        })).setMimeType(ContentService.MimeType.JSON);
      }
      try {
        var scsRes = setDiscussionCommentStatusLocked_(scsQid, scsTs, scsStatus);
        if (!scsRes.ok) {
          return ContentService.createTextOutput(JSON.stringify({ result: 'error', message: scsRes.message })).setMimeType(ContentService.MimeType.JSON);
        }
        return ContentService.createTextOutput(JSON.stringify({ result: 'success' })).setMimeType(ContentService.MimeType.JSON);
      } finally {
        scsLock.releaseLock();
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
    // IMAGE UPLOAD GROUP (Drive I/O — ไม่จับ script lock)
    // เดิมอยู่ใต้ adminLock 25s: อัปโหลดลง Drive ใช้เวลา 2-5 วิ/รูป ทำให้คนอัปโหลดพร้อมกัน 2-3 คน
    // ชน Admin Lock Timeout และบล็อก editQuestion ของแอดมินคนอื่นไปด้วย
    // Drive ไม่แตะชีตข้อสอบ จึงไม่ต้องใช้ sheet lock; writeAdminLog ใช้ appendRow ซึ่ง atomic อยู่แล้ว
    // (ไม่มี read-modify-write) — การจับ script lock ครอบจะทำให้กลับไป serialize กับ admin CRUD โดยไม่ได้ correctness เพิ่ม
    // ----------------------------------------------------
    if (action === 'uploadImage' || action === 'uploadImagesBatch') {
      var upUser = null;
      if (data.sessionToken) {
        upUser = verifySessionToken(data.sessionToken);
      } else if (data.googleIdToken) {
        var upPayload = verifyGoogleToken(data.googleIdToken);
        if (upPayload) upUser = findAdminByEmail(upPayload.email);
      } else {
        upUser = verifyAdmin(data.username, data.adminPass);
      }

      if (!upUser) {
        return ContentService.createTextOutput(JSON.stringify({
          'result': 'error',
          'message': 'token_expired'
        })).setMimeType(ContentService.MimeType.JSON);
      }

      if (action === 'uploadImage') {
        try {
          var fileUrl = uploadQuestionImageToDrive(data.data.base64, data.data.questionId, data.data.type, data.data.subject, data.data.year);
          writeAdminLog(upUser.username, upUser.role, "IMAGE", "UPLOAD", data.data.questionId, "Uploaded new " + data.data.type + " image", "", fileUrl, "");

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

      // T2.5: อัปโหลดหลายรูปในการเรียกครั้งเดียว (สูงสุด 10 รูป) — พารามิเตอร์ต่อรายการเหมือน uploadImage
      // แต่ละรายการอยู่ใน data.images[] = { base64, questionId, type, subject, year }
      // คืน urls[] เรียงตามลำดับ input; รายการที่ล้มเหลวจะเป็น { error: "..." } (ไม่ทำให้ทั้ง batch ล้ม)
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
      // ใช้ร่วมกันทั้ง batch — กัน getQuestionRoutingInfo/getFoldersByName รันซ้ำต่อรูป
      var batchRouteCache = {};
      var batchFolderCache = {};
      for (var bi = 0; bi < images.length; bi++) {
        var item = images[bi] || {};
        try {
          if (!item.base64) { urls.push({ error: 'missing base64' }); continue; }
          var batchUrl = uploadQuestionImageToDrive(item.base64, item.questionId, item.type, item.subject, item.year, batchRouteCache, batchFolderCache);
          urls.push(batchUrl);
          successCount++;
        } catch (err) {
          urls.push({ error: err.message });
        }
      }

      writeAdminLog(upUser.username, upUser.role, "IMAGE", "UPLOAD_BATCH",
        (images[0] && images[0].questionId) || "",
        "Batch uploaded " + successCount + "/" + images.length + " images", "", "", "");

      return ContentService.createTextOutput(JSON.stringify({
        'result': 'success',
        'urls': urls
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // ----------------------------------------------------
    // LOCALIZED LOCK GROUP (Locks briefly for writes, tryLock 15s)
    // ----------------------------------------------------
    var localizedActions = ['submitVote', 'submitReport', 'voteOnReport', 'deleteSession', 'saveStudentId', 'syncSubjectPopularity'];
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

        // ยืนยันตัวตนด้วยรหัสนักศึกษา — ผู้ใช้ที่ auto-enroll ผ่าน Google SSO กรอกรหัส นศ. หลังล็อกอิน
        // ต้องมี session ที่ใช้ได้ (Admin หรือ Student ก็ได้) แล้วเขียนลงคอลัมน์ StudentID (col 9, index 8) ของแถวตนเองในชีต Admins
        if (action === 'saveStudentId') {
          var sidUser = verifyAnySession(data.sessionToken);
          if (!sidUser || !sidUser.email) {
            return ContentService.createTextOutput(JSON.stringify({ 'result': 'error', 'message': 'session_expired' })).setMimeType(ContentService.MimeType.JSON);
          }
          var newSid = String(data.studentId || '').trim();
          if (!/^\d{6,12}$/.test(newSid)) {
            return ContentService.createTextOutput(JSON.stringify({ 'result': 'error', 'message': 'รหัสนักศึกษาไม่ถูกต้อง (ต้องเป็นตัวเลข 6-12 หลัก)' })).setMimeType(ContentService.MimeType.JSON);
          }
          var adminsSheet = doc.getSheetByName("Admins");
          if (!adminsSheet) {
            return ContentService.createTextOutput(JSON.stringify({ 'result': 'error', 'message': 'ไม่พบชีต Admins' })).setMimeType(ContentService.MimeType.JSON);
          }
          var adminsData = adminsSheet.getDataRange().getValues();
          var myRow = -1;
          var myEmail = String(sidUser.email).trim().toLowerCase();
          for (var si = 1; si < adminsData.length; si++) {
            var rowSid = String(adminsData[si][8]).trim();
            // กันรหัสซ้ำ: มีคนอื่น (คนละอีเมล) ใช้รหัสนี้แล้ว
            if (rowSid && rowSid === newSid && String(adminsData[si][5]).trim().toLowerCase() !== myEmail) {
              return ContentService.createTextOutput(JSON.stringify({ 'result': 'error', 'message': 'รหัสนักศึกษานี้ถูกใช้ยืนยันตัวตนโดยบัญชีอื่นแล้ว' })).setMimeType(ContentService.MimeType.JSON);
            }
            if (String(adminsData[si][5]).trim().toLowerCase() === myEmail) myRow = si + 1;
          }
          if (myRow === -1) {
            return ContentService.createTextOutput(JSON.stringify({ 'result': 'error', 'message': 'ไม่พบบัญชีของคุณในระบบ' })).setMimeType(ContentService.MimeType.JSON);
          }
          adminsSheet.getRange(myRow, 9).setValue(newSid); // col 9 = StudentID (index 8)
          updateVersion();
          writeAdminLog(sidUser.displayName || myEmail, sidUser.role || "", "AUTH", "VERIFY_SID", "Admins", "ยืนยันตัวตนด้วยรหัสนักศึกษา", "", "", "");
          return ContentService.createTextOutput(JSON.stringify({ 'result': 'success', 'studentId': newSid })).setMimeType(ContentService.MimeType.JSON);
        }

        // syncSubjectPopularity — อัปเดตจำนวนครั้งที่เลือกวิชา (upsert ต่อคู่ email+subjectId)
        // ยิงมาจาก beacon ทุกครั้งที่ผู้ใช้เลือกวิชา = ความถี่สูงและใช้ shared lock ร่วมกับ submitVote/submitReport
        // จึงต้องมี rate limit เหมือน localized action อื่น (rl_kb_/rl_appfb_) กันคนเดียวยิงรัวจนแย่ง lock
        if (action === 'syncSubjectPopularity') {
          if (!checkActionRateLimit('rl_subjpop_', data.sessionToken || 'anon', 60)) {
            // fire-and-forget ฝั่ง client → ตอบ success (dropped) ไม่ให้ขึ้น error ให้ผู้ใช้เห็น
            return ContentService.createTextOutput(JSON.stringify({ 'result': 'success', 'dropped': true })).setMimeType(ContentService.MimeType.JSON);
          }
          var popUser = verifyAnySession(data.sessionToken);
          if (!popUser || !popUser.email) {
            return ContentService.createTextOutput(JSON.stringify({ 'result': 'error', 'message': 'session_expired' })).setMimeType(ContentService.MimeType.JSON);
          }
          var popSubjectId = String(data.subjectId || '').trim();
          var popCount = Number(data.count) || 0;
          if (!popSubjectId || popCount < 1) {
            return ContentService.createTextOutput(JSON.stringify({ 'result': 'error', 'message': 'invalid_subject' })).setMimeType(ContentService.MimeType.JSON);
          }
          var popEmail = String(popUser.email);
          var popSheet = doc.getSheetByName('Subjects_Popularity');
          if (!popSheet) {
            popSheet = doc.insertSheet('Subjects_Popularity');
            popSheet.appendRow(['Email', 'SubjectId', 'Count', 'LastUsed']);
          }
          var popData = popSheet.getDataRange().getValues();
          var popRowIdx = -1;
          for (var pi = 1; pi < popData.length; pi++) {
            if (String(popData[pi][0]) === popEmail && String(popData[pi][1]) === popSubjectId) {
              popRowIdx = pi + 1;
              break;
            }
          }
          if (popRowIdx !== -1) {
            // max() ไม่ใช่ทับตรงๆ — client ส่งยอดสะสมของ "เครื่องนั้น" มา
            // ถ้าผู้ใช้ล้าง localStorage หรือเปิดเครื่องใหม่ ยอดที่ส่งมาจะต่ำกว่าของจริง การทับตรงๆ = ข้อมูลหาย
            var popExisting = Number(popSheet.getRange(popRowIdx, 3).getValue()) || 0;
            if (popCount > popExisting) {
              popSheet.getRange(popRowIdx, 3).setValue(popCount);
              popSheet.getRange(popRowIdx, 4).setValue(new Date());
            }
          } else {
            popSheet.appendRow([popEmail, popSubjectId, popCount, new Date()]);
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
                sbMirrorVoteDeleted_(data.questionId, category);
              } else {
                voteSheet.getRange(foundRowIndex, 4).setValue(newVote);
                voteSheet.getRange(foundRowIndex, 5).setValue(timestamp);
                // VoteCount = ค่าสุดท้าย ไม่ใช่ delta — RPC ไม่บวกซ้ำให้
                sbMirrorVoteRow_({
                  QuestionID: data.questionId, Question: voteData[foundRowIndex - 1][1],
                  SuggestedTopic: category, VoteCount: newVote,
                  Time: timestamp.toISOString(), Status: voteData[foundRowIndex - 1][5]
                });
              }
            } else if (delta > 0) {
              voteSheet.appendRow([data.questionId, data.questionText, category, 1, timestamp, "Pending"]);
              sbMirrorVoteRow_({
                QuestionID: data.questionId, Question: data.questionText,
                SuggestedTopic: category, VoteCount: 1,
                Time: timestamp.toISOString(), Status: "Pending"
              });
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

          // Time คือคีย์ธรรมชาติของ reports (§9.11 ข้อ 7) ⇒ ต้องเป็นค่าเดียวกันทั้งชีทและ mirror
          var reportTime = new Date().toISOString();

          sheet.appendRow([
            data.from || "User",
            data.category || "",
            data.questionId || "",
            data.question || "",
            qImg,
            data.allChoices || "",
            ansSug,
            data.report || "",
            reportTime,
            "Pending",
            "",
            "FALSE",
            data.suggestedExplain || "",
            1
          ]);

          updateVotesVersion();
          sbMirrorReportRow_({
            Time: reportTime, From: data.from || "User", Category: data.category || "",
            QuestionID: data.questionId || "", Question: data.question || "", Image: qImg,
            Choices: data.allChoices || "", SuggestedAnswer: ansSug,
            ReportDetail: data.report || "", Status: "Pending", AdminNote: "", Done: "FALSE",
            SuggestedExplain: data.suggestedExplain || "", VoteCount: 1
          });
          return ContentService.createTextOutput(JSON.stringify({ 'result': 'success' })).setMimeType(ContentService.MimeType.JSON);
        }

        if (action === 'voteOnReport') {
          // Auth + rate limit — mirror syncSubjectPopularity (same localized tier).
          // เดิมไม่มีเช็คเลย → คนไม่ล็อกอินยิงซ้ำ 5 ครั้งดัน report แตะ threshold ได้ (T-report-vote)
          if (!checkActionRateLimit('rl_reportvote_', data.sessionToken || 'anon', 30)) {
            return ContentService.createTextOutput(JSON.stringify({ result: 'error', message: 'rate_limited' })).setMimeType(ContentService.MimeType.JSON);
          }
          var rvUser = verifyAnySession(data.sessionToken);
          if (!rvUser || !rvUser.email) {
            return ContentService.createTextOutput(JSON.stringify({ result: 'error', message: 'session_expired' })).setMimeType(ContentService.MimeType.JSON);
          }
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
          // ส่งแค่คีย์ + คอลัมน์ที่เปลี่ยน — คีย์ที่ไม่ส่ง = คอลัมน์ที่ไม่ถูกแตะ (§9.11 ข้อ 3)
          sbMirrorReportRow_({ Time: targetTs, VoteCount: newVotes });

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
        sbFlush_(); // dual-write ไป Postgres — นอก lock เสมอ (D14) และกลืน error ทุกชนิด
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
          // Auto-enroll: บัญชี KKU ทุกคนที่ login ครั้งแรก → เพิ่มเข้าชีต Admins เป็น role Admin ทันที
          // Password ใส่ค่าสุ่ม (SSO-only) — ห้ามเว้นว่าง เพราะ verifyAdmin เทียบตรงตัว ค่าว่างจะ login ผ่านด้วยรหัสว่าง
          var adminsSheet = doc.getSheetByName("Admins");
          var newUsername = String(email).split("@")[0];
          var newDisplayName = tokenPayload.name || newUsername;
          adminsSheet.appendRow([
            newUsername,
            "SSO_ONLY_" + Utilities.getUuid(),
            newDisplayName,
            "https://api.dicebear.com/7.x/avataaars/svg?seed=" + newUsername,
            "Admin",
            email,
            "", "", "", "", ""
          ]);
          updateVersion();
          var newAdmin = findAdminByEmail(email);
          var newToken = createSession(email, newAdmin);
          writeAdminLog(newDisplayName, "Admin", "AUTH", "AUTO_ENROLL", "Session", "Auto-enrolled KKU account as Admin via Google SSO", "", "", "");
          return ContentService.createTextOutput(JSON.stringify({
            'result': 'success',
            'user': newAdmin,
            'sessionToken': newToken
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
        var userObj = verifyUser(data); // verifyUser รับ data object (sessionToken > googleIdToken > username+adminPass) — เดิมส่ง (username, adminPass) ผิด signature ทำให้ auth ไม่ผ่านทุกกรณี
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
          sbMarkSheet_('Announcements');
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
              sbMarkSheet_('Announcements');
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
              sbMarkSheet_('Announcements'); // แถวหายจากชีท ⇒ delete-absent ของ replace_* จัดการให้
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

              var splitRes = null;
              try {
                const catsForSplit = Array.isArray(data.data.category) ? data.data.category : JSON.parse(catToSave);
                splitRes = autoCreateSplitCategories(data.data.id, catsForSplit);
              } catch (e) { console.log("Split error in editQuestion: " + e); }

              updateVersion();
              writeAdminLog(user, userRole, "QUESTION", "EDIT", data.data.id, "Question Updated", oldRowData, data.data, metadata);
              // ส่ง data.data.category (array ดิบ) ไม่ใช่ catToSave ที่ stringify แล้ว
              sbMirrorQuestion_(data.data, splitRes);

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
          var appliedIds = []; // qid จริงสำหรับ delta-feed (getChangedSince split ด้วย comma)
          var sbMirrorRows = [];
          var sbSplitChangedSheets = false;
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

            var bulkSplit = null;
            try { bulkSplit = autoCreateSplitCategories(upd.id, cats, true); } // skipSort=true — sort ทีเดียวตอนจบ
            catch (e) { console.log("Split error in bulkAddQuestionCategories: " + e); }
            applied++;
            appliedIds.push(String(upd.id).trim());
            // §9.11 ข้อ 3: ส่งแค่ {questionId, category} — คีย์ที่ไม่ส่ง = คอลัมน์ที่ไม่ถูกแตะ
            // ใช้ finalCategories ถ้ามี เพราะ split เขียนทับคอลัมน์นี้ต่อจากเรา
            sbMirrorRows.push({
              questionId: String(upd.id).trim(),
              category: (bulkSplit && bulkSplit.finalCategories) || cats
            });
            if (bulkSplit && bulkSplit.sheetsChanged) sbSplitChangedSheets = true;
          }

          if (applied > 0) {
            try { sortCategorySheet(); } catch (e) { console.log("Sort error in bulkAddQuestionCategories: " + e); }
            updateVersion();
            sbMirrorQuestionRows_(sbMirrorRows);
            // มาร์คเฉพาะเมื่อ autoCreateSplitCategories เพิ่มแถวจริง — ไม่งั้นทุกรอบ AI categorize
            // จะลาก replace_categories_all (~1,420 แถว) ไปด้วยโดยไม่มีอะไรเปลี่ยนเลย
            if (sbSplitChangedSheets) {
              sbMarkSheet_('Category');
              sbMarkSheet_('Structure');   // ขั้นที่ 6 ของมัน append ชีท Structure ได้เช่นกัน
            }
          }
          // targetId = comma-joined qid จริง เพื่อให้ delta-sync เห็นข้อที่เปลี่ยน (เดิม "N items" ทำ delta หลุด)
          writeAdminLog(user, userRole, "QUESTION", "BULK_CATEGORIZE", appliedIds.join(","), "AI batch categorize (" + updates.length + " items)", "", { applied: applied, skipped: skipped }, metadata);

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
              // แถวในชีทหายไปแล้ว ⇒ sweep มองไม่เห็นจาก "สภาพปัจจุบัน" ได้อีก
              // ตัวนี้คือสัญญาณหลัก ส่วน sweep เล่นซ้ำจาก log แถว QUESTION/DELETE เป็นตาข่ายรอง
              sbMirrorQuestionDeleted_(data.data.id);

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
              sbMirrorQuestionSheetRows_(importData); // แถวดิบคอลัมน์ 0..6, sbFlush_ แบ่งก้อนให้เอง

              // Delta-feed: log แถว group QUESTION พร้อม qid จริง (comma-joined) ให้ getChangedSince เห็นข้อที่ import
              // แบ่งรอบละ 1000 qid กัน 50k char/cell limit ของ Sheets
              var importedQids = importData.map(function (row) { return String(row[0]).trim(); }).filter(Boolean);
              for (var qi = 0; qi < importedQids.length; qi += 1000) {
                writeAdminLog(user, userRole, "QUESTION", "IMPORT",
                  importedQids.slice(qi, qi + 1000).join(","),
                  "Imported questions batch", "", "", metadata);
              }

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
                sbMarkSheet_(realSheetName); // Structure / Category → replace ทั้ง slice

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
          var targetQid = data.data.questionId ? String(data.data.questionId).trim() : null;
          var updatedCount = 0;

          for (var i = 1; i < rows.length; i++) {
            var sTime = rows[i][8] instanceof Date ? rows[i][8].toISOString() : String(rows[i][8]);
            var rowQid = String(rows[i][2] || "").trim();

            var shouldUpdate = false;
            if (targetQid) {
              if (rowQid === targetQid) shouldUpdate = true;
            } else if (sTime === String(data.data.timestamp)) {
              shouldUpdate = true;
            }

            if (shouldUpdate) {
              var oldStatus = rows[i][9];

              sheet.getRange(i + 1, 10, 1, 3).setValues([
                [data.data.status, data.data.adminNote, data.data.done]
              ]);

              updatedCount++;
              writeAdminLog(user, userRole, "REPORT", "UPDATE", "Report_Row_" + (i + 1), "Updated Report Status (batch)", oldStatus, data.data.status, metadata);
            }
          }

          if (updatedCount > 0) {
            updateVersion();
            return ContentService.createTextOutput(JSON.stringify({ result: 'success', updated: updatedCount })).setMimeType(ContentService.MimeType.JSON);
          }
          return ContentService.createTextOutput(JSON.stringify({ result: 'success', updated: 0, message: 'No matching reports found' })).setMimeType(ContentService.MimeType.JSON);
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
              sbMarkSheet_('Category');
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
              sbMarkSheet_('Category');
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
          sbMarkSheet_('Category');
          sbMarkSheet_('Structure');
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
          sbMarkSheet_('Category');
          sbMarkSheet_('Structure');
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
          sbMarkSheet_('Structure');
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
              sbMarkSheet_('Structure');
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
          sbMarkSheet_('Structure');
          sbMarkSheet_('Category');
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
          sbMarkSheet_('Category');
          sbMarkSheet_('Structure'); // อาจ append กลุ่มใหม่ลง Structure ด้านบน
          return ContentService.createTextOutput(JSON.stringify({ 'result': 'success' })).setMimeType(ContentService.MimeType.JSON);
        }
      }
      return ContentService.createTextOutput(JSON.stringify({
        'result': 'error',
        'message': 'Action "' + action + '" not found or logic failed'
      })).setMimeType(ContentService.MimeType.JSON);
    } finally {
      // ★ ลำดับนี้สำคัญ: ถ่ายภาพชีท "ขณะยังถือ lock" แล้วค่อยปลด แล้วค่อยยิง HTTP
      //   อ่านชีทหลังปลด lock อาจได้ snapshot ที่ขาดแถว (doPost อีกตัว deleteRow เลื่อนแถว)
      //   ซึ่งไม่ว่างจึงผ่านการ์ด §Q ของ 004 แล้วแถวที่ขาดจะถูกลบจริงใน Postgres
      sbSnapshotDirtySheets_();
      adminLock.releaseLock();
      sbFlush_();
    }

  } catch (e) {
    return ContentService.createTextOutput(JSON.stringify({
      'result': 'error',
      'message': e.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

