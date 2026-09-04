/* 
   =========================================
   ส่วนที่ 5: ระบบช่วยเหลือ (Utils) & Logging System
   =========================================
*/

// ฟังก์ชัน Log เก่า (เก็บไว้เพื่อความเข้ากันได้)
function writelog(user, action, targetId, details) {
    // แปลงให้ไปเรียก writeAdminLog แบบง่ายๆ
    writeAdminLog(user, "LEGACY", "SYSTEM", action, targetId, details, "", "", "");
}

/**
 * ฟังก์ชันใหม่: บันทึก Log ของ Admin โดยละเอียด
 */
function writeAdminLog(user, role, group, type, targetId, details, oldVal, newVal, meta) {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName("Logs") || ss.insertSheet("Logs");
    
    // Create header if empty
    if (sheet.getLastRow() == 0) {
      sheet.appendRow(["Timestamp", "User", "Role", "ActionGroup", "ActionType", "TargetID", "Details", "OldValue", "NewValue", "Metadata"]);
      sheet.getRange(1, 1, 1, 10).setFontWeight("bold").setBackground("#f3f3f3");
    }

    var oldStr = (typeof oldVal === 'object') ? JSON.stringify(oldVal) : String(oldVal || "");
    var newStr = (typeof newVal === 'object') ? JSON.stringify(newVal) : String(newVal || "");

    var logTime = new Date();
    sheet.appendRow([
      logTime,
      user,
      role,
      group,
      type,
      targetId,
      details,
      oldStr,
      newStr,
      meta || ""
    ]);

    // จุดเขียน log จุดเดียวของทั้งระบบ ⇒ hook ที่นี่ครอบทุก action ของแอดมิน รวม REPORT_AUTOFIX
    // ยกเว้น POSTGRES_MIRROR_FAIL เอง ไม่งั้น mirror ที่ล้มจะพยายาม mirror ความล้มเหลวของตัวเอง
    if (type !== 'POSTGRES_MIRROR_FAIL') {
      sbMirrorLogRow_({
        Timestamp: logTime.toISOString(), User: String(user || ''), Role: String(role || ''),
        ActionGroup: String(group || ''), ActionType: String(type || ''),
        TargetID: String(targetId || ''), Details: String(details || ''),
        OldValue: oldStr, NewValue: newStr, Metadata: String(meta || '')
      });
    }
  } catch (e) { console.error("Admin Log Error: " + e.message); }
}

/**
 * บันทึกสถิติการใช้งานแบบ "ไม่ระบุตัวตน" ลงไฟล์ audit ที่แยกจากคลังข้อสอบ (getAuditSheetId)
 * - append-only ต่อแถว (atomic ในตัว → เรียกได้ใน lock-free tier ไม่ต้องพึ่ง LockService)
 * - eventType === 'ai_intent' → แท็บ AI_Intents (tag + model), อื่น ๆ → แท็บ Interactions (feature)
 * - PRIVACY: ไม่รับ/ไม่เขียน email, studentId, clientId, userAgent, prompt ดิบ — เก็บแค่ tag/feature/app/เวลา
 */
function writeInteractionEvents_(appId, events) {
  try {
    var ss = SpreadsheetApp.openById(getAuditSheetId());
    var features = getAuditTab_(ss, 'Features', ["Timestamp", "AppId", "FeatureName"]);
    var intents = getAuditTab_(ss, 'AI_Intents', ["Timestamp", "AppId", "IntentTag", "Model"]);
    var now = new Date();

    for (var i = 0; i < events.length; i++) {
      var ev = events[i] || {};
      var type = String(ev.eventType || 'feature_use').slice(0, 40);

      if (type === 'ai_intent') {
        intents.appendRow([
          now, appId,
          String(ev.tag || 'other').slice(0, 40),
          String(ev.model || '').slice(0, 80)
        ]);
      } else {
        features.appendRow([
          now, appId,
          String(ev.feature || '').slice(0, 200)
        ]);
      }
    }
  } catch (e) { console.error("Audit Log Error: " + e.message); }
}


function processVotes() {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var voteSheet = ss.getSheetByName("Votes");
    var qSheet = ss.getSheetByName("Questions");
    if (!voteSheet || !qSheet) return;

    var voteValues = voteSheet.getDataRange().getValues();
    var qValues = qSheet.getDataRange().getValues();
    var qIdMap = {};
    for (var i = 1; i < qValues.length; i++) {
        qIdMap[qValues[i][0]] = i + 1;
    }

    var hasChanged = false; 

    for (var i = 1; i < voteValues.length; i++) {
        var voteCount = parseInt(voteValues[i][3]) || 0;
        var qId = voteValues[i][0];
        var categoryToAdd = voteValues[i][2];
        var status = String(voteValues[i][5]).trim();
        var currentRow = i + 1;

        // 1. ถ้าโหวตถึงเกณฑ์ Confirm (เช่น 2 คนขึ้นไป) -> Verified (เขียวเข้ม)
        if (voteCount >= VOTE_THRESHOLD_CONFIRM && status !== "Verified") {
            updateQuestionCategory(qSheet, qIdMap, qId, categoryToAdd);
            voteSheet.getRange(currentRow, 6).setValue("Verified");
            voteSheet.getRange(currentRow, 1, 1, 6).setBackground("#6aa84f"); 
            hasChanged = true;
        }
        // 2. ปรับใหม่: ถ้ามีคะแนนตั้งแต่ 1 ขึ้นไป และยังเป็น Pending -> Approved (สถานะรอเกณฑ์)
        //    หมายเหตุ: "ไม่" apply category ที่ 1 โหวตอีกต่อไป — ต้องถึง VOTE_THRESHOLD_CONFIRM (Verified) เท่านั้น
        //    ปิดบั๊ก "1 โหวต hijack หมวดถาวร" + หยุด updateVersion() ล้าง cache ทุกโหวต
        else if (voteCount >= 1 && (status === "Pending" || status === "")) {
            voteSheet.getRange(currentRow, 6).setValue("Approved");
            voteSheet.getRange(currentRow, 1, 1, 6).setBackground(null); // ล้างสี (สีขาว)
        }
    }

    if (hasChanged) {
        updateVersion();
        sortCategorySheet();
    }
}

function processReports(doc) {
    var reportSheet = doc.getSheetByName("Report");
    var qSheet = doc.getSheetByName("Questions");
    if (!reportSheet || !qSheet) return;
    var rv = reportSheet.getDataRange().getValues();
    var qv = qSheet.getDataRange().getValues();
    var qIdMap = {};
    for (var i = 1; i < qv.length; i++) qIdMap[qv[i][0]] = i + 1;

    var changed = false;
    for (var i = 1; i < rv.length; i++) {
        if (String(rv[i][9]).trim() !== "Pending") continue;
        var voteCount = parseInt(rv[i][13]) || 0;
        if (voteCount < REPORT_VOTE_THRESHOLD) continue;

        var qId = rv[i][2];
        var suggestedAns = String(rv[i][6] || "").trim();
        var suggestedExplain = String(rv[i][12] || "").trim();
        var qRowIndex = qIdMap[qId];
        if (!qRowIndex) continue;

        // Safety: suggestedChoice must exactly match an existing choice (no free-text)
        var choicesArray = String(qv[qRowIndex-1][3] || "").split("///").map(function(s){return s.trim();}).filter(Boolean);
        if (choicesArray.indexOf(suggestedAns) === -1) continue;

        var questionText = String(qv[qRowIndex-1][1] || "");
        applyReportCorrection(qSheet, qRowIndex, suggestedAns, suggestedExplain, questionText, choicesArray, qId);

        // Status(10) + AdminNote(11) + Done(12) — Done ต้องเป็น TRUE ด้วย ไม่งั้นแดชบอร์ดยังนับเป็น pending
        reportSheet.getRange(i+1, 10, 1, 3).setValues([["AutoResolved", "Auto-applied by community vote (" + voteCount + "/" + REPORT_VOTE_THRESHOLD + ")", "TRUE"]]);
        changed = true;
    }
    if (changed) updateVersion();
}

function applyReportCorrection(qSheet, qRowIndex, newAnswer, suggestedExplain, questionText, choicesArray, qId) {
    qSheet.getRange(qRowIndex, 5).setValue(newAnswer);
    var newExplain = suggestedExplain || "";
    try {
        var apiKeyInfo = getAvailableAIKey("Gemini");
        if (apiKeyInfo) {
            var prompt = buildExplainPrompt(questionText, choicesArray, newAnswer);
            var aiText = callGeminiAI(prompt, apiKeyInfo, null);
            newExplain = aiText.replace(/\r?\n/g, " ").trim();
        }
    } catch(e) {
        console.warn("Gemini explain failed: " + e.message);
    }
    qSheet.getRange(qRowIndex, 6).setValue(newExplain);
    // Delta-feed: แถว group QUESTION + qid จริง ให้ getChangedSince เห็นการ auto-apply correction
    if (!qId) qId = qSheet.getRange(qRowIndex, 1).getValue();
    writeAdminLog("SYSTEM", "SYSTEM", "QUESTION", "REPORT_AUTOFIX", String(qId), "Answer auto-corrected by community report vote: " + newAnswer, "", "", "");
}

function buildExplainPrompt(questionText, choicesArray, correctAnswer) {
    var choicesText = choicesArray.map(function(c) {
        var display = (c.startsWith('http') || c.startsWith('<svg')) ? '[รูปภาพ]' : c;
        return "- " + display;
    }).join("\n");
    return "คุณเป็นอาจารย์แพทย์ผู้เชี่ยวชาญ กรุณาเขียนคำอธิบายเฉลยข้อสอบแพทย์ต่อไปนี้เป็น paragraph เดียวต่อเนื่อง " +
        "(ห้ามใช้ bullet points หรือขึ้นบรรทัดใหม่) โดยใช้ภาษาไทยผสมคำศัพท์ทางการแพทย์ภาษาอังกฤษ ห้ามใช้ภาษาอังกฤษล้วน\n\n" +
        "โจทย์: " + questionText + "\n\n" +
        "ตัวเลือก:\n" + choicesText + "\n\n" +
        "เฉลยที่ถูกต้อง: " + correctAnswer + "\n\n" +
        "คำอธิบายต้องครอบคลุมตามลำดับดังนี้:\n" +
        "1) ชี้ diagnostic clues ในโจทย์และกลไกพยาธิสรีรวิทยา (Causal mechanism X → Y → Z) ที่นำไปสู่เฉลยที่ถูกต้อง\n" +
        "2) อธิบายแจกแจงตัวเลือกที่ผิด 'ครบทุกข้อที่เหลือ' (Process of elimination) โดยระบุชัดเจนว่าแต่ละข้อผิดเพราะอะไร และถ้าจะถูกต้องเป็นโรค/ภาวะใด " +
        "**ห้ามอ้างอิงตัวเลือกด้วยตัวอักษร A/B/C/D หรือหมายเลขข้อโดยเด็ดขาด** เพราะลำดับตัวเลือกถูกสลับใหม่ทุกครั้งที่แสดงผล ให้อ้างอิงด้วย 'ข้อความของตัวเลือก' ในเครื่องหมายคำพูดแทนเสมอ " +
        "(เช่น 'ส่วนตัวเลือก \"...\" ผิดเพราะ... ซึ่งจะพบในภาวะ...')\n" +
        "3) ปิดท้ายด้วย Clinical pearl หรือ High-yield point สั้นๆ\n\n" +
        "กรณีเฉลยถูกมากกว่า 1 ข้อ: หากในโจทย์หรือข้อมูลที่ให้มาระบุชัดเจนว่าข้อนี้มีตัวเลือกที่ถูกต้องได้มากกว่า 1 ข้อ " +
        "หรือเป็นข้อที่ให้คะแนนฟรี (อาจารย์เฉลยถูกหลายข้อ) ให้ขึ้นต้นย่อหน้าด้วยแท็ก **[⚠️ หมายเหตุข้อสอบจริง]:** " +
        "ชี้แจงสั้นๆ ว่าทำไมตัวเลือกเหล่านั้นถึงถูกได้ทั้งคู่ และทำไมระบบจึงตั้ง \"" + correctAnswer + "\" เป็นเฉลยหลัก " +
        "(เช่น เป็น First-line treatment / Gold standard ตาม Guideline ปัจจุบัน) แล้วจึงอธิบายกลไกต่อในย่อหน้าเดียวกัน " +
        "ห้ามอนุมานเงื่อนไขนี้เองจากความรู้สึกว่าโจทย์กำกวม ถ้าไม่มีหลักฐานชัดเจนให้เขียนคำอธิบายตามปกติ ห้ามใส่ /// ในข้อความหมายเหตุ\n\n" +
        "เขียนเป็น 1 paragraph ต่อเนื่อง ความยาวกระชับ ไม่เกิน ~300 คำ ห้ามมี newline หรือ bullet ในคำตอบ:";
}

function updateQuestionCategory(qSheet, qIdMap, qId, categoryToAdd) {
    var qRowIndex = qIdMap[qId];
    if (qRowIndex) {
        var catCell = qSheet.getRange(qRowIndex, 7);
        var currentCats = [];
        try {
            var val = catCell.getValue().toString();
            currentCats = (val && val !== "") ? JSON.parse(val.replace(/'/g, '"')) : [];
        } catch (e) {
            currentCats = [catCell.getValue().toString()];
        }

        if (currentCats.indexOf(categoryToAdd) === -1) {
            currentCats.push(categoryToAdd);
            catCell.setValue(JSON.stringify(currentCats));
            autoCreateSplitCategories(qId, currentCats);
            // Delta-feed: getChangedSince อ่านเฉพาะแถว group QUESTION + qid จริง — ไม่มีแถวนี้ delta-sync จะไม่เห็นการ auto-confirm
            writeAdminLog("SYSTEM", "SYSTEM", "QUESTION", "VOTE_CONFIRM", qId, "Category auto-confirmed by votes: " + categoryToAdd, "", "", "");
        }
    }

}

