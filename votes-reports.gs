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

    sheet.appendRow([
      new Date(), 
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
  } catch (e) { console.error("Admin Log Error: " + e.message); }
}

/**
 * ฟังก์ชันใหม่: บันทึกกิจกรรมผู้ใช้ (User Activity)
 */
function writeUserActivity(data) {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName("UserActivity") || ss.insertSheet("UserActivity");
    
    // Create header if empty
    if (sheet.getLastRow() == 0) {
      sheet.appendRow(["Timestamp", "SessionID", "Action", "TargetID", "Result", "TimeSpent", "Metadata"]);
      sheet.getRange(1, 1, 1, 7).setFontWeight("bold").setBackground("#e6f7ff");
    }

    sheet.appendRow([
      new Date(),
      data.session || "N/A",
      data.action || "",
      data.target || "",
      data.result || "",
      data.timeSpent || 0,
      data.metadata || ""
    ]);
  } catch (e) { console.error("User Log Error: " + e.message); }
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
        applyReportCorrection(qSheet, qRowIndex, suggestedAns, suggestedExplain, questionText, choicesArray);

        reportSheet.getRange(i+1, 10).setValue("AutoResolved");
        reportSheet.getRange(i+1, 11).setValue("Auto-applied by community vote (" + voteCount + "/" + REPORT_VOTE_THRESHOLD + ")");
        changed = true;
    }
    if (changed) updateVersion();
}

function applyReportCorrection(qSheet, qRowIndex, newAnswer, suggestedExplain, questionText, choicesArray) {
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
}

function buildExplainPrompt(questionText, choicesArray, correctAnswer) {
    var choicesText = choicesArray.map(function(c, i) {
        var display = (c.startsWith('http') || c.startsWith('<svg')) ? '[รูปภาพ]' : c;
        return String.fromCharCode(65 + i) + ". " + display;
    }).join("\n");
    return "คุณเป็นอาจารย์แพทย์ผู้เชี่ยวชาญ กรุณาเขียนคำอธิบายเฉลยข้อสอบแพทย์ต่อไปนี้เป็น paragraph เดียวต่อเนื่อง " +
        "(ห้ามใช้ bullet points หรือขึ้นบรรทัดใหม่) โดยใช้ภาษาไทยผสมคำศัพท์ทางการแพทย์ภาษาอังกฤษ ห้ามใช้ภาษาอังกฤษล้วน\n\n" +
        "โจทย์: " + questionText + "\n\n" +
        "ตัวเลือก:\n" + choicesText + "\n\n" +
        "เฉลยที่ถูกต้อง: " + correctAnswer + "\n\n" +
        "คำอธิบายต้องครอบคลุม: 1) Key concept/การวินิจฉัย 2) เหตุผลที่เฉลยถูก พร้อมชี้ clues จากโจทย์ " +
        "3) อธิบายว่าทำไมตัวเลือกที่ผิดแต่ละข้อถึงผิด 4) Clinical pearl ถ้ามี\n\n" +
        "เขียนเป็น paragraph เดียว ห้ามมี newline ในคำตอบ:";
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
        }
    }

}

