function autoCreateSplitCategories(questionId, categories, skipSort) {
  if (!categories || categories.length < 2) return;

  // 1. กรอง "by AI" ออก และตรวจสอบว่ามาจาก Subject เดียวกันหรือไม่
  const validCats = categories.filter(c => !c.toLowerCase().includes("by ai"));
  if (validCats.length < 2) return;

  const firstCatId = validCats[0];
  const subjectId = firstCatId.split('_')[0]; // เช่น GI

  // ตรวจสอบว่าทุกอันขึ้นต้นด้วย Subject เดียวกัน
  const sameSubject = validCats.every(c => c.startsWith(subjectId));
  if (!sameSubject) return;

  // 2. ระบุ Source (ชุดข้อสอบเก่า) และ Lecture (หัวข้อเรียน)
  // สมมติ: อันแรกคือชุดข้อสอบ (GI_51MCQ1...), อันที่สองคือหัวข้อ (GI_ANA_...)
  const sourceCatId = validCats[0];
  const lectureCatId = validCats[1];

  // 3. Mapping หมวดหมู่ 6 กลุ่ม
  let splitSuffix = "";
  let groupKey = "";

  const upperLect = lectureCatId.toUpperCase();

  if (upperLect.includes("_ANA_")) { groupKey = "ANA"; splitSuffix = "ANATOMY (Extracted)"; }
  else if (upperLect.includes("_PHY_") || upperLect.includes("_PHYSIO_") || upperLect.includes("_BIOCHEM_")) { groupKey = "PHYSIO and BIOCHEM"; splitSuffix = "PHYSIO and BIOCHEM (Extracted)"; }
  else if (upperLect.includes("_PARASITO_") || upperLect.includes("_MICRO_")) { groupKey = "PARASITO and MICRO"; splitSuffix = "PARASITO and MICRO (Extracted)"; }
  else if (upperLect.includes("_PATHO_")) { groupKey = "PATHO"; splitSuffix = "PATHO (Extracted)"; }
  else if (upperLect.includes("_PHARM_") || upperLect.includes("_PHARMACO_")) { groupKey = "PHARM"; splitSuffix = "PHARM (Extracted)"; }
  else if (upperLect.includes("_IMAGE_") || upperLect.includes("_RADIO_") || upperLect.includes("_CLINICAL_")) { groupKey = "RADIO and CLINICAL"; splitSuffix = "RADIO and CLINICAL (Extracted)"; }

  if (!groupKey) return; // ถ้าไม่ตรงกับ 6 กลุ่มที่กำหนด ไม่ต้องทำต่อ

  // 4. สร้าง ID และชื่อใหม่
  // ตัวอย่าง: GI_51MCQ1_ANA_Extracted
  const newSplitCatId = `${sourceCatId}_${groupKey.replace(/\s+/g, '')}_Extracted`;
  const newSplitCatName = `${sourceCatId} (${splitSuffix})`;
  const newAccordionGroup = `${subjectId} (Extracted)`;

  const ss = SpreadsheetApp.openById(SHEET_ID);
  const catSheet = ss.getSheetByName("Category");
  const structSheet = ss.getSheetByName("Structure");

  // 5. ตรวจสอบและเพิ่มลงในแผ่นงาน Category
  const catValues = catSheet.getDataRange().getValues();
  let catExists = false;
  for (let i = 1; i < catValues.length; i++) {
    if (catValues[i][0] === newSplitCatId) { catExists = true; break; }
  }

  if (!catExists) {
    catSheet.appendRow([newSplitCatId, subjectId, newAccordionGroup, newSplitCatName]);
  }

  // 6. ตรวจสอบและเพิ่มลงในแผ่นงาน Structure
  const structValues = structSheet.getDataRange().getValues();
  let structExists = false;
  for (let i = 1; i < structValues.length; i++) {
    if (structValues[i][1] === subjectId && structValues[i][3] === newAccordionGroup) {
      structExists = true;
      break;
    }
  }

  if (!structExists) {
    // ดึง Year จากอันเดิมมาใส่ (ถ้าหาเจอ)
    let year = "0";
    for (let i = 1; i < structValues.length; i++) {
      if (structValues[i][1] === subjectId) { year = structValues[i][0]; break; }
    }
    structSheet.appendRow([year, subjectId, subjectId, newAccordionGroup]);
  }

  // 7. เพิ่ม NewSplitCatId เข้าไปในคำถามนั้น (ถ้ายังไม่มี)
  const qSheet = ss.getSheetByName("Questions");
  const qData = qSheet.getDataRange().getValues();
  for (let i = 1; i < qData.length; i++) {
    if (qData[i][0] === questionId) {
      let currentCats = [];
      try {
        currentCats = JSON.parse(qData[i][6].replace(/'/g, '"'));
      } catch (e) { currentCats = [qData[i][6]]; }

      if (!currentCats.includes(newSplitCatId)) {
        currentCats.push(newSplitCatId);
        qSheet.getRange(i + 1, 7).setValue(JSON.stringify(currentCats));
      }
      break;
    }
  }

  // บังคับข้ามการจัดเรียงหากทำงานอยู่ภายใต้คำสั่งประมวลผลเป็นกลุ่ม (Deferred Sorting)
  if (!skipSort) {
    sortCategorySheet();
  }
}

/**
 * ฟังก์ชันรันตรวจสอบคัดแยกวิชากลุ่มย่อย (Extracted) สำหรับคำถามทั้งหมดในแผ่นงาน Questions แบบ Manual
 */
function runManualSplitExtraction() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var ui = SpreadsheetApp.getUi();

  // ============================================================
  // STEP 1: โหลดข้อมูลทั้งหมดเข้า Memory ครั้งเดียว
  // ============================================================
  var qSheet   = ss.getSheetByName("Questions");
  var catSheet  = ss.getSheetByName("Category");
  var structSheet = ss.getSheetByName("Structure");

  if (!qSheet || !catSheet || !structSheet) {
    ui.alert("❌ ไม่พบ Sheet ที่จำเป็น");
    return;
  }

  var qData     = qSheet.getDataRange().getValues();
  var catData   = catSheet.getDataRange().getValues();
  var structData = structSheet.getDataRange().getValues();

  if (qData.length < 2) {
    ui.alert("ℹ️ ไม่พบข้อมูลคำถาม");
    return;
  }

  // ============================================================
  // STEP 2: สร้าง Lookup Maps จาก Memory (ไม่ต้องเปิด Sheet อีก)
  // ============================================================

  // Map: categoryId -> subjectId
  var catToSubject = {};
  // Set: subjectId|accordionGroup ที่มีอยู่ใน Structure แล้ว
  var existingStructKeys = {};
  // Set: categoryId ที่มีอยู่ใน Category แล้ว
  var existingCatIds = {};
  // Map: subjectId -> year (สำหรับ Structure row ใหม่)
  var subjectToYear = {};

  for (var i = 1; i < catData.length; i++) {
    var cid = String(catData[i][0]).trim();
    existingCatIds[cid] = true;
    catToSubject[cid] = String(catData[i][1]).trim();
  }
  for (var i = 1; i < structData.length; i++) {
    var sid = String(structData[i][1]).trim();
    var grp = String(structData[i][3]).trim();
    existingStructKeys[sid + "|" + grp] = true;
    if (!subjectToYear[sid]) subjectToYear[sid] = structData[i][0];
  }

  // ============================================================
  // STEP 3: คำนวณทุกอย่างใน Memory — ไม่แตะ Sheet เลยใน loop
  // ============================================================

  // ผลลัพธ์ที่จะเขียนทีเดียวตอนท้าย
  var newCatRows    = [];   // แถวใหม่สำหรับ Category Sheet
  var newStructRows = [];   // แถวใหม่สำหรับ Structure Sheet
  // Map: questionRowIndex -> categories array ที่อัปเดตแล้ว
  var qUpdates = {};

  var processCount  = 0;
  var skippedCount  = 0;

  var groupKeyMap = {
    "_ANA_":       { key: "ANA",             suffix: "ANATOMY (Extracted)" },
    "_PHY_":       { key: "PHYSIOandBIOCHEM", suffix: "PHYSIO and BIOCHEM (Extracted)" },
    "_PHYSIO_":    { key: "PHYSIOandBIOCHEM", suffix: "PHYSIO and BIOCHEM (Extracted)" },
    "_BIOCHEM_":   { key: "PHYSIOandBIOCHEM", suffix: "PHYSIO and BIOCHEM (Extracted)" },
    "_PARASITO_":  { key: "PARASITOandMICRO", suffix: "PARASITO and MICRO (Extracted)" },
    "_MICRO_":     { key: "PARASITOandMICRO", suffix: "PARASITO and MICRO (Extracted)" },
    "_PATHO_":     { key: "PATHO",            suffix: "PATHO (Extracted)" },
    "_PHARM_":     { key: "PHARM",            suffix: "PHARM (Extracted)" },
    "_PHARMACO_":  { key: "PHARM",            suffix: "PHARM (Extracted)" },
    "_IMAGE_":     { key: "RADIOandCLINICAL", suffix: "RADIO and CLINICAL (Extracted)" },
    "_RADIO_":     { key: "RADIOandCLINICAL", suffix: "RADIO and CLINICAL (Extracted)" },
    "_CLINICAL_":  { key: "RADIOandCLINICAL", suffix: "RADIO and CLINICAL (Extracted)" },
  };

  for (var i = 1; i < qData.length; i++) {
    var qId    = String(qData[i][0]).trim();
    var catRaw = String(qData[i][6]).trim();
    if (!qId || !catRaw) continue;

    // Parse categories
    var categories = [];
    try {
      categories = (catRaw.indexOf("[") > -1)
        ? JSON.parse(catRaw.replace(/'/g, '"'))
        : [catRaw];
    } catch (e) { continue; }

    // Guard: ต้องมี >= 2 categories
    if (categories.length < 2) { skippedCount++; continue; }

    // กรอง "by ai" และ "_Extracted" ออก
    var validCats = categories.filter(function(c) {
      var cu = c.toLowerCase();
      return !cu.includes("by ai") && !c.endsWith("_Extracted");
    });

    if (validCats.length < 2) { skippedCount++; continue; }

    var sourceCatId  = validCats[0];
    var lectureCatId = validCats[1];
    var subjectId    = sourceCatId.split('_')[0];

    // ตรวจว่าทุกอัน startsWith subjectId เดียวกัน
    var sameSubject = validCats.every(function(c) { return c.startsWith(subjectId); });
    if (!sameSubject) { skippedCount++; continue; }

    // หา groupKey จาก lectureCatId
    var upperLect = lectureCatId.toUpperCase();
    var matched = null;
    var keys = Object.keys(groupKeyMap);
    for (var k = 0; k < keys.length; k++) {
      if (upperLect.indexOf(keys[k]) > -1) { matched = groupKeyMap[keys[k]]; break; }
    }
    if (!matched) { skippedCount++; continue; }

    var newCatId       = sourceCatId + "_" + matched.key + "_Extracted";
    var newCatName     = sourceCatId + " (" + matched.suffix + ")";
    var newAccordion   = subjectId + " (Extracted)";

    // Guard: ข้ามถ้า extract แล้วและ question มี newCatId อยู่แล้ว
    if (existingCatIds[newCatId] && categories.indexOf(newCatId) > -1) {
      skippedCount++;
      continue;
    }

    // --- เพิ่ม Category ถ้ายังไม่มี (บันทึกใน memory) ---
    if (!existingCatIds[newCatId]) {
      newCatRows.push([newCatId, subjectId, newAccordion, newCatName]);
      existingCatIds[newCatId] = true;  // อัปเดต Map ใน memory ด้วย
      catToSubject[newCatId] = subjectId;
    }

    // --- เพิ่ม Structure ถ้ายังไม่มี (บันทึกใน memory) ---
    var structKey = subjectId + "|" + newAccordion;
    if (!existingStructKeys[structKey]) {
      var year = subjectToYear[subjectId] || "0";
      newStructRows.push([year, subjectId, subjectId, newAccordion]);
      existingStructKeys[structKey] = true;
    }

    // --- อัปเดต categories ของคำถาม (บันทึกใน memory) ---
    if (categories.indexOf(newCatId) === -1) {
      var updatedCats = categories.concat([newCatId]);
      qUpdates[i] = updatedCats;  // i = row index ใน qData array
    }

    processCount++;
  }

  // ============================================================
  // STEP 4: เขียนทุกอย่างลง Sheet ครั้งเดียว (Batch Write)
  // ============================================================

  if (newCatRows.length > 0) {
    var catLastRow = catSheet.getLastRow();
    catSheet.getRange(catLastRow + 1, 1, newCatRows.length, 4).setValues(newCatRows);
  }

  if (newStructRows.length > 0) {
    var structLastRow = structSheet.getLastRow();
    structSheet.getRange(structLastRow + 1, 1, newStructRows.length, 4).setValues(newStructRows);
  }

  // อัปเดต Questions Sheet: เขียนเฉพาะ row ที่เปลี่ยนแปลง
  var qRowIndices = Object.keys(qUpdates);
  for (var j = 0; j < qRowIndices.length; j++) {
    var rowIdx = parseInt(qRowIndices[j]);
    var sheetRow = rowIdx + 1; // +1 เพราะ getValues() เริ่มที่ index 0 = header
    qSheet.getRange(sheetRow, 7).setValue(JSON.stringify(qUpdates[rowIdx]));
  }

  // เรียก sort และ version update แค่ครั้งเดียวตอนท้าย
  if (processCount > 0 || newCatRows.length > 0) {
    updateVersion();
    sortCategorySheet();
  }

  ui.alert(
    "✅ เสร็จสิ้น\n" +
    "• ประมวลผล: " + processCount + " ข้อ\n" +
    "• Category ใหม่: " + newCatRows.length + " รายการ\n" +
    "• Structure ใหม่: " + newStructRows.length + " รายการ\n" +
    "• ข้าม: " + skippedCount + " ข้อ"
  );
}

function sortCategorySheet(ss) {
  if (!ss) ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName("Category");
  if (!sheet) return;

  // บังคับล้างลบแคชของ Category และตารางจัดกลุ่มความสัมพันธ์ในทันทีก่อนเรียงลำดับใหม่
  var cache = CacheService.getScriptCache();
  var v = getVersionCached();
  cache.remove("category_sheet_raw_" + v);
  cache.remove("cat_to_subj_map_" + v);

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  var range = sheet.getRange(2, 1, lastRow - 1, 4);
  var data = range.getValues();

  var extractYear = function (id) {
    var match = String(id).match(/\d+/);
    return match ? parseInt(match[0]) : 0;
  };

  var getSubSubjectPriority = function (group, id, name) {
    var text = (String(group) + " " + String(id) + " " + String(name)).toUpperCase();
    if (text.includes("_ANA_")) return 1;
    if (text.includes("_PHYSIO") || text.includes("BIOCHEM_")) return 2;
    if (text.includes("_MICRO") || text.includes("PARASITO_")) return 3;
    if (text.includes("_PATHO_")) return 4;
    if (text.includes("_PHARM_")) return 5;
    if (text.includes("_RADIO_") || text.includes("_CLINIC_")) return 6;
    return 7;
  };

  var getNumberSuffix = function (group, id, keyword) {
    var text = (String(group) + " " + String(id)).toUpperCase();
    var regex = new RegExp(keyword.toUpperCase() + "(\\d+)");
    var match = text.match(regex);
    if (match) return parseInt(match[1]);
    return 0;
  };

  var getGroupPriority = function (group, id) {
    var g = String(group).toUpperCase();
    var i = String(id).toUpperCase();

    if (g.includes("FMT")) return 10;
    if (g.includes("EXTRACTED") || i.includes("EXTRACTED")) return 30;
    if (g.includes("MCQ") || i.includes("MCQ")) return 20;
    if (g.includes("BY AI")) return 50;
    if (g.includes("LEC")) return 40;

    return 99;
  };

  data.sort(function (a, b) {
    var subA = String(a[1]);
    var subB = String(b[1]);
    if (subA !== subB) return subA.localeCompare(subB);

    var prioA = getGroupPriority(a[2], a[0]);
    var prioB = getGroupPriority(b[2], b[0]);

    if (prioA !== prioB) return prioA - prioB;

    if (prioA === 30 || prioA === 40 || prioA === 50) {
      var sRankA = getSubSubjectPriority(a[2], a[0], a[3]);
      var sRankB = getSubSubjectPriority(b[2], b[0], b[3]);
      if (sRankA !== sRankB) return sRankA - sRankB;
    }

    var yearA = extractYear(a[0]);
    var yearB = extractYear(b[0]);
    if (yearA !== yearB) return yearB - yearA;

    if (prioA === 10 || prioA === 20) {
      var keyword = (prioA === 10) ? "FMT" : "MCQ";
      var nA = getNumberSuffix(a[2], a[0], keyword);
      var nB = getNumberSuffix(b[2], b[0], keyword);
      if (nA !== nB) return nA - nB;
    }

    return 0;
  });

  // บังคับ Flush ข้อมูลลงชีตหลักให้เรียบร้อยก่อนเขียนทับ เพื่อความปลอดภัยของข้อมูล
  SpreadsheetApp.flush();
  range.setValues(data);
  updateVersion();
}
function verifyAllSingleCategoryVotes() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var voteSheet = ss.getSheetByName("Votes");
  var qSheet = ss.getSheetByName("Questions");
  if (!voteSheet || !qSheet) {
    SpreadsheetApp.getUi().alert("❌ ไม่พบแผ่นงาน Votes หรือ Questions");
    return;
  }

  var voteValues = voteSheet.getDataRange().getValues();
  var qValues = qSheet.getDataRange().getValues();
  
  // 1. สร้าง Map ของคำถามเพื่อความรวดเร็วในการตรวจสอบ
  // qMap[questionId] = { row: rowIndex, categories: [cat1, cat2] }
  var qMap = {};
  for (var i = 1; i < qValues.length; i++) {
    var qId = qValues[i][0];
    var catRaw = String(qValues[i][6] || "");
    var categories = [];
    
    try {
      if (catRaw !== "") {
        categories = (catRaw.indexOf("[") > -1) ? JSON.parse(catRaw.replace(/'/g, '"')) : [catRaw];
      }
    } catch (e) {
      categories = [catRaw];
    }
    qMap[qId] = { 
      row: i + 1, 
      categories: categories 
    };
  }

  var verifiedCount = 0;
  var qIdMapForUpdate = {}; // ใช้สำหรับฟังก์ชัน updateQuestionCategory เดิม
  for (var i = 1; i < qValues.length; i++) {
    qIdMapForUpdate[qValues[i][0]] = i + 1;
  }

  // 2. ไล่ดูรายการโหวต
  for (var j = 1; j < voteValues.length; j++) {
    var qId = voteValues[j][0];
    var categoryIdInVote = voteValues[j][2];
    var status = String(voteValues[j][5]).trim();
    
    // ข้ามถ้าเป็น Verified ไปแล้ว
    if (status === "Verified") continue;

    var qInfo = qMap[qId];
    if (qInfo) {
      // เงื่อนไข: มีแค่ 1 category
      if (qInfo.categories.length === 1) {
        var currentCat = qInfo.categories[0].toLowerCase();
        
        // เงื่อนไข: ต้องไม่ใช่ AI
        if (currentCat.indexOf("by ai") === -1) {
          
          // ทำการอัปเดตข้อมูล (เรียกใช้ฟังก์ชันที่มีอยู่แล้ว)
          updateQuestionCategory(qSheet, qIdMapForUpdate, qId, categoryIdInVote);
          
          // อัปเดตสถานะในหน้าโหวต
          voteSheet.getRange(j + 1, 6).setValue("Verified");
          voteSheet.getRange(j + 1, 1, 1, 6).setBackground("#6aa84f"); // สีเขียวเข้ม
          
          verifiedCount++;
        }
      }
    }
  }

  // 3. สรุปผล
  if (verifiedCount > 0) {
    updateVersion();
    sortCategorySheet();
    SpreadsheetApp.getUi().alert('✅ ดำเนินการ Verified เรียบร้อยแล้ว ' + verifiedCount + ' รายการ');
  } else {
    SpreadsheetApp.getUi().alert('ℹ️ ไม่พบคำถามที่ตรงตามเงื่อนไข (1 Category & No AI)');
  }
}

/**
 * สคริปต์ตรวจสอบรูปภาพในระบบเปรียบเทียบกับไฟล์ใน Google Drive
 * สร้างโดยอ้างอิง Root Folder ID: 1nzLH2ia2lL2TMxfrr6Kv-5fhsWwOWSCm
 */

function generateImageVerificationReport() {
  var TARGET_DRIVE_ID = '1nzLH2ia2lL2TMxfrr6Kv-5fhsWwOWSCm'; // ลิงก์ที่คุณให้มา
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var qSheet = ss.getSheetByName("Questions");
  
  // 1. สร้าง/เตรียมแผ่นงาน Report
  var reportSheet = ss.getSheetByName("Image_Migration_Report");
  if (reportSheet) {
    reportSheet.clear();
  } else {
    reportSheet = ss.insertSheet("Image_Migration_Report");
  }
  
  // เขียน Header ของ Report
  reportSheet.appendRow(["QuestionID", "Source Column", "Original URL", "Extracted ID", "Status", "Filename in Drive"]);
  reportSheet.getRange(1, 1, 1, 6).setFontWeight("bold").setBackground("#d9ead3");

  // 2. Brute Force สแกนหาไฟล์ทั้งหมดใน Drive (รวม Folder ย่อย)
  console.log("Starting Drive Scan... might take a while.");
  var driveMap = {}; // เก็บ {id: filename}
  var rootFolder = DriveApp.getFolderById(TARGET_DRIVE_ID);
  recursiveMapFolder(rootFolder, driveMap);
  console.log("Drive Scan Complete. Found " + Object.keys(driveMap).length + " files.");

  // 3. ดึงข้อมูลจาก Questions Sheet
  var data = qSheet.getDataRange().getValues();
  var reportRows = [];

  // เริ่มวนลูปข้อมูล (ข้าม Header แถวที่ 0)
  for (var i = 1; i < data.length; i++) {
    var qid = data[i][0];
    var imgCol = data[i][2];     // คอลัมน์ Image (โจทย์)
    var choicesCol = data[i][3]; // คอลัมน์ Choices (เผื่อมีรูปในตัวเลือก)

    // ตรวจสอบคอลัมน์รูปโจทย์
    processVerification(qid, "Image (Main)", imgCol, driveMap, reportRows);
    
    // ตรวจสอบคอลัมน์ตัวเลือก
    processVerification(qid, "Choices", choicesCol, driveMap, reportRows);
  }

  // 4. บันทึกผลลัพธ์ลงใน Sheet Report
  if (reportRows.length > 0) {
    reportSheet.getRange(2, 1, reportRows.length, 6).setValues(reportRows);
        console.log("ตรวจสอบเสร็จสิ้น! พบลิงก์รูปภาพทั้งหมด " + reportRows.length + " รายการ ใน Sheet 'Image_Migration_Report'");
return "Success: Found " + reportRows.length + " images.";
  } else {
    SpreadsheetApp.getUi().alert("ไม่พบลิงก์รูปภาพในฐานข้อมูล");
    return "No images found.";
  }
}

/**
 * ฟังก์ชันช่วยวนลูปหาไฟล์ในโฟลเดอร์ย่อยทั้งหมด (Recursive)
 */
function recursiveMapFolder(folder, driveMap) {
  var files = folder.getFiles();
  while (files.hasNext()) {
    var file = files.next();
    driveMap[file.getId()] = file.getName();
  }
  
  var subFolders = folder.getFolders();
  while (subFolders.hasNext()) {
    recursiveMapFolder(subFolders.next(), driveMap);
  }
}

/**
 * ฟังก์ชันสกัด ID และตรวจสอบสถานะ
 */
function processVerification(qid, colName, cellValue, driveMap, reportRows) {
  if (!cellValue || cellValue == "") return;
  
  // แยกส่วนด้วย /// กรณีมีหลายรูป
  var parts = String(cellValue).split("///");
  
  parts.forEach(function(part) {
    var trimmedPart = part.trim();
    if (trimmedPart.includes("drive.google.com") || trimmedPart.includes("id=")) {
      var fileId = extractId(trimmedPart);
      var status = "Not Found";
      var fileName = "-";
      
      if (driveMap[fileId]) {
        status = "Found";
        fileName = driveMap[fileId];
      }
      
      reportRows.push([qid, colName, trimmedPart, fileId, status, fileName]);
    }
  });
}

/**
 * ฟังก์ชัน Regex สกัด ID จาก URL หลากหลายรูปแบบ
 */
function extractId(url) {
  var match = url.match(/\/d\/(.*?)\//) || 
              url.match(/id=([^&]+)/) || 
              url.match(/\/d\/([^\/\?]+)/);
  return (match && match[1]) ? match[1] : url;
}

/**
 * เมนู: ถามรหัสวิชา แล้วสั่งจัดระเบียบรูปภาพของวิชานั้น
 * รันซ้ำได้เรื่อยๆ จนกว่าจะขึ้นว่าเสร็จ (ระบบจำแถวล่าสุดไว้ให้)
 */
function promptMigrateSubjectImages() {
  var ui = SpreadsheetApp.getUi();
  var res = ui.prompt('จัดระเบียบรูปภาพตามวิชา', 'ใส่รหัสวิชา (ขึ้นต้นของ QuestionID) เช่น SKIN, MS, GI', ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;

  var prefix = res.getResponseText().trim();
  if (!prefix) {
    ui.alert('ยังไม่ได้ใส่รหัสวิชา');
    return;
  }
  ui.alert(migrateSubjectImages(prefix));
}

/**
 * สคริปต์จัดระเบียบรูปภาพของวิชาที่ระบุ (ตั้งชื่อ Q_<qid>_Main_<n> / Q_<qid>_Choice_<A-E>)
 * หากไม่ได้เป็นเจ้าของไฟล์ จะย้ายไปที่โฟลเดอร์ MD > Unknown
 * รันเกิน 5 นาทีจะหยุดและจำแถวล่าสุดไว้ใน PropertiesService รันใหม่เพื่อทำต่อ
 */
function migrateSubjectImages(subjectPrefix, yearFolder) {
  var ROOT_FOLDER_ID = DRIVE_FOLDER_ID;
  var TIME_BUDGET_MS = 5 * 60 * 1000;
  var startTime = Date.now();
  var props = PropertiesService.getScriptProperties();
  var cursorKey = 'migrate_cursor_' + subjectPrefix;

  yearFolder = yearFolder || "Y2";

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var qSheet = ss.getSheetByName("Questions");
  var data = qSheet.getDataRange().getValues();
  var myEmail = Session.getEffectiveUser().getEmail(); // อีเมลผู้รันสคริปต์
  var startRow = Number(props.getProperty(cursorKey)) || 1;

  try {
    var rootFolder = DriveApp.getFolderById(ROOT_FOLDER_ID);
    console.log("--- Starting " + subjectPrefix + " Migration from row " + startRow + " (Active User: " + myEmail + ") ---");

    // โฟลเดอร์หลักและโฟลเดอร์ Unknown เตรียมครั้งเดียว
    var mdFolder = getOrCreateSubFolder(rootFolder, "MD");
    var unknownFolder = getOrCreateSubFolder(mdFolder, "Unknown");
    var folderCache = {}; // path -> Folder กันเรียก Drive ซ้ำ
    var processed = 0;

    for (var i = startRow; i < data.length; i++) {
      if (Date.now() - startTime > TIME_BUDGET_MS) {
        props.setProperty(cursorKey, String(i));
        var msg = "หยุดชั่วคราวที่แถว " + i + " จาก " + (data.length - 1) + " (จัดการไปแล้ว " + processed + " ข้อในรอบนี้) — รันเมนูเดิมซ้ำเพื่อทำต่อ";
        console.log(msg);
        return msg;
      }

      var qid = String(data[i][0]);
      var imgCell = String(data[i][2]);
      var choicesCell = String(data[i][3]);

      // 1. กรองเฉพาะวิชาที่ระบุ
      if (!qid || qid.split('_')[0] !== subjectPrefix) continue;

      var hasMainImg = (imgCell && imgCell != "" && !imgCell.toLowerCase().includes("require_img"));
      var hasChoiceImg = (choicesCell && choicesCell.includes("drive.google.com"));

      if (!hasMainImg && !hasChoiceImg) continue;

      // 2. วิเคราะห์ Path
      var parts = qid.split('_');
      if (parts.length < 2) continue;

      var yearType = parts[1].trim();
      var pathNames = ["MD", yearFolder, subjectPrefix];

      if (/^\d{2}/.test(yearType)) {
        pathNames.push(yearType.substring(0, 2)); // Year
        pathNames.push(yearType.substring(2) || "General"); // Type
      } else {
        pathNames.push(yearType);
      }

      // 3. เตรียมโฟลเดอร์ปลายทางตามโครงสร้างปกติ
      var pathKey = pathNames.join('/');
      var targetFolder = folderCache[pathKey];
      if (!targetFolder) {
        targetFolder = rootFolder;
        pathNames.forEach(function(name) {
          targetFolder = getOrCreateSubFolder(targetFolder, name);
        });
        folderCache[pathKey] = targetFolder;
      }

      // 4. จัดการรูปโจทย์
      if (hasMainImg) {
        var imgUrls = imgCell.split("///");
        imgUrls.forEach(function(url, idx) {
          var newName = "Q_" + qid + "_Main_" + (idx + 1);
          moveAndRenameFile(url, newName, targetFolder, unknownFolder);
        });
      }

      // 5. จัดการรูปในตัวเลือก
      if (hasChoiceImg) {
        var choiceParts = choicesCell.split("///");
        choiceParts.forEach(function(content, idx) {
          if (content.includes("drive.google.com") || content.includes("id=")) {
            var letter = String.fromCharCode(65 + idx);
            var newName = "Q_" + qid + "_Choice_" + letter;
            moveAndRenameFile(content, newName, targetFolder, unknownFolder);
          }
        });
      }
      processed++;
    }

    props.deleteProperty(cursorKey);
    console.log("--- Finished " + subjectPrefix + " Migration ---");
    return "เสร็จสิ้น! จัดการรูปของวิชา " + subjectPrefix + " ไปแล้ว " + processed + " ข้อในรอบนี้";
  } catch (e) {
    console.error("Critical Error: " + e.message);
    return "เกิดข้อผิดพลาด: " + e.message + " (ตำแหน่งล่าสุดถูกบันทึกไว้แล้ว รันซ้ำเพื่อทำต่อ)";
  }
}

/**
 * ฟังก์ชันย้ายไฟล์และเปลี่ยนชื่อ
 * พยายามย้ายเข้าโฟลเดอร์ปลายทางจริงก่อนเสมอ (รูปส่วนใหญ่แชร์ "anyone: writer" ถึงจะคนละเจ้าของก็ย้ายได้)
 * ถ้า Drive ปฏิเสธจริงๆ ค่อยตกไปที่ MD > Unknown
 */
function moveAndRenameFile(url, newName, targetFolder, unknownFolder) {
  try {
    var fileId = extractIdFromUrl(url);
    if (!fileId) return;

    var file = DriveApp.getFileById(fileId);
    var currentParents = file.getParents();
    var currentParentId = currentParents.hasNext() ? currentParents.next().getId() : "";

    // ถ้าชื่อตรงและอยู่ปลายทางแล้ว ให้ข้าม (ไฟล์ที่ตกค้างใน Unknown จะไม่เข้าเงื่อนไขนี้ จึงถูกย้ายต่อ)
    if (file.getName() === newName && currentParentId === targetFolder.getId()) {
      return;
    }

    // พยายามเปลี่ยนชื่อ (ถ้าสิทธิ์ไม่พอจะติด Catch แต่ยังย้ายต่อได้)
    try {
      file.setName(newName);
    } catch(e) {
      console.warn("Cannot rename (Permission): " + newName);
    }

    if (currentParentId === targetFolder.getId()) {
      console.log("Renamed: " + newName + " (at " + targetFolder.getName() + ")");
      return;
    }

    // 1. ลองย้ายเข้าโฟลเดอร์ปลายทางจริง
    try {
      targetFolder.addFile(file);
      if (currentParentId) {
        try {
          DriveApp.getFolderById(currentParentId).removeFile(file);
        } catch(e) {
          // Drive ใช้ parent เดียวตั้งแต่ปี 2020 addFile ย้ายให้แล้ว removeFile อาจไม่จำเป็น
          console.warn("Move limited (remove old parent failed): " + newName);
        }
      }
      console.log("Processed: " + newName + " (at " + targetFolder.getName() + ")");
      return;
    } catch (e) {
      var owner = file.getOwner() ? file.getOwner().getEmail() : "Unknown";
      console.warn("Cannot move to target (owner " + owner + "): " + newName + " — " + e.message);
    }

    // 2. ย้ายไม่ได้จริง ค่อยพักไว้ที่ Unknown
    try {
      unknownFolder.addFile(file);
      console.log("Fallback to Unknown: " + newName);
    } catch (e) {
      console.warn("Cannot move at all: " + newName + " — " + e.message);
    }

  } catch (e) {
    console.warn("Error processing " + newName + ": " + e.message);
  }
}

function getOrCreateSubFolder(parentFolder, name) {
  var folders = parentFolder.getFoldersByName(name);
  if (folders.hasNext()) {
    return folders.next();
  } else {
    return parentFolder.createFolder(name);
  }
}

// Helper: ดึง File ID จาก URL
function extractIdFromUrl(url) {
  var match = url.match(/\/d\/(.*?)\//) || url.match(/id=([^&]+)/) || url.match(/\/d\/([^\/\?]+)/);
  return (match && match[1]) ? match[1] : null;
}

// ลบแถวใน Questions ที่มี img = "require_img" (placeholder ที่ไม่ถูก patch ก่อน import)
// ตั้ง Time Trigger รันทุกวัน 3-4 AM จาก Apps Script console
function cleanupStaging() {
  var doc = SpreadsheetApp.openById(SHEET_ID);
  var sheet = doc.getSheetByName('Questions');
  if (!sheet) return;
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  var deleted = 0;
  // วนลูปจากล่างขึ้นบนเพื่อไม่ให้ index เลื่อน
  for (var i = lastRow; i >= 2; i--) {
    var imgVal = String(sheet.getRange(i, 3).getValue()).trim();
    if (imgVal === 'require_img') {
      sheet.deleteRow(i);
      deleted++;
    }
  }
  console.log('cleanupStaging: deleted ' + deleted + ' require_img rows');
}

/*
   =========================================
   ส่วนที่ 3: KKU IntelSphere Shared Key Pool
   (Idea/interested-using-kkuintel.md — v7)
   =========================================
*/

var INTELSPHERE_SHEET_NAME = "IntelSphere_Keys";
var AI_FEEDBACK_SHEET_NAME = "AI_Feedback";

// ── Feature 4: Related-Questions relations (token-free v1) ─────────────
var QUESTION_RELATIONS_SHEET_NAME = "Question_Relations";
var RELATIONS_TOPK = 5;                 // เก็บสูงสุด k ความสัมพันธ์ต่อข้อ
var RELATIONS_MIN_SCORE = 2;            // เกณฑ์คะแนนขั้นต่ำ (จำนวน token ร่วม) จึงจะถือว่าเกี่ยวข้อง
var RELATIONS_MIN_SHARED_TOKENS = 2;    // prefilter: คู่ที่จะนำมาให้คะแนนต้องแชร์ token >= ค่านี้
var RELATIONS_MAX_POSTINGS = 120;       // token ที่ปรากฏในเอกสารมากกว่านี้ = stopword-like ข้ามไป (กัน N^2 ระเบิด)
var RELATIONS_CANDIDATE_CAP = 60;       // เพดานจำนวน candidate ที่ให้คะแนนต่อข้อ
var RELATIONS_BATCH_BUDGET_MS = 300000; // งบเวลาต่อรอบ ~5 นาที (ต่ำกว่าลิมิต 6 นาทีของ GAS)
var RELATIONS_CHECKPOINT_KEY = "RELATIONS_BATCH_CHECKPOINT";

// ── §1.8: Knowledge-base corpus (Markdown textbook/lecture chunks, token-free v1) ──
// เสิร์ฟผ่าน getKB (doGet, chunked cache); เขียนผ่าน ingestKB (doPost, ต้อง login ก่อน)
var KB_CHUNKS_SHEET_NAME = "KB_Chunks";
var KB_CHUNK_MAX_WORDS = 500;   // section ที่ยาวเกินนี้ถูกตัดเป็น chunk ย่อย (คุม top-k ให้ถูก); ~200-500 คำ/chunk

// ── Feature 2: Glossary (root-word + Thai↔English, §2.1–§2.6) ──
// tap/select miss-path = askGlossaryTerm (public, standalone block, LLM lock-free); เสิร์ฟผ่าน getGlossary
var GLOSSARY_SHEET_NAME = "Glossary";
var GLOSSARY_MODEL = "gemini-3.5-flash-lite";  // flash-lite tier (ปริมาณสูง/ต้นทุนต่ำ); executeChatbotQuery rotate ต่อถ้า Gemini หมดโควต้า
var GLOSSARY_ASK_RATE_LIMIT = 20;         // ต่อ token/'anon' ต่อชั่วโมง (public token-spending + write; กัน spam)
var GLOSSARY_BATCH_BUDGET_MS = 300000;    // งบเวลา nightly ~5 นาที (ต่ำกว่าลิมิต 6 นาทีของ GAS)
var GLOSSARY_CHECKPOINT_KEY = "GLOSSARY_BATCH_CHECKPOINT";
var GLOSSARY_BATCH_CHARS = 6000;          // ขนาดก้อนข้อความต่อ 1 LLM call ตอน batch (กันโพรมป์ยาวเกิน context)

// ── Feature 3: High-yield cram sheet (§3.1–§3.6) ──
// lazy-generate miss-path = generateHighYield (public, standalone block, LLM lock-free) — mirror askGlossaryTerm
// mnemonic vote = voteHighYieldMnemonic (standalone block, 15s tryLock = localized tier); เสิร์ฟผ่าน getHighYield (doGet, chunked cache)
// batch = generateHighYieldForSubject / runHighYieldBatch (checkpointed) + runHighYieldBatchManual (admin tier). trigger เว้นไว้ไม่ติดตั้ง
var HIGHYIELD_SHEET_NAME = "HighYield_Cache";
var HIGHYIELD_MODEL = "gemini-3.5-flash-lite";  // flash-lite tier — เหมือน glossary; executeChatbotQuery rotate ต่อถ้าโควต้าหมด
var HIGHYIELD_MAX_TOKENS = 8192;           // output ก้อนใหญ่ (summary+mnemonics+keywords) — 8192 รองรับ Thai Unicode overhead + JSON เต็ม
var HIGHYIELD_GEN_RATE_LIMIT = 6;          // ต่อ token/'anon' ต่อชั่วโมง (call ใหญ่/แพงกว่า glossary มาก → เข้มกว่า 20)
var HIGHYIELD_VOTE_RATE_LIMIT = 40;        // 👍/🚩 mnemonic เบามาก — กัน spam อย่างเดียว
var HIGHYIELD_MAX_QUESTIONS = 80;          // เพดานจำนวนข้อที่รวมต่อ 1 หมวด (IntelSphere context เล็ก) — เลือกข้อมีเฉลยก่อน
var HIGHYIELD_MAX_CHARS = 6000;            // เพดานตัวอักษรที่ป้อน LLM (บังคับก่อน MAX_QUESTIONS)
var HIGHYIELD_MAX_MNEMONICS = 6;           // จำกัดจำนวน mnemonics ที่เก็บ (คุม output + vote index)
var HIGHYIELD_MAX_KEYWORDS = 15;           // จำกัดจำนวน keywords ที่เก็บ
var HIGHYIELD_BATCH_BUDGET_MS = 300000;    // งบเวลา nightly ~5 นาที (ต่ำกว่าลิมิต 6 นาทีของ GAS)
var HIGHYIELD_CHECKPOINT_KEY = "HIGHYIELD_BATCH_CHECKPOINT";

// ── Feature 6: Frequently-tested keyword index (§6.1–§6.4) — token-free ทั้งหมด (ไม่ยิง LLM เลย) ──
// term set = HighYield keywords ของหมวด ∪ Glossary terms ของวิชา (สกัดจาก 2 pass เดิม, ไม่ pass ที่ 3)
// นับความถี่ lexical: EN → word-boundary regex, TH → substring (ไทยไม่มี word boundary; v1 characteristic)
// เสิร์ฟผ่าน getKeywordIndex (doGet, chunked cache); gen ผ่าน runKeywordIndexBatchManual (admin tier เท่านั้น — §6.2 idle-day/admin, ไม่มี public endpoint, ไม่มี trigger)
var KEYWORD_INDEX_SHEET_NAME = "Keyword_Index";
var KEYWORD_MIN_LEN_EN = 2;   // §6.4 min length กัน keyword สั้น/กำกวม (EN มี word-boundary ป้องกันอยู่แล้ว → 2 พอ เช่น "MI")
var KEYWORD_MIN_LEN_TH = 3;   // TH ใช้ substring (อันตรายกว่า) → ต้องยาว ≥3 กัน match มั่วทั่ว

// ── Feature 4 (main-task): Peer Discussion Thread (per-question) — Idea/active/peer-discussion-thread.md ──
// อ่าน = getDiscussion (doGet, public, cache disc_<qid> 5 นาที — ไม่ rate-limit เพราะ cache กันซ้ำอยู่แล้ว)
// เขียน = postComment/deleteComment (doPost, localized-15s tier, ต้อง login — verifyAnySession)
var DISCUSSION_SHEET_NAME = "Discussion";
var DISCUSSION_MAX_CHARS = 500;        // เพดานความยาวข้อความต่อ 1 comment (บังคับ backend, counter ฝั่ง frontend)
var DISCUSSION_MAX_COMMENTS = 100;     // เพดานจำนวน comment ที่ยัง visible ต่อ 1 คำถาม — เช็คใต้ lock กัน race
var DISCUSSION_CACHE_TTL_SEC = 300;    // 5 นาที ต่อ qid (payload เล็ก ไม่ต้อง chunk)

var INTELSPHERE_ENDPOINT = "https://gen.ai.kku.ac.th/api/v1/chat/completions";
var INTELSPHERE_QUOTA_FLOOR = 0.05; // skip a provider whose remaining < 5% of its daily limit

var INTELSPHERE_LIMITS = {
  "Deepseek": 1000000, "Gemini": 350000, "Nova": 200000, "xAI": 100000,
  "Qwen": 100000, "OpenAI": 200000, "Claude": 200000, "MiniMax": 100000,
  "MoonshotAI": 100000, "Meta": 100000, "Mistral": 100000
  // Perplexity intentionally excluded — no published model ID
};

var INTELSPHERE_PROVIDER_PRIORITY = [
  "Deepseek", "Gemini", "Nova", "xAI", "Qwen", "OpenAI", "Claude", "MiniMax", "MoonshotAI", "Meta", "Mistral"
];

// One flagship model per provider — used ONLY when rotation moves to a provider
// other than the one the student explicitly requested.
var PROVIDER_MODEL_MAP = {
  "Deepseek": "deepseek-v4-pro",  "Gemini": "gemini-3.6-flash",
  "Nova":     "nova-pro-v1",       "xAI":    "grok-4",             "Qwen": "qwen3.7-plus",
  "OpenAI":   "gpt-5-mini",        "Claude": "claude-sonnet-4.5",
  "MiniMax":  "minimax-m3",        "MoonshotAI": "kimi-k3",
  "Meta":     "llama-4-maverick",  "Mistral": "mistral-medium-3"
};

// Hardcoded fallback catalog — used ONLY when the live GET /models fetch fails.
var PROVIDER_MODELS_FALLBACK = {
  "Claude":   ["claude-sonnet-5","claude-sonnet-4.6","claude-sonnet-4.5","claude-haiku-4.5","claude-sonnet-4","claude-3.7-sonnet"],
  "Deepseek": ["deepseek-v4-pro","deepseek-v4-flash","deepseek-v3.2","deepseek-v3.2-exp","deepseek-chat-v3.1"],
  "Gemini":   ["gemini-3.6-flash","gemini-3.5-flash","gemini-3.1-pro-preview","gemini-3.1-flash-lite","gemini-3.1-flash-lite-preview","gemini-3-flash-preview","gemini-2.5-pro","gemini-2.5-flash","gemini-2.5-flash-lite","gemini-3-pro-preview"],
  "Meta":     ["llama-4-maverick","llama-4-scout"],
  "MiniMax":  ["minimax-m3"],
  "Mistral":  ["mistral-small-2603","mistral-large-2512","mistral-medium-3","codestral-2508","codestral-2501"],
  "MoonshotAI": ["kimi-k3"],
  "Nova":     ["nova-2-lite-v1","nova-pro-v1"],
  "OpenAI":   ["gpt-5.4","gpt-5.4-mini","gpt-5.4-nano","gpt-5.2","gpt-5.1","gpt-5.1-codex","gpt-5","gpt-5-mini","gpt-5-nano","gpt-5.5"],
  "Qwen":     ["qwen3.7-plus","qwen3.7-max","qwen3.6-flash","qwen3.5-9b","qwen3-235b-a22b-2507","qwen3-next-80b-a3b-instruct","qwen3-coder-flash","qwen3-coder","qwen3-vl-32b-instruct"],
  "xAI":      ["grok-4.3","grok-4.1-fast","grok-4","grok-3"]
};

