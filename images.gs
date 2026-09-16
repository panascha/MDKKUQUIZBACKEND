/* 
   ========================================================
   ส่วนที่ 3 IMAGE CRUD SYSTEM (Upload, Recycle Bin, Restore)
   ========================================================
*/

// 1. ฟังก์ชันช่วยหาหรือสร้าง Folder (MD > Y[ปี] > [วิชา])
// folderCache: optional {} shared across a batch call — memoizes getFoldersByName lookups
// so repeated images to the same folder don't re-hit Drive every iteration
// CacheService ทำหน้าที่แทน folderCache ข้าม execution: แต่ละ POST รันคนละ container
// folderCache (in-memory) จึงว่างเสมอเมื่อคนละ request → getFoldersByName ยิง Drive 4-5 ครั้งต่อรูป
// อัปโหลดพร้อมกันหลายคนจะชน Drive rate limit (429 / "Service error: Drive")
// โครงสร้างโฟลเดอร์นิ่งเมื่อสร้างแล้ว จึงแคช folder ID ได้ยาว (6 ชม.)
var FOLDER_ID_CACHE_TTL = 21600;

function getOrCreateFolder(parentFolder, folderName, folderCache) {
  var parentId = parentFolder.getId();
  var memKey = parentId + '::' + folderName;
  if (folderCache && folderCache[memKey]) return folderCache[memKey];

  var cache = CacheService.getScriptCache();
  var cacheKey = 'fldr_' + memKey;
  var cachedId = null;
  try { cachedId = cache.get(cacheKey); } catch (e) { /* cache ใช้ไม่ได้ → ตกไปหาแบบเดิม */ }
  if (cachedId) {
    try {
      var hit = DriveApp.getFolderById(cachedId);
      if (!hit.isTrashed()) {
        if (folderCache) folderCache[memKey] = hit;
        return hit;
      }
      cache.remove(cacheKey);
    } catch (e) {
      try { cache.remove(cacheKey); } catch (e2) { /* ignore */ }
    }
  }

  var folders = parentFolder.getFoldersByName(folderName);
  var folder = folders.hasNext() ? folders.next() : parentFolder.createFolder(folderName);
  try { cache.put(cacheKey, folder.getId(), FOLDER_ID_CACHE_TTL); } catch (e) { /* ไม่ critical */ }
  if (folderCache) folderCache[memKey] = folder;
  return folder;
}

// 2. ฟังก์ชันแกะรหัสเพื่อหาที่อยู่โฟลเดอร์
// subjectHint/yearHint: ผู้เรียกที่รู้ปลายทางอยู่แล้ว (เช่น converter ที่ยังไม่ได้บันทึกข้อลงชีต)
// ส่งมาได้เลย → ข้ามการอ่าน Questions/Category/Structure ทั้งใบ (3 full-sheet read ต่อรูป)
function getQuestionRoutingInfo(questionId, subjectHint, yearHint) {
  if (subjectHint && yearHint) {
    return { year: String(yearHint).trim(), subject: String(subjectHint).trim() };
  }

  // เดิมสแกนชีต Questions ทั้งใบ (24k แถว) + Category + Structure ต่อรูป 1 ใบ เพื่อหาวิชา
  // ข้อที่เพิ่งแปลงจาก PDF ยังไม่มีแถวในชีตด้วยซ้ำ → สแกนจบแล้วได้ "General"/"Unknown" อยู่ดี
  // questionId ใช้รูปแบบ <SubjectCode>_<Batch>_<No> เสมอ จึงแกะวิชาจาก prefix แล้วหาปีจาก Structure ที่แคชไว้
  var subjectId = subjectHint
    ? String(subjectHint).trim()
    : ((questionId && String(questionId).indexOf('_') > -1) ? String(questionId).split('_')[0].trim() : "");

  var year = yearHint ? String(yearHint).trim() : "";
  if (!year && subjectId) {
    var structRows = getStructureSheetDataCached();
    for (var i = 1; i < structRows.length; i++) {
      if (String(structRows[i][1]).trim().toUpperCase() === subjectId.toUpperCase()) {
        year = String(structRows[i][0]).trim();
        break;
      }
    }
  }

  return { year: year || "Unknown", subject: subjectId || "General" };
}

// 3. ฟังก์ชันอัปโหลดรูป
// routeCache/folderCache: optional {} shared across a uploadImagesBatch call — memoizes
// getQuestionRoutingInfo (SpreadsheetApp.openById) and folder lookups (DriveApp.getFoldersByName)
// so they don't re-run per image when a batch shares the same question/subject/year
function uploadQuestionImageToDrive(base64Data, questionId, typeIdentifier, subjectHint, yearHint, routeCache, folderCache) {
  var maxRetries = 3;
  var lastError;

  for (var i = 0; i < maxRetries; i++) {
    try {
      var routeInfo;
      if (subjectHint && yearHint) {
        // hint ครบ → ไม่ต้องแตะชีตและไม่ต้องใช้ cache เลย
        routeInfo = getQuestionRoutingInfo(questionId, subjectHint, yearHint);
      } else if (routeCache && routeCache[questionId]) {
        routeInfo = routeCache[questionId];
      } else {
        routeInfo = getQuestionRoutingInfo(questionId);
        if (routeCache) routeCache[questionId] = routeInfo;
      }

      var rootFolder;
      if (folderCache && folderCache['__root__']) {
        rootFolder = folderCache['__root__'];
      } else {
        rootFolder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
        if (folderCache) folderCache['__root__'] = rootFolder;
      }

      var mdFolder = getOrCreateFolder(rootFolder, "MD", folderCache);
      var yearFolder = getOrCreateFolder(mdFolder, "Y" + routeInfo.year, folderCache);
      var targetFolder = getOrCreateFolder(yearFolder, routeInfo.subject, folderCache);

      // Logic จัดเก็บแยกลง Sub-folder
      if (questionId && questionId.indexOf('_') > -1) {
        var parts = questionId.split('_');
        if (parts.length >= 2) {
          var yearType = parts[1].trim();
          if (/^\d{2}/.test(yearType)) {
            var examYear = yearType.substring(0, 2);
            var examGroup = yearType.substring(2) || "General";
            var examYearFolder = getOrCreateFolder(targetFolder, examYear, folderCache);
            targetFolder = getOrCreateFolder(examYearFolder, examGroup, folderCache);
          } else {
            targetFolder = getOrCreateFolder(targetFolder, yearType, folderCache);
          }
        }
      }

      if (typeIdentifier === 'Explain') {
        targetFolder = getOrCreateFolder(targetFolder, 'Explanation', folderCache);
      }

      var mimeType = "image/png";
      var fileExtension = "png";
      if (base64Data.indexOf("data:") === 0) {
        var partsMime = base64Data.split(";")[0].split(":");
        if (partsMime.length > 1) {
          mimeType = partsMime[1];
          if (mimeType === "application/pdf") {
            fileExtension = "pdf";
          } else {
            fileExtension = mimeType.split("/")[1] || "png";
          }
        }
      }

      var cleanBase64 = base64Data.split(',')[1] || base64Data;
      var uniqueID = new Date().getTime() + "_" + Math.floor(Math.random() * 10000);
      var fileName = "Q_" + questionId + "_" + typeIdentifier + "_" + uniqueID + "." + fileExtension;
      
      var blob = Utilities.newBlob(Utilities.base64Decode(cleanBase64), mimeType, fileName);
      var file = targetFolder.createFile(blob);
      
      if (mimeType === 'application/pdf') {
        return 'https://drive.google.com/file/d/' + file.getId() + '/preview';
      } else {
        return 'https://drive.google.com/uc?export=view&id=' + file.getId();
      }

    } catch (e) {
      lastError = e;
      if (e.message.indexOf("Service error") !== -1 || e.message.indexOf("ข้อผิดพลาดของบริการ") !== -1) {
        Utilities.sleep(1500);
        continue;
      }
      throw e;
    }
  }
  throw new Error("Drive ยังคงไม่ตอบสนองหลังจากพยายาม 3 ครั้ง: " + lastError.message);
}

function handleImageTrashAction(fileUrl, user) {
  try {
    var fileId = fileUrl.match(/id=([^&]+)/);
    if (!fileId) return "Invalid URL";
    
    var file = DriveApp.getFileById(fileId[1]);
    var originalParentId = file.getParents().next().getId();
    var rootFolder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
    var recycleFolder = getOrCreateFolder(rootFolder, "RecycleBin");

    recycleFolder.addFile(file);
    DriveApp.getFolderById(originalParentId).removeFile(file);

    var ss = SpreadsheetApp.openById(SHEET_ID);
    var logSheet = ss.getSheetByName("RecycleLog") || ss.insertSheet("RecycleLog");
    if(logSheet.getLastRow() == 0) logSheet.appendRow(["FileID", "Type", "OriginalParentID", "DeletedDate", "User"]);
    
    logSheet.appendRow([fileId[1], "Image", originalParentId, new Date().toISOString(), user]);
    return "Moved to Recycle Bin";
  } catch (e) {
    throw new Error("Trash Error: " + e.message);
  }
}

// 4. ระบบถังขยะ (Soft Delete)
function deleteImageToRecycleBin(fileUrl, user) {
  try {
    var match = fileUrl.match(/id=([^&]+)/) || fileUrl.match(/\/d\/([^\/]+)/);
    if (!match) throw new Error("ID ไฟล์ไม่ถูกต้อง");
    var fileId = match[1];
    
    var file = DriveApp.getFileById(fileId);
    
    // ค้นหาโฟลเดอร์หลัก
    var rootFolder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
    
    // ค้นหาหรือสร้าง RecycleBin
    var recycleFolder;
    var folders = rootFolder.getFoldersByName("RecycleBin");
    if (folders.hasNext()) {
      recycleFolder = folders.next();
    } else {
      recycleFolder = rootFolder.createFolder("RecycleBin");
    }

    // ย้ายไฟล์
    recycleFolder.addFile(file);
    
    // พยายามดึงไฟล์ออกจากโฟลเดอร์เดิมทั้งหมด
    var firstParentId = "";
    var parents = file.getParents();
    while (parents.hasNext()) {
      var parent = parents.next();
      if (firstParentId === "") firstParentId = parent.getId();
      parent.removeFile(file);
    }

    // เขียนบันทึกประวัติเพื่ออนุญาตให้กู้คืนได้ผ่านระบบ Restore
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var logSheet = ss.getSheetByName("RecycleLog") || ss.insertSheet("RecycleLog");
    if(logSheet.getLastRow() == 0) logSheet.appendRow(["FileID", "Type", "OriginalParentID", "DeletedDate", "User"]);
    logSheet.appendRow([fileId, "ExplainMedia", firstParentId, new Date().toISOString(), user]);

    return "Moved to Recycle Bin successfully";
  } catch (e) {
    throw new Error("Drive Error: " + e.message);
  }
}

// 5. ระบบกู้คืน (Restore)
function restoreImageFromRecycleBin(fileUrl) {
  var match = fileUrl.match(/id=([^&]+)/) || fileUrl.match(/\/d\/([^\/]+)/);
  if (!match) throw new Error("URL ไม่ถูกต้อง");
  var fileId = match[1];

  var ss = SpreadsheetApp.openById(SHEET_ID);
  var logSheet = ss.getSheetByName("RecycleLog");
  if (!logSheet) throw new Error("ไม่พบประวัติการลบรูปภาพ");
  
  var data = logSheet.getDataRange().getValues();
  var parentId = "";
  var rowIdx = -1;

  for (var i = 1; i < data.length; i++) {
    if (data[i][0] == fileId) { 
      parentId = data[i][2]; 
      rowIdx = i + 1; 
      break; 
    }
  }

  if (parentId) {
    var file = DriveApp.getFileById(fileId);
    var targetFolder = DriveApp.getFolderById(parentId);
    
    // ย้ายกลับที่เดิม
    targetFolder.addFile(file);
    
    // ลบออกจากโฟลเดอร์ถังขยะ
    var recycleBin = getOrCreateFolder(DriveApp.getFolderById(DRIVE_FOLDER_ID), "RecycleBin");
    try { recycleBin.removeFile(file); } catch(e) {}

    // ลบ Log ออกจาก Sheet
    logSheet.deleteRow(rowIdx);
    
    // ส่งลิงก์ UC กลับไปให้หน้าเว็บใช้งานต่อ
    return 'https://drive.google.com/uc?export=view&id=' + fileId;
  }
  throw new Error("ไม่พบประวัติไฟล์นี้ในถังขยะ (อาจเกิน 10 วันและถูกลบถาวรไปแล้ว)");
}

// 6. ลบถาวรเมื่อเกิน 10 วัน (ตั้ง Trigger รันทุกวัน)
function autoCleanupRecycleBin() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var logSheet = ss.getSheetByName("RecycleLog");
  if(!logSheet) return;
  var data = logSheet.getDataRange().getValues();
  var now = new Date();
  for(var i = data.length - 1; i >= 1; i--) {
    var diff = (now - new Date(data[i][3])) / (1000*60*60*24);
    if(diff > 10) {
      try { DriveApp.getFileById(data[i][0]).setTrashed(true); } catch(e) {}
      logSheet.deleteRow(i + 1);
    }
  }
}

function setupPermissions() {
  DriveApp.getRootFolder();
  Logger.log("ได้รับสิทธิ์เข้าถึง Google Drive เรียบร้อยแล้ว");
}

