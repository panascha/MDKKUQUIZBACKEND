// bulk-questions.gs — bulkDeleteQuestions + bulkSetQuestionCategories (DATABASE Phase 2)
// router-doPost.gs ทำแค่ auth/role/dispatch — logic อยู่ที่นี่ (รับ spreadsheet เพื่อ stub-test ได้)
var QUESTIONS_TRASH_SHEET = 'Questions_Trash';
var BULK_QUESTION_CAP = 100;

// validate + normalize + dedupe ids — คืน {error} หรือ {ids}
function bulkNormalizeIds_(rawIds) {
  if (!Array.isArray(rawIds) || !rawIds.length) return { error: 'ไม่มีรายการ ids' };
  if (rawIds.length > BULK_QUESTION_CAP) return { error: 'เกิน 100 ข้อต่อรอบ — แบ่ง chunk จาก frontend' };
  var seen = {}, ids = [];
  for (var i = 0; i < rawIds.length; i++) {
    var id = String(rawIds[i]).trim();
    if (!id || seen[id]) continue;
    seen[id] = true;
    ids.push(id);
  }
  if (!ids.length) return { error: 'ไม่มีรายการ ids' };
  return { ids: ids };
}

// ย้ายข้อที่เลือกไปชีต Questions_Trash แล้วลบจาก Questions (DEVELOPER only — gate อยู่ที่ router)
// Restore: ยังไม่มี action — กู้คืนด้วยการ copy แถวกลับจาก Questions_Trash ในชีตเอง (ตัด 2 คอลัมน์ท้าย DeletedAt/DeletedBy)
function bulkDeleteQuestionsCore_(doc, rawIds, user, userRole, metadata) {
  var norm = bulkNormalizeIds_(rawIds);
  if (norm.error) return { result: 'error', message: norm.error };
  var ids = norm.ids;

  var qSheet = doc.getSheetByName('Questions');
  var rows = qSheet.getDataRange().getValues();
  var headers = rows[0];
  var nCols = headers.length;
  var want = {};
  for (var w = 0; w < ids.length; w++) want[ids[w]] = true;

  var found = [], foundIds = {}, deletedIds = [];
  for (var i = 1; i < rows.length; i++) {
    var rid = String(rows[i][0]).trim();
    if (!want[rid]) continue;
    found.push({ rowNum: i + 1, id: rid, vals: rows[i] });
    if (!foundIds[rid]) { foundIds[rid] = true; deletedIds.push(rid); }
  }
  if (!found.length) return { result: 'success', applied: 0, skipped: ids.length, deletedIds: [] };

  var trashHeader = headers.concat(['DeletedAt', 'DeletedBy']);
  var tSheet = doc.getSheetByName(QUESTIONS_TRASH_SHEET);
  if (!tSheet) {
    tSheet = doc.insertSheet(QUESTIONS_TRASH_SHEET);
    tSheet.appendRow(trashHeader);
    tSheet.getRange(1, 1, 1, trashHeader.length).setFontWeight('bold').setBackground('#f3f3f3');
  } else if (tSheet.getLastRow() === 0) {
    tSheet.appendRow(trashHeader);
    tSheet.getRange(1, 1, 1, trashHeader.length).setFontWeight('bold').setBackground('#f3f3f3');
  } else {
    var tHead = tSheet.getRange(1, 1, 1, nCols).getValues()[0];
    for (var h = 0; h < nCols; h++) {
      if (String(tHead[h]) !== String(headers[h])) {
        return { result: 'error', message: 'Questions_Trash header ไม่ตรงกับ Questions — ยกเลิก ไม่มีการลบ' };
      }
    }
  }

  // ต้อง append trash ก่อนลบเสมอ — ถ้าพังกลางทางได้แค่แถวซ้ำ ไม่มีของหาย
  var now = new Date();
  var out = found.map(function (f) { return f.vals.concat([now, user]); });
  tSheet.getRange(tSheet.getLastRow() + 1, 1, out.length, nCols + 2).setValues(out);
  SpreadsheetApp.flush();

  // ลบล่างขึ้นบน (deleteRows ทำให้เลขแถวเลื่อน) — รวมแถวติดกันเป็น run
  var nums = found.map(function (f) { return f.rowNum; }).sort(function (a, b) { return b - a; });
  var s = 0;
  while (s < nums.length) {
    var e = s;
    while (e + 1 < nums.length && nums[e + 1] === nums[e] - 1) e++;
    qSheet.deleteRows(nums[e], e - s + 1);
    s = e + 1;
  }

  updateVersion();
  // ไม่ log เนื้อแถว (appendRow จำกัด 50,000 ตัวอักษร) — เนื้ออยู่ใน Questions_Trash
  writeAdminLog(user, userRole, 'QUESTION', 'DELETE', deletedIds.join(','),
    'Bulk delete -> Questions_Trash (' + deletedIds.length + ' items)',
    { movedTo: QUESTIONS_TRASH_SHEET, ids: deletedIds }, 'DELETED', metadata);
  sbMirrorQuestionsDeleted_(deletedIds);

  return { result: 'success', applied: deletedIds.length, skipped: ids.length - deletedIds.length, deletedIds: deletedIds };
}

// แทนที่ category ทั้งรายการของข้อที่เลือก (replace — ต่างจาก bulkAddQuestionCategories ที่ append)
function bulkSetQuestionCategoriesCore_(doc, rawIds, rawCatIds, user, userRole, metadata) {
  var norm = bulkNormalizeIds_(rawIds);
  if (norm.error) return { result: 'error', message: norm.error };
  var ids = norm.ids;

  var catIds = [], catSeen = {};
  if (Array.isArray(rawCatIds)) {
    for (var c = 0; c < rawCatIds.length; c++) {
      var cid = String(rawCatIds[c]).trim();
      if (!cid || catSeen[cid]) continue;
      if (cid.indexOf('///') !== -1 || cid.indexOf('"') !== -1) {
        return { result: 'error', message: 'category id ไม่ถูกต้อง' };
      }
      catSeen[cid] = true;
      catIds.push(cid);
    }
  }
  if (catIds.length < 1 || catIds.length > 20) return { result: 'error', message: 'ต้องเลือก category 1-20 รายการ' };

  var catSheet = doc.getSheetByName('Category');
  var catRows = catSheet.getDataRange().getValues();
  var known = {};
  for (var k = 1; k < catRows.length; k++) known[String(catRows[k][0]).trim()] = true;
  var unknown = catIds.filter(function (x) { return !known[x]; });
  if (unknown.length) return { result: 'error', message: 'ไม่พบ category: ' + unknown.slice(0, 5).join(', ') };

  var qSheet = doc.getSheetByName('Questions');
  var qData = qSheet.getDataRange().getValues();
  var qIdMap = {};
  for (var r = 1; r < qData.length; r++) qIdMap[String(qData[r][0]).trim()] = r + 1;

  var newJson = JSON.stringify(catIds);
  var applied = 0, skipped = 0;
  var appliedIds = [], oldMap = {}, finalMap = {}, mirrorRows = [];
  var sbChanged = false;
  for (var n = 0; n < ids.length; n++) {
    var id = ids[n];
    var rowIdx = qIdMap[id];
    if (!rowIdx) { skipped++; continue; }

    var oldRaw = String(qData[rowIdx - 1][6] || '');
    var oldCats = [];
    try { oldCats = oldRaw ? JSON.parse(oldRaw.replace(/'/g, '"')) : []; }
    catch (e) { oldCats = oldRaw ? [oldRaw] : []; }
    if (JSON.stringify(oldCats) === newJson) { skipped++; continue; }

    qSheet.getRange(rowIdx, 7).setValue(newJson);
    var split = null;
    try { split = autoCreateSplitCategories(id, catIds.slice(), true, rowIdx); } // skipSort=true — sort ทีเดียวตอนจบ
    catch (e2) { console.log('Split error in bulkSetQuestionCategories: ' + e2); }
    applied++;
    appliedIds.push(id);
    oldMap[id] = oldRaw;
    finalMap[id] = (split && split.finalCategories) || catIds;
    mirrorRows.push({ questionId: id, category: finalMap[id] });
    if (split && split.sheetsChanged) sbChanged = true;
  }

  if (applied > 0) {
    try { sortCategorySheet(); } catch (e3) { console.log('Sort error in bulkSetQuestionCategories: ' + e3); }
    updateVersion();
    sbMirrorQuestionRows_(mirrorRows);
    if (sbChanged) {
      sbMarkSheet_('Category');
      sbMarkSheet_('Structure');
    }
    writeAdminLog(user, userRole, 'QUESTION', 'BULK_SET_CATEGORY', appliedIds.join(','),
      'Bulk set categories (' + ids.length + ' items)', oldMap,
      { categoryIds: catIds, applied: applied, skipped: skipped }, metadata);
  }

  return { result: 'success', applied: applied, skipped: skipped, finalCategories: finalMap };
}
