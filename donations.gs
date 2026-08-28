/* =========================================================================
   SLIP-VERIFIED DONATIONS — บริจาคค่าเซิร์ฟเวอร์ (write-only จาก REAL → admin audit)
   ★ ไม่มี public wall / ไม่มี getPublicDonations (Q6 scrapped) — สลิปเป็นความลับ owner-only
   Sheet: Donations [A-J] (แช่แข็ง)
     Timestamp | DonorName | Amount | TransRef | RecipientMatch | Message |
     IsAnonymous | SlipDriveUrl | Status | AdminNote
   - submitDonation      : localized-15s. OCR (Gemini) + Drive upload "นอก lock" (invariant:
                           ห้าม UrlFetchApp ใต้ LockService) → dedup(TransRef) + append ใต้ lock
   - updateDonationStatus: admin 25s override (MarkVerified/Hide/Delete/…)
   - getDonationsForAdmin : dual-auth read (รวม SlipDriveUrl — admin เท่านั้น)
   สถานะ: SlipReviewed(auto high) | PendingAdmin(ชื่อคลุมเครือ) | Rejected | Verified | Hidden | Deleted
   แผน: mdkkuquiz-reviews-donations-handoff.md (Q1–Q5 locked) + decision 2 (ชื่อคลุมเครือ → PendingAdmin ไม่ reject)
   ========================================================================= */

var DONATION_STATUS_WHITELIST = ['SlipReviewed', 'PendingAdmin', 'Rejected', 'Verified', 'Hidden', 'Deleted'];

// idempotent: สร้างชีต Donations + header ถ้ายังไม่มี
function setupDonationsSheet() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName("Donations");
  if (!sheet) sheet = ss.insertSheet("Donations");
  var headers = ["Timestamp", "DonorName", "Amount", "TransRef", "RecipientMatch", "Message", "IsAnonymous", "SlipDriveUrl", "Status", "AdminNote"];
  if (!sheet.getRange(1, 1).getValue()) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#fff0e6");
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// เซฟสลิปลง subfolder "Donations" ใต้ DRIVE_FOLDER_ID (owner-only เหมือน Feedback)
// ★ URL นี้ห้ามส่งออกใน response สาธารณะใดๆ — ไม่มี endpoint public donation อยู่แล้ว, admin เท่านั้นอ่านผ่าน getDonations
function saveDonationSlipToDrive_(base64Data) {
  var rootFolder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
  var folder = getOrCreateFolder(rootFolder, "Donations");
  var mimeType = "image/jpeg", fileExtension = "jpg";
  if (base64Data.indexOf("data:") === 0) {
    var partsMime = base64Data.split(";")[0].split(":");
    if (partsMime.length > 1 && partsMime[1].indexOf("image/") === 0) {
      mimeType = partsMime[1];
      fileExtension = mimeType.split("/")[1] || "jpg";
    }
  }
  var cleanBase64 = base64Data.split(',')[1] || base64Data;
  var fileName = "SLIP_" + new Date().getTime() + "_" + Math.floor(Math.random() * 10000) + "." + fileExtension;
  var blob = Utilities.newBlob(Utilities.base64Decode(cleanBase64), mimeType, fileName);
  var file = folder.createFile(blob);
  return 'https://drive.google.com/uc?export=view&id=' + file.getId();
}

// normalize ชื่อไทย/อังกฤษ: ตัดคำนำหน้า + ตัด mask (* x .) → คงเฉพาะตัวที่เห็น
function normalizeThaiName_(s) {
  s = String(s || '');
  s = s.replace(/(นางสาว|นาย|นาง|น\.ส\.|ด\.ช\.|ด\.ญ\.|เด็กชาย|เด็กหญิง|Mr\.?|Mrs\.?|Ms\.?|Miss)\s*/gi, '');
  s = s.replace(/[*xX×•]+/g, ' '); // mask ปิดบัง → ช่องว่าง
  s = s.replace(/\s+/g, ' ').trim().toLowerCase();
  return s;
}

// เทียบชื่อผู้รับ (อาจถูก bank ปิดบัง เช่น "ปาณั*** จ***") กับ DONATION_RECIPIENT_NAME
// คืน 'high' | 'low' | 'none' — ไม่เคยตัดสิน reject (decision 2: ชื่อคลุมเครือ → PendingAdmin เสมอ)
function checkRecipientMatch_(rawName) {
  var raw = normalizeThaiName_(rawName);
  if (!raw) return 'none';
  var targetTokens = normalizeThaiName_(DONATION_RECIPIENT_NAME).split(' ').filter(Boolean);
  var rawTokens = raw.split(' ').filter(Boolean);
  if (!rawTokens.length || !targetTokens.length) return 'none';
  var matched = 0;
  for (var i = 0; i < targetTokens.length; i++) {
    for (var j = 0; j < rawTokens.length; j++) {
      var rt = rawTokens[j], tt = targetTokens[i];
      var shorter = Math.min(rt.length, tt.length);
      if (shorter >= 2 && (tt.indexOf(rt) === 0 || rt.indexOf(tt) === 0)) { matched++; break; }
    }
  }
  if (matched >= targetTokens.length) return 'high';
  if (matched >= 1) return 'low';
  return 'none';
}

// student/anon write (localized-15s). ลำดับ: OCR → validate → dedup precheck → Drive → lock → dedup recheck → append
// OCR ก่อน Drive โดยตั้งใจ — path reject/dup จะไม่สร้างไฟล์ค้างใน Drive เลย (มีแค่ race หายากที่อาจค้าง 1 ไฟล์)
function submitDonation(doc, data) {
  var email = 'anonymous';
  if (data.sessionToken) { var u = verifyAnySession(data.sessionToken); if (u && u.email) email = u.email; }

  var slipB64 = String(data.slipImage || '');
  if (slipB64.indexOf('data:image') !== 0) return rdJson_({ result: 'error', message: 'กรุณาแนบรูปสลิปโอนเงิน' });
  if (slipB64.length > 8 * 1024 * 1024) return rdJson_({ result: 'error', message: 'รูปสลิปใหญ่เกินไป กรุณาลดขนาดภาพ' });

  var donorName = String(data.donorName || '').trim().slice(0, 80).replace(/\/\/\//g, '/');
  var isAnon = data.isAnonymous === true || String(data.isAnonymous).toLowerCase() === 'true';
  var message = String(data.message || '').trim().slice(0, 500).replace(/\/\/\//g, '/').replace(/\r\n/g, '\n');

  // (1) OCR — นอก lock
  var ocr = callGeminiForSlipOCR(slipB64);
  if (!ocr.ok) return rdJson_({ result: 'error', message: ocr.error || 'อ่านสลิปไม่สำเร็จ กรุณาถ่ายใหม่ให้ชัด' });
  var d = ocr.data || {};
  var transRef = String(d.transRef || '').replace(/\s+/g, '').trim();
  var amount = Number(String(d.amount).replace(/[^\d.]/g, '')) || 0; // strip คอมม่า/บาท/ช่องว่าง กันยอด "1,500" → NaN → reject ผิด

  // (2) strict reject — เฉพาะ transRef อ่านไม่ได้ หรือ amount<=0 (ชื่อคลุมเครือ "ไม่" reject)
  if (!transRef || transRef.length < 6) return rdJson_({ result: 'error', message: 'สลิปไม่ชัดเจน กรุณาถ่ายใหม่ให้เห็นเลขอ้างอิง (Ref) ให้ชัด' });
  if (amount <= 0) return rdJson_({ result: 'error', message: 'อ่านยอดเงินไม่ได้ กรุณาถ่ายสลิปให้เห็นจำนวนเงินชัดเจน' });

  // (3) dedup precheck — นอก lock (กันอัป Drive ทิ้งถ้าซ้ำชัดๆ)
  var sheet0 = doc.getSheetByName("Donations");
  if (sheet0 && sheet0.getLastRow() > 1) {
    var refCol = sheet0.getRange(2, 4, sheet0.getLastRow() - 1, 1).getValues();
    for (var i = 0; i < refCol.length; i++) {
      if (String(refCol[i][0]).replace(/\s+/g, '') === transRef) return rdJson_({ result: 'error', message: 'สลิปนี้ถูกบันทึกไว้แล้ว ขอบคุณสำหรับการสนับสนุน 🙏' });
    }
  }

  // (4) recipient match → tri-state (decision 2: high+model ตรง = auto SlipReviewed, ที่เหลือ = PendingAdmin)
  var confidence = checkRecipientMatch_(d.recipientNameRaw);
  var modelSaysMatch = d.isRecipientMatch === true || String(d.isRecipientMatch).toLowerCase() === 'true';
  var status = (confidence === 'high' && modelSaysMatch) ? 'SlipReviewed' : 'PendingAdmin';

  // (5) Drive upload — นอก lock
  var slipUrl = '';
  try { slipUrl = saveDonationSlipToDrive_(slipB64); } catch (err) { slipUrl = 'upload_failed'; }

  // (6) lock 15s: dedup recheck (กัน race) + append 1 แถว
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) return rdJson_({ result: 'error', message: 'เซิร์ฟเวอร์ไม่ตอบสนอง กรุณาลองใหม่' });
  try {
    var sheet = setupDonationsSheet();
    if (sheet.getLastRow() > 1) {
      var refs = sheet.getRange(2, 4, sheet.getLastRow() - 1, 1).getValues();
      for (var k = 0; k < refs.length; k++) {
        if (String(refs[k][0]).replace(/\s+/g, '') === transRef) return rdJson_({ result: 'error', message: 'สลิปนี้ถูกบันทึกไว้แล้ว ขอบคุณสำหรับการสนับสนุน 🙏' });
      }
    }
    var autoNote = (status === 'PendingAdmin')
      ? '[AUTO] recipient confidence=' + confidence + (slipUrl === 'upload_failed' ? ' | slip upload FAILED' : '')
      : (slipUrl === 'upload_failed' ? '[AUTO] slip upload FAILED' : '');
    sheet.appendRow([new Date().toISOString(), donorName, amount, transRef, confidence, message, isAnon, slipUrl, status, autoNote]);
    return rdJson_({ result: 'success', status: status, amount: amount, transRef: transRef });
  } finally {
    lock.releaseLock();
  }
}

// admin override (25s). auth ทำใน router (dual-auth) ก่อนเรียก
function updateDonationStatus(doc, data) {
  var newStatus = String(data.status || '').trim();
  if (DONATION_STATUS_WHITELIST.indexOf(newStatus) < 0) return rdJson_({ result: 'error', message: 'status ไม่ถูกต้อง' });
  var rowIndex = parseInt(data.rowIndex, 10);
  var transRef = String(data.transRef || '').replace(/\s+/g, '').trim();
  var noteText = String(data.note || '').trim().slice(0, 500);

  var lock = LockService.getScriptLock();
  if (!lock.tryLock(25000)) return rdJson_({ result: 'error', message: 'ระบบไม่ว่าง กรุณาลองใหม่' });
  try {
    var sheet = doc.getSheetByName("Donations");
    if (!sheet || sheet.getLastRow() <= 1) return rdJson_({ result: 'error', message: 'ไม่พบชีต Donations' });
    var targetRow = -1;
    if (rowIndex > 1) targetRow = rowIndex;
    else if (transRef) {
      var vals = sheet.getRange(2, 4, sheet.getLastRow() - 1, 1).getValues();
      for (var i = 0; i < vals.length; i++) {
        if (String(vals[i][0]).replace(/\s+/g, '') === transRef) { targetRow = i + 2; break; }
      }
    }
    if (targetRow < 2) return rdJson_({ result: 'error', message: 'ไม่พบรายการบริจาค' });
    sheet.getRange(targetRow, 9).setValue(newStatus); // col I
    if (noteText) {
      var prev = String(sheet.getRange(targetRow, 10).getValue() || '');
      var stamp = '[' + Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM-dd HH:mm') + ' Admin] ' + noteText;
      sheet.getRange(targetRow, 10).setValue(prev ? (stamp + '\n' + prev) : stamp);
    }
    return rdJson_({ result: 'success' });
  } finally {
    lock.releaseLock();
  }
}

// admin read: ทุกแถว รวม SlipDriveUrl (admin เท่านั้น). auth ทำใน router ก่อนเรียก
function getDonationsForAdmin() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName("Donations");
  var rows = [];
  if (sheet && sheet.getLastRow() > 1) {
    var values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 10).getValues();
    for (var i = 0; i < values.length; i++) {
      rows.push({
        rowIndex: i + 2,
        timestamp: values[i][0] instanceof Date ? values[i][0].toISOString() : String(values[i][0]),
        donorName: values[i][1], amount: values[i][2], transRef: values[i][3],
        recipientMatch: values[i][4], message: values[i][5], isAnonymous: values[i][6],
        slipDriveUrl: values[i][7], status: values[i][8], adminNote: values[i][9]
      });
    }
  }
  return rdJson_({ result: 'success', donations: rows });
}
