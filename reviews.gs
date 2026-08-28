/* =========================================================================
   COURSE REVIEWS — รีวิว/ให้ดาววิชา (public read + student write + admin moderate)
   Sheet: Reviews [A-J] (แช่แข็ง ห้ามสลับ/เพิ่ม/ลบคอลัมน์)
     Timestamp | SubjectID | Rating | ReviewText | DisplayName |
     StudentYearAtReview | StudentIdHash | IsAnonymous | Status | AdminNote
   - getReviews         : public, chunked cache (cache.gs::getReviewsDataCached), Approved-only,
                          avg เฉพาะเมื่อ totalReviews >= 3
   - submitReview       : localized-15s, verifyAnySession + rl_review_ 5/hr → upsert (idHash+subject)
                          → year snapshot ฝั่ง server (แช่ในชีต) → invalidate v_reviews
   - updateReviewStatus : admin 25s, flip Status + prepend AdminNote
   - getReviewsForAdmin : dual-auth read (ทุกสถานะ + rowIndex/hash เป็น identity)
   แผน: mdkkuquiz-reviews-donations-handoff.md (Q7–Q8 locked)
   ========================================================================= */

var REVIEW_STATUS_WHITELIST = ['Approved', 'Hidden', 'Rejected'];
var REVIEW_MIN_COUNT_FOR_AVG = 3; // Q7c: ต่ำกว่านี้ซ่อนค่าเฉลี่ยตัวเลข (ยังโชว์การ์ดรายรีวิว)

// helper JSON output ร่วมของ reviews + donations (GAS global scope — โหลดก่อน router ได้)
function rdJson_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// idempotent: สร้างชีต Reviews + header ถ้ายังไม่มี (ไม่ seed — seed อยู่ setupReviewsWithSeed_)
function setupReviewsSheet() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName("Reviews");
  if (!sheet) sheet = ss.insertSheet("Reviews");
  var headers = ["Timestamp", "SubjectID", "Rating", "ReviewText", "DisplayName", "StudentYearAtReview", "StudentIdHash", "IsAnonymous", "Status", "AdminNote"];
  if (!sheet.getRange(1, 1).getValue()) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#e6ffe6");
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// GET setupReviews (one-off): สร้างชีต + seed mock ~6 แถวใต้วิชา 'DEMO' (ถ้าว่าง)
// ใช้วิชา DEMO โดยตั้งใจ — เพื่อไม่ให้ค่าเฉลี่ยปลอมไปปนวิชาจริงเลย (ลบทิ้งง่าย = ลบแถว DEMO)
function setupReviewsWithSeed_() {
  var sheet = setupReviewsSheet();
  if (sheet.getLastRow() <= 1) {
    var t = new Date().toISOString();
    var mock = [
      [t, 'DEMO', 5, 'เนื้อหาแน่น อาจารย์สอนดีมาก', 'พี่ปีสาม', 'ปี 3', 'seed_hash_1', false, 'Approved', 'SEED'],
      [t, 'DEMO', 4, 'ข้อสอบยากแต่ยุติธรรม', '', 'ปี 2', 'seed_hash_2', true, 'Approved', 'SEED'],
      [t, 'DEMO', 4, 'สไลด์ครบ ทบทวนง่าย', 'anon', 'ปี 4', 'seed_hash_3', true, 'Approved', 'SEED'],
      [t, 'DEMO', 5, 'lab สนุก ได้ลงมือทำจริง', 'หมอน้อย', 'ปี 3', 'seed_hash_4', false, 'Approved', 'SEED'],
      [t, 'DEMO', 3, 'เนื้อหาเยอะไปนิด', '', 'ศิษย์เก่า', 'seed_hash_5', true, 'Approved', 'SEED'],
      [t, 'DEMO', 4, 'แนะนำให้อ่าน sheet ก่อนเข้าเรียน', 'รุ่นพี่', 'ปี 5', 'seed_hash_6', false, 'Approved', 'SEED']
    ];
    sheet.getRange(2, 1, mock.length, 10).setValues(mock);
    updateReviewsVersion();
  }
  return rdJson_({ result: 'success', rows: sheet.getLastRow() - 1, note: 'seed อยู่ใต้วิชา DEMO — ลบก่อน go-live (Step 3 prereq)' });
}

// year snapshot ฝั่ง server (Q8a) — fixed UTC+7 offset, ไม่ใช้ toLocaleString re-parse
// คืน 'ปี N' | 'ศิษย์เก่า' | '' (ไม่มี/แปลงไม่ได้ → ไม่โชว์ badge)
function computeStudentYearAtReview_(studentId) {
  var sid = String(studentId || '').trim();
  if (!/^\d{2}/.test(sid)) return '';
  var bangkokOffsetMs = 7 * 60 * 60 * 1000; // ไทยไม่มี DST
  var now = new Date(Date.now() + bangkokOffsetMs); // อ่าน .getUTCMonth()/.getUTCFullYear()
  var academicYearBE = now.getUTCFullYear() + 543 - (now.getUTCMonth() >= 5 ? 0 : 1); // ตัดปีการศึกษาที่ มิ.ย. (index 5)
  var entryYearBE = 2500 + parseInt(sid.slice(0, 2), 10);
  var yearLevel = academicYearBE - entryYearBE + 1;
  if (isNaN(yearLevel)) return '';
  if (yearLevel > 6) return 'ศิษย์เก่า';
  if (yearLevel >= 1) return 'ปี ' + yearLevel;
  return '';
}

// SHA-256 hex ของ input + STUDENT_ID_SALT (dedup key — salt อยู่ฝั่ง server เท่านั้น)
function hashStudentId_(input) {
  var bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(input) + STUDENT_ID_SALT, Utilities.Charset.UTF_8);
  var hex = '';
  for (var i = 0; i < bytes.length; i++) {
    var b = (bytes[i] < 0 ? bytes[i] + 256 : bytes[i]).toString(16);
    hex += (b.length === 1 ? '0' : '') + b;
  }
  return hex;
}

// หา StudentID ของผู้ใช้จากอีเมล (ชีต Admins เก็บผู้ใช้ SSO ทุกคน; col6=email idx5, col9=StudentID idx8)
// อ่านฝั่ง server เท่านั้น — ห้ามรับ studentId จาก client (ปลอมได้ = ปลอมทั้ง year badge + dedup key)
function lookupStudentIdByEmail_(doc, email) {
  var sheet = doc.getSheetByName("Admins");
  if (!sheet) return '';
  var rows = sheet.getDataRange().getValues();
  var target = String(email).trim().toLowerCase();
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][5]).trim().toLowerCase() === target) {
      return String(rows[i][8] || '').trim();
    }
  }
  return '';
}

// public read: Approved-only, project เฉพาะ field ที่ปลอดภัย (ห้ามส่ง StudentIdHash col G)
// avg เฉพาะเมื่อ totalReviews >= 3 (Q7c); ต่ำกว่านั้น avgRating = null (frontend โชว์ '—'/'ใหม่')
function getReviewsData(subjectId, startTime) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName("Reviews");
  var subjFilter = subjectId ? String(subjectId).trim().toUpperCase() : '';
  var out = { result: 'success', subject: subjFilter || 'all', avgRating: null, totalReviews: 0, reviews: [] };
  if (!sheet || sheet.getLastRow() <= 1) {
    return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
  }
  var values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 10).getValues();
  var sum = 0, n = 0;
  for (var i = 0; i < values.length; i++) {
    var row = values[i];
    if (String(row[8]) !== 'Approved') continue; // col I Status
    if (subjFilter && String(row[1]).trim().toUpperCase() !== subjFilter) continue;
    var rating = Number(row[2]) || 0;
    var isAnon = row[7] === true || String(row[7]).toLowerCase() === 'true';
    out.reviews.push({
      timestamp: row[0] instanceof Date ? row[0].toISOString() : String(row[0]),
      rating: rating,
      reviewText: String(row[3] || ''),
      displayName: String(row[4] || ''), // write path การันตี: anon = nickname-or-empty, ไม่มีชื่อจริงอัตโนมัติ; ว่าง → frontend เรนเดอร์ "นิรนาม"
      yearLabel: String(row[5] || ''),
      isAnonymous: isAnon
    });
    sum += rating; n++;
  }
  out.totalReviews = n;
  if (n >= REVIEW_MIN_COUNT_FOR_AVG) out.avgRating = Math.round((sum / n) * 10) / 10;
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

// student write (localized-15s) — auth+compute นอก lock, lock เฉพาะ dedup-check + upsert
function submitReview(doc, data) {
  var user = verifyAnySession(data.sessionToken);
  if (!user || !user.email) return rdJson_({ result: 'error', message: 'login_required' });

  var subjectId = String(data.subjectId || '').trim().toUpperCase();
  var rating = parseInt(data.rating, 10);
  if (!subjectId) return rdJson_({ result: 'error', message: 'กรุณาเลือกวิชา' });
  if (!(rating >= 1 && rating <= 5)) return rdJson_({ result: 'error', message: 'กรุณาให้คะแนน 1-5 ดาว' });

  var reviewText = String(data.reviewText || '').trim().slice(0, 2000).replace(/\/\/\//g, '/').replace(/\r\n/g, '\n');
  var isAnon = data.isAnonymous === true || String(data.isAnonymous).toLowerCase() === 'true';
  // Q8b: เก็บเฉพาะ nickname ที่ผู้ใช้พิมพ์เอง (ทั้ง anon และไม่ anon) — ไม่ดึงชื่อ Google มาเก็บ = ไม่มี PII อัตโนมัติ
  var displayName = String(data.displayName || '').trim().slice(0, 60).replace(/\/\/\//g, '/');

  // studentId lookup ฝั่ง server (ไม่เชื่อ client). ไม่มีรหัส → hash จากอีเมลแทน (ยังกัน dup ได้) + ไม่โชว์ปี
  var studentId = lookupStudentIdByEmail_(doc, user.email);
  var idHash = studentId ? hashStudentId_(studentId) : hashStudentId_('email:' + String(user.email).toLowerCase());
  var yearLabel = studentId ? computeStudentYearAtReview_(studentId) : '';

  var lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) return rdJson_({ result: 'error', message: 'เซิร์ฟเวอร์ไม่ตอบสนองเนื่องจากโหลดสูง กรุณาลองใหม่' });
  try {
    var sheet = setupReviewsSheet();
    var values = sheet.getLastRow() > 1 ? sheet.getRange(2, 1, sheet.getLastRow() - 1, 10).getValues() : [];
    var foundRow = -1;
    for (var i = 0; i < values.length; i++) {
      if (String(values[i][6]) === idHash && String(values[i][1]).trim().toUpperCase() === subjectId) { foundRow = i + 2; break; }
    }
    var nowIso = new Date().toISOString();
    var rowVals = [nowIso, subjectId, rating, reviewText, displayName, yearLabel, idHash, isAnon, 'Approved', ''];
    if (foundRow > -1) {
      rowVals[9] = String(sheet.getRange(foundRow, 10).getValue() || ''); // คง AdminNote เดิม (Q8d append-only)
      sheet.getRange(foundRow, 1, 1, 10).setValues([rowVals]); // แก้ = กลับเป็น Approved ให้ moderator ตรวจซ้ำ
    } else {
      sheet.appendRow(rowVals);
    }
    updateReviewsVersion();
    return rdJson_({ result: 'success', updated: foundRow > -1, yearLabel: yearLabel });
  } finally {
    lock.releaseLock();
  }
}

// admin moderate (25s) — flip Status + prepend AdminNote. auth ทำใน router (dual-auth) ก่อนเรียก
function updateReviewStatus(doc, data) {
  var newStatus = String(data.status || '').trim();
  if (REVIEW_STATUS_WHITELIST.indexOf(newStatus) < 0) return rdJson_({ result: 'error', message: 'status ไม่ถูกต้อง' });
  var rowIndex = parseInt(data.rowIndex, 10);
  var idHash = String(data.studentIdHash || '').trim();
  var subjectId = String(data.subjectId || '').trim().toUpperCase();
  var noteText = String(data.note || '').trim().slice(0, 500);

  var lock = LockService.getScriptLock();
  if (!lock.tryLock(25000)) return rdJson_({ result: 'error', message: 'ระบบไม่ว่าง กรุณาลองใหม่' });
  try {
    var sheet = doc.getSheetByName("Reviews");
    if (!sheet || sheet.getLastRow() <= 1) return rdJson_({ result: 'error', message: 'ไม่พบชีต Reviews' });
    var targetRow = -1;
    if (rowIndex > 1) targetRow = rowIndex;
    else {
      var vals = sheet.getRange(2, 1, sheet.getLastRow() - 1, 10).getValues();
      for (var i = 0; i < vals.length; i++) {
        if (String(vals[i][6]) === idHash && String(vals[i][1]).trim().toUpperCase() === subjectId) { targetRow = i + 2; break; }
      }
    }
    if (targetRow < 2) return rdJson_({ result: 'error', message: 'ไม่พบรีวิว' });
    sheet.getRange(targetRow, 9).setValue(newStatus); // col I
    if (noteText) {
      var prev = String(sheet.getRange(targetRow, 10).getValue() || '');
      var stamp = '[' + Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM-dd HH:mm') + ' Admin] ' + noteText;
      sheet.getRange(targetRow, 10).setValue(prev ? (stamp + '\n' + prev) : stamp); // prepend ล่าสุดบนสุด (Q8d)
    }
    updateReviewsVersion();
    return rdJson_({ result: 'success' });
  } finally {
    lock.releaseLock();
  }
}

// admin read: ทุกสถานะ + rowIndex/hash (identity สำหรับปุ่ม moderate). auth ทำใน router ก่อนเรียก
function getReviewsForAdmin() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName("Reviews");
  var rows = [];
  if (sheet && sheet.getLastRow() > 1) {
    var values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 10).getValues();
    for (var i = 0; i < values.length; i++) {
      rows.push({
        rowIndex: i + 2,
        timestamp: values[i][0] instanceof Date ? values[i][0].toISOString() : String(values[i][0]),
        subjectId: values[i][1], rating: values[i][2], reviewText: values[i][3],
        displayName: values[i][4], yearLabel: values[i][5], studentIdHash: values[i][6],
        isAnonymous: values[i][7], status: values[i][8], adminNote: values[i][9]
      });
    }
  }
  return rdJson_({ result: 'success', reviews: rows });
}
