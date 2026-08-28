// --- Caching Engine for Large Value Chunking (Apps Script 100KB Limit Workaround) ---

function getVersionCached() {
  var cache = CacheService.getScriptCache();
  var v = cache.get("v_cache");
  if (v == null) {
    v = PropertiesService.getScriptProperties().getProperty('v') || "0";
    try {
      cache.put("v_cache", v, 60); // Cache version check for 60 seconds
    } catch(e) {
      console.warn("Version cache write error: " + e.message);
    }
  }
  return v;
}

function getVotesVersionCached() {
  var cache = CacheService.getScriptCache();
  var v = cache.get("v_votes_cache");
  if (v == null) {
    v = PropertiesService.getScriptProperties().getProperty('v_votes') || "0";
    try {
      cache.put("v_votes_cache", v, 60); // Cache votes version for 60 seconds
    } catch (e) {
      console.warn("Votes version cache write error: " + e.message);
    }
  }
  return v;
}

function updateVotesVersion() {
  var cache = PropertiesService.getScriptProperties();
  var newVer = new Date().getTime().toString();
  cache.setProperty('v_votes', newVer);
  try {
    CacheService.getScriptCache().put("v_votes_cache", newVer, 60);
  } catch (e) {
    console.warn("Votes version cache put failed: " + e.message);
  }
  return newVer;
}

// --- Reviews version key (แยกจาก v/v_votes) — submitReview/updateReviewStatus bump แล้วล้าง cache รีวิวทันที ---
function getReviewsVersionCached() {
  var cache = CacheService.getScriptCache();
  var v = cache.get("v_reviews_cache");
  if (v == null) {
    v = PropertiesService.getScriptProperties().getProperty('v_reviews') || "0";
    try { cache.put("v_reviews_cache", v, 60); } catch (e) { console.warn("Reviews version cache write error: " + e.message); }
  }
  return v;
}

function updateReviewsVersion() {
  var newVer = new Date().getTime().toString();
  PropertiesService.getScriptProperties().setProperty('v_reviews', newVer);
  try { CacheService.getScriptCache().put("v_reviews_cache", newVer, 60); } catch (e) { console.warn("Reviews version cache put failed: " + e.message); }
  return newVer;
}

// public getReviews cache (chunked, 10 นาที) — คีย์ผูก v_reviews จึงถูกล้างทันทีเมื่อมีรีวิวใหม่/เปลี่ยนสถานะ
function getReviewsDataCached(subjectId, startTime) {
  var v = getReviewsVersionCached();
  var cleanFilter = subjectId ? String(subjectId).trim().toUpperCase() : "all";
  var cacheKey = "reviews_" + v + "_" + cleanFilter;
  var cachedStr = getLargeCache(cacheKey);
  if (cachedStr != null) return ContentService.createTextOutput(cachedStr).setMimeType(ContentService.MimeType.JSON);
  var response = getReviewsData(subjectId, startTime);
  putLargeCache(cacheKey, response.getContent(), 600); // 10 นาที
  return response;
}

function putLargeCache(key, value, ttl, startTime) {
  if (!value) return;
  var cache = CacheService.getScriptCache();
  // Downsized to 25KB character slices to protect against multi-byte (Thai) UTF-8 expansion (up to 3x bytes per char)
  var chunkSize = 25 * 1024;
  var chunks = Math.ceil(value.length / chunkSize);

  try {
    cache.put(key + "_chunks", String(chunks), ttl);
    // putAll in batches of ≤100 keys — one RPC per batch instead of one per chunk
    var batch = {};
    var batchCount = 0;
    for (var i = 0; i < chunks; i++) {
      if (startTime) assertNotTimedOut_(startTime, 'putLargeCache:' + key);
      batch[key + "_chunk_" + i] = value.substring(i * chunkSize, (i + 1) * chunkSize);
      batchCount++;
      if (batchCount >= 100) {
        cache.putAll(batch, ttl);
        batch = {};
        batchCount = 0;
      }
    }
    if (batchCount > 0) cache.putAll(batch, ttl);
  } catch (e) {
    if (e.message && e.message.indexOf(DOGET_TIMEOUT_MARK) !== -1) throw e;
    console.warn("putLargeCache failed for key " + key + ": " + e.message);
  }
}

function getLargeCache(key) {
  var cache = CacheService.getScriptCache();
  var chunksStr = cache.get(key + "_chunks");
  if (!chunksStr) return null;
  
  var chunks = parseInt(chunksStr, 10);
  var keys = [];
  for (var i = 0; i < chunks; i++) keys.push(key + "_chunk_" + i);
  var chunkMap = cache.getAll(keys); // one RPC for all chunks instead of one per chunk
  var value = "";
  for (var j = 0; j < chunks; j++) {
    var chunk = chunkMap[key + "_chunk_" + j];
    if (chunk == null) return null; // If any chunk is lost, treat as cache miss
    value += chunk;
  }
  return value;
}

function getPriorYearAuditDataCached(startTime) {
  var v = getVersionCached();
  var cacheKey = "prior_year_audit_" + v;

  var cachedStr = getLargeCache(cacheKey);
  if (cachedStr != null) {
    return ContentService.createTextOutput(cachedStr).setMimeType(ContentService.MimeType.JSON);
  }

  if (startTime) assertNotTimedOut_(startTime, 'getPriorYearAuditDataCached:start');
  var response = getPriorYearAuditData(startTime);
  var responseStr = response.getContent();
  putLargeCache(cacheKey, responseStr, 1800); // 30 minutes — no startTime: data already built, write regardless of time
  return response;
}

function getStructureDataCached(filterSubject, startTime) {
  var v = getVersionCached();
  var cleanFilter = filterSubject ? String(filterSubject).trim().toUpperCase() : "all";
  var cacheKey = "struct_" + v + "_" + cleanFilter;
  
  var cachedStr = getLargeCache(cacheKey);
  if (cachedStr != null) {
    return ContentService.createTextOutput(cachedStr).setMimeType(ContentService.MimeType.JSON);
  }
  
  if (startTime) assertNotTimedOut_(startTime, 'getStructureDataCached:start');
  var response = getStructureData(filterSubject, startTime);
  var responseStr = response.getContent();
  putLargeCache(cacheKey, responseStr, 1800); // 30 minutes — no startTime: data already built, write regardless of time
  return response;
}

function getQuestionsDataCached(filterSubject, ss, startTime) {
  var v = getVersionCached();
  var cleanFilter = filterSubject ? String(filterSubject).trim().toUpperCase() : "all";
  var cacheKey = "questions_" + v + "_" + cleanFilter;

  var cachedStr = getLargeCache(cacheKey);
  if (cachedStr != null) {
    return ContentService.createTextOutput(cachedStr).setMimeType(ContentService.MimeType.JSON);
  }

  if (startTime) assertNotTimedOut_(startTime, 'getQuestionsDataCached:start');
  var response = getQuestionsData(filterSubject, ss, startTime);
  var responseStr = response.getContent();
  putLargeCache(cacheKey, responseStr, 1800); // 30 minutes — no startTime: data already built, write regardless of time
  return response;
}

function getAllDataForAdminCached(startTime) {
  startTime = startTime || Date.now();
  var v = getVersionCached();
  var cacheKey = "admin_all_data_" + v;

  var cachedStr = getLargeCache(cacheKey);
  if (cachedStr != null) {
    return ContentService.createTextOutput(cachedStr).setMimeType(ContentService.MimeType.JSON);
  }

  // Single-flight advisory flag: กันหลาย doGet execution ที่มาชนกันตอน cache miss (เช่นหลังแก้ข้อมูล v เปลี่ยน)
  // จาก compute payload 26MB ซ้ำกันคนละ execution พร้อมกัน (ตรงกับ log: 4-5 execution ค้าง 100-311s)
  // ไม่ใช่ LockService.getScriptLock() เพราะจะไปแย่ง lock กับ doPost admin tier (25s tryLock) จนเขียนข้อมูลค้าง
  var scriptCache = CacheService.getScriptCache();
  var inflightKey = "inflight_" + cacheKey;
  if (scriptCache.get(inflightKey) != null) {
    return ContentService.createTextOutput(JSON.stringify({
      result: 'error',
      message: 'Server is already building this data - please retry shortly or use Delta Sync'
    })).setMimeType(ContentService.MimeType.JSON);
  }

  try {
    scriptCache.put(inflightKey, "1", 150); // TTL กันธงค้างถ้า execution ตายกลางคัน (โควต้าจริง 360s แต่ guard ตัดที่ 90s)
    var response = getAllDataForAdmin(startTime);
    var responseStr = response.getContent();
    putLargeCache(cacheKey, responseStr, 1800); // 30 minutes — no startTime: data already built, write cache regardless of time elapsed
    return response;
  } finally {
    scriptCache.remove(inflightKey);
  }
}

// ────────────────────────────────────────────────────────────────────
// NEW: ADVANCED CACHED SHEET LOADERS (Bypasses Sheets API contention)
// ────────────────────────────────────────────────────────────────────

function getCategorySheetDataCached(ss, startTime) {
  var v = getVersionCached();
  var cacheKey = "category_sheet_raw_" + v;
  var cached = getLargeCache(cacheKey);
  if (cached) {
    return JSON.parse(cached);
  }
  if (startTime) assertNotTimedOut_(startTime, 'getCategorySheetDataCached');
  if (!ss) ss = SpreadsheetApp.openById(SHEET_ID);
  var catSheet = ss.getSheetByName('Category');
  if (!catSheet) return [];
  var rows = catSheet.getDataRange().getValues();
  putLargeCache(cacheKey, JSON.stringify(rows), 1800); // 30 minutes — no startTime on write-back
  return rows;
}

function getStructureSheetDataCached(ss, startTime) {
  var v = getVersionCached();
  var cacheKey = "structure_sheet_raw_" + v;
  var cached = getLargeCache(cacheKey);
  if (cached) {
    return JSON.parse(cached);
  }
  if (startTime) assertNotTimedOut_(startTime, 'getStructureSheetDataCached');
  if (!ss) ss = SpreadsheetApp.openById(SHEET_ID);
  var structSheet = ss.getSheetByName('Structure');
  if (!structSheet) return [];
  var rows = structSheet.getDataRange().getValues();
  putLargeCache(cacheKey, JSON.stringify(rows), 1800); // 30 minutes — no startTime on write-back
  return rows;
}

function getAllQuestionsCached(ss, startTime) {
  var v = getVersionCached();
  var cacheKey = "all_questions_raw_" + v;
  var cached = getLargeCache(cacheKey);
  if (cached) {
    return JSON.parse(cached);
  }
  if (startTime) assertNotTimedOut_(startTime, 'getAllQuestionsCached:before_sheet');
  if (!ss) ss = SpreadsheetApp.openById(SHEET_ID);
  var qSheet = ss.getSheetByName('Questions');
  var qLastRow = qSheet.getLastRow();
  if (qLastRow <= 1) return [];

  if (startTime) assertNotTimedOut_(startTime, 'getAllQuestionsCached:before_getValues');
  var qData = qSheet.getRange(2, 1, qLastRow - 1, 7).getValues();
  // No timeout check before putLargeCache — data already fetched; always write cache regardless of elapsed time
  putLargeCache(cacheKey, JSON.stringify(qData), 1800); // 30 minutes
  return qData;
}

function getCategoryToSubjectMapCached(ss, startTime) {
  var v = getVersionCached();
  var cacheKey = "cat_to_subj_map_" + v;
  var cached = getLargeCache(cacheKey);
  if (cached) {
    return JSON.parse(cached);
  }
  if (startTime) assertNotTimedOut_(startTime, 'getCategoryToSubjectMapCached');
  var catRows = getCategorySheetDataCached(ss, startTime);
  var categoryToSubjectMap = {};
  for (var i = 1; i < catRows.length; i++) {
    categoryToSubjectMap[String(catRows[i][0]).trim()] = String(catRows[i][1]).trim().toUpperCase();
  }
  putLargeCache(cacheKey, JSON.stringify(categoryToSubjectMap), 1800); // 30 minutes — no startTime on write-back
  return categoryToSubjectMap;
}
