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

function putLargeCache(key, value, ttl) {
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

function getStructureDataCached(filterSubject) {
  var v = getVersionCached();
  var cleanFilter = filterSubject ? String(filterSubject).trim().toUpperCase() : "all";
  var cacheKey = "struct_" + v + "_" + cleanFilter;
  
  var cachedStr = getLargeCache(cacheKey);
  if (cachedStr != null) {
    return ContentService.createTextOutput(cachedStr).setMimeType(ContentService.MimeType.JSON);
  }
  
  var response = getStructureData(filterSubject);
  var responseStr = response.getContent();
  putLargeCache(cacheKey, responseStr, 1800); // 30 minutes
  return response;
}

function getQuestionsDataCached(filterSubject, ss) {
  var v = getVersionCached();
  var cleanFilter = filterSubject ? String(filterSubject).trim().toUpperCase() : "all";
  var cacheKey = "questions_" + v + "_" + cleanFilter;

  var cachedStr = getLargeCache(cacheKey);
  if (cachedStr != null) {
    return ContentService.createTextOutput(cachedStr).setMimeType(ContentService.MimeType.JSON);
  }

  var response = getQuestionsData(filterSubject, ss);
  var responseStr = response.getContent();
  putLargeCache(cacheKey, responseStr, 1800); // ขยาย Cache เป็น 30 นาที เนื่องจากคีย์ผูกกับเวอร์ชันอยู่แล้ว (เมื่อมีข้อมูลใหม่แคชจะรีเซ็ตอัตโนมัติ)
  return response;
}

function getAllDataForAdminCached() {
  var v = getVersionCached();
  var cacheKey = "admin_all_data_" + v;

  var cachedStr = getLargeCache(cacheKey);
  if (cachedStr != null) {
    return ContentService.createTextOutput(cachedStr).setMimeType(ContentService.MimeType.JSON);
  }

  var response = getAllDataForAdmin();
  var responseStr = response.getContent();
  putLargeCache(cacheKey, responseStr, 1800); // 30 minutes — key is version-scoped (admin_all_data_<v>) so staleness impossible
  return response;
}

// ────────────────────────────────────────────────────────────────────
// NEW: ADVANCED CACHED SHEET LOADERS (Bypasses Sheets API contention)
// ────────────────────────────────────────────────────────────────────

function getCategorySheetDataCached(ss) {
  var v = getVersionCached();
  var cacheKey = "category_sheet_raw_" + v;
  var cached = getLargeCache(cacheKey);
  if (cached) {
    return JSON.parse(cached);
  }
  if (!ss) ss = SpreadsheetApp.openById(SHEET_ID);
  var catSheet = ss.getSheetByName('Category');
  if (!catSheet) return [];
  var rows = catSheet.getDataRange().getValues();
  putLargeCache(cacheKey, JSON.stringify(rows), 1800); // 30 minutes
  return rows;
}

function getStructureSheetDataCached(ss) {
  var v = getVersionCached();
  var cacheKey = "structure_sheet_raw_" + v;
  var cached = getLargeCache(cacheKey);
  if (cached) {
    return JSON.parse(cached);
  }
  if (!ss) ss = SpreadsheetApp.openById(SHEET_ID);
  var structSheet = ss.getSheetByName('Structure');
  if (!structSheet) return [];
  var rows = structSheet.getDataRange().getValues();
  putLargeCache(cacheKey, JSON.stringify(rows), 1800); // 30 minutes
  return rows;
}

function getAllQuestionsCached(ss) {
  var v = getVersionCached();
  var cacheKey = "all_questions_raw_" + v;
  var cached = getLargeCache(cacheKey);
  if (cached) {
    return JSON.parse(cached);
  }
  if (!ss) ss = SpreadsheetApp.openById(SHEET_ID);
  var qSheet = ss.getSheetByName('Questions');
  var qLastRow = qSheet.getLastRow();
  if (qLastRow <= 1) return [];

  var qData = qSheet.getRange(2, 1, qLastRow - 1, 7).getValues();
  putLargeCache(cacheKey, JSON.stringify(qData), 1800); // 30 minutes
  return qData;
}

function getCategoryToSubjectMapCached(ss) {
  var v = getVersionCached();
  var cacheKey = "cat_to_subj_map_" + v;
  var cached = getLargeCache(cacheKey);
  if (cached) {
    return JSON.parse(cached);
  }
  var catRows = getCategorySheetDataCached(ss);
  var categoryToSubjectMap = {};
  for (var i = 1; i < catRows.length; i++) {
    categoryToSubjectMap[String(catRows[i][0]).trim()] = String(catRows[i][1]).trim().toUpperCase();
  }
  putLargeCache(cacheKey, JSON.stringify(categoryToSubjectMap), 1800); // 30 minutes
  return categoryToSubjectMap;
}
