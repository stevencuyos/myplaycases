/**
 * Shared-data cache (10-minute TTL). Chunked because CacheService caps
 * individual values at ~100KB.
 */
const CACHE_TTL_SECONDS = 600;
const CACHE_CHUNK_SIZE = 90000;
const CACHE_KEY = 'mpc_shared_data';

function writeToCache_(cache, key, obj) {
  try {
    const json = JSON.stringify(obj);
    const chunks = [];
    for (let i = 0; i < json.length; i += CACHE_CHUNK_SIZE) chunks.push(json.slice(i, i + CACHE_CHUNK_SIZE));
    const entries = {};
    entries[key + '_meta'] = String(chunks.length);
    chunks.forEach(function (chunk, idx) { entries[key + '_' + idx] = chunk; });
    cache.putAll(entries, CACHE_TTL_SECONDS);
  } catch (e) {
    Logger.log('Cache write skipped: %s', e.message);
  }
}

function readFromCache_(cache, key) {
  try {
    const countStr = cache.get(key + '_meta');
    if (!countStr) return null;
    const count = parseInt(countStr, 10);
    const keys = [];
    for (let i = 0; i < count; i++) keys.push(key + '_' + i);
    const parts = cache.getAll(keys);
    let json = '';
    for (let i = 0; i < count; i++) {
      const part = parts[key + '_' + i];
      if (!part) return null;
      json += part;
    }
    return JSON.parse(json);
  } catch (e) {
    return null;
  }
}

/**
 * Returns cached shared data, recomputing under a lock if missing or
 * bypassed. The lock stops a cache-expiry "stampede" where several
 * concurrent users all recompute at once.
 */
function getSharedData_(bypassCache) {
  const cache = CacheService.getScriptCache();
  if (!bypassCache) {
    const cached = readFromCache_(cache, CACHE_KEY);
    if (cached) return cached;
  }

  const lock = LockService.getScriptLock();
  const gotLock = lock.tryLock(15000);
  try {
    if (!bypassCache) {
      const cachedAfterWait = readFromCache_(cache, CACHE_KEY);
      if (cachedAfterWait) return cachedAfterWait;
    }
    const fresh = computeSharedData_();
    writeToCache_(cache, CACHE_KEY, fresh);
    return fresh;
  } finally {
    if (gotLock) lock.releaseLock();
  }
}
