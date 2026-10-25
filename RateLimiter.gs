/**
 * Fixed-window rate limiter backed by the per-user cache. Returns true
 * if the call is allowed, false if the caller is over the limit.
 */
function checkRateLimit_(key, maxCalls, windowSeconds) {
  const cache = CacheService.getUserCache();
  const countStr = cache.get(key);
  const count = countStr ? parseInt(countStr, 10) : 0;
  if (count >= maxCalls) return false;
  cache.put(key, String(count + 1), windowSeconds);
  return true;
}
