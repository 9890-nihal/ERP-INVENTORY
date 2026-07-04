/**
 * CacheHelper.gs
 * Centralized Smart Caching System for ERP
 * TTL = 120 seconds (2 minutes) by default
 * 
 * Usage:
 *   const data = getCachedData('material_list', loadMaterialFromSheet, 120);
 *   invalidateCache('material_list');
 */

// ====================== CONFIG ======================
const CACHE_TTL_SECONDS = 120;           // Default 2 minutes
const CACHE_PREFIX = 'ERP_';             // Namespace to avoid conflicts

// ====================== CORE FUNCTIONS ======================

/**
 * Generic cached data getter
 * @param {string} key - Unique cache key
 * @param {function} loaderFunction - Function that loads fresh data if cache miss
 * @param {number} [ttl] - Time to live in seconds (optional)
 */
function getCachedData(key, loaderFunction, ttl) {
  const cache = CacheService.getScriptCache();
  const cacheKey = CACHE_PREFIX + key;
  const cached = cache.get(cacheKey);

  if (cached) {
    try {
      return JSON.parse(cached);
    } catch (e) {
      // Corrupted cache, fall through to reload
    }
  }

  // Cache miss or expired → load fresh data
  const freshData = loaderFunction();
  const cacheTTL = ttl || CACHE_TTL_SECONDS;

  try {
    cache.put(cacheKey, JSON.stringify(freshData), cacheTTL);
  } catch (e) {
    // Data too large for cache, skip caching silently
    console.warn('Cache put failed for key: ' + key);
  }

  return freshData;
}

/**
 * Invalidate (clear) a specific cache key
 */
function invalidateCache(key) {
  const cache = CacheService.getScriptCache();
  cache.remove(CACHE_PREFIX + key);
}

/**
 * Clear ALL ERP caches (use carefully)
 */
function clearAllErpCache() {
  const cache = CacheService.getScriptCache();
  // We can't easily clear all keys, so we use a version bump technique
  const versionKey = CACHE_PREFIX + 'CACHE_VERSION';
  const currentVersion = cache.get(versionKey) || '1';
  cache.put(versionKey, String(Number(currentVersion) + 1), 86400); // 24 hours
  console.log('ERP Cache cleared (version bumped)');
}

/**
 * Check if cache is enabled (useful for debugging)
 */
function isCacheEnabled() {
  return true; // Change to false temporarily during heavy debugging
}

// ====================== SPECIFIC CACHED LOADERS ======================

/**
 * Cached Material List (most used master data)
 */
function getCachedMaterialList_() {
  if (!isCacheEnabled()) {
    return getMaterialRows_(); // fallback to direct
  }
  return getCachedData('material_list', getMaterialRows_, 300); // 5 minutes for master data
}

/**
 * Cached Recent Inward Entries
 */
function getCachedInwardRecentEntries_() {
  if (!isCacheEnabled()) {
    return getInwardRecentEntries_();
  }
  return getCachedData('inward_recent', getInwardRecentEntries_, 60); // 1 minute
}

/**
 * Cached Inward Suggestions (vendors + locations)
 */
function getCachedInwardSuggestions_() {
  if (!isCacheEnabled()) {
    return getInwardSuggestions_();
  }
  return getCachedData('inward_suggestions', getInwardSuggestions_, 180);
}

/**
 * Cached Inventory Data for a specific period
 */
function getCachedInventoryData_(period) {
  const cacheKey = 'inventory_' + (period || 'current');
  if (!isCacheEnabled()) {
    return getInventoryInitialData({ period: period });
  }
  return getCachedData(cacheKey, function() {
    return getInventoryInitialData({ period: period });
  }, 120);
}

// ====================== INVALIDATION HELPERS ======================

/**
 * Call this after saving/updating Material List
 */
function invalidateMaterialCache() {
  invalidateCache('material_list');
}

/**
 * Call this after saving new Inward entry
 */
function invalidateInwardCache() {
  invalidateCache('inward_recent');
  invalidateCache('inward_suggestions');
}

/**
 * Call this after Inventory location update or period change
 */
function invalidateInventoryCache(period) {
  invalidateCache('inventory_' + (period || 'current'));
}
