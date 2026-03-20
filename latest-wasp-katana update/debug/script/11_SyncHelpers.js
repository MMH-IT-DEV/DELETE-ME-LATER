// ============================================
// 11_SyncHelpers.gs - INVENTORY SYNC HELPERS
// ============================================
// Core sync logic: map building, delta calculation,
// and adjustment application between Katana and WASP.
// ============================================
// DEPENDENCIES:
//   fetchAllKatanaInventory()  — 98_ItemComparison.gs
//   fetchAllKatanaLocations()  — 98_ItemComparison.gs
//   fetchAllKatanaVariants()   — 98_ItemComparison.gs
//   fetchAllKatanaProducts()   — 98_ItemComparison.gs
//   waspApiCall(url, payload)  — 05_WaspAPI.gs
//   isSkippedSku(sku)          — 00_Config.gs
//   CONFIG                     — 00_Config.gs
//   SYNC_CONFIG                — 00_Config.gs
//   SYNC_LOCATION_MAP          — 00_Config.gs
// ============================================

// ============================================
// MAP BUILDERS
// ============================================

/**
 * Build a Katana inventory map keyed by 'SKU|site|location'.
 *
 * Fetches all Katana inventory records, resolves each to a SKU,
 * Katana location name, and item type, then looks up the
 * corresponding WASP site + location via SYNC_LOCATION_MAP.
 * Quantities are accumulated when multiple Katana variants
 * resolve to the same composite key.
 *
 * @return {Object} Map of { 'SKU|site|location': qty }
 */
function syncBuildKatanaMap() {
  Logger.log('--- syncBuildKatanaMap: fetching Katana data ---');

  var inventory = fetchAllKatanaInventory();
  var locations = fetchAllKatanaLocations();
  var variants  = fetchAllKatanaVariants();
  var products  = fetchAllKatanaProducts();

  // Build variant lookup: variantId → { sku, productId }
  var variantMap = {};
  for (var v = 0; v < variants.length; v++) {
    var vr = variants[v];
    variantMap[vr.id] = {
      sku: vr.sku || '',
      productId: vr.product_id || ''
    };
  }

  // Build product lookup: productId → { type }
  var productMap = {};
  for (var p = 0; p < products.length; p++) {
    var prod = products[p];
    productMap[prod.id] = {
      type: prod.type || prod.product_type || prod.category || 'unknown'
    };
  }

  var map = {};
  var totalQty = 0;
  var skipped = 0;

  for (var i = 0; i < inventory.length; i++) {
    var record = inventory[i];
    var qty = record.quantity_in_stock || 0;

    // Resolve variant
    var vInfo = variantMap[record.variant_id];
    if (!vInfo || !vInfo.sku) {
      skipped++;
      continue;
    }

    var sku = vInfo.sku;

    // Skip service / virtual SKUs (e.g. OP- prefix)
    if (isSkippedSku(sku)) {
      skipped++;
      continue;
    }

    // Resolve Katana location name
    var katanaLocName = locations[record.location_id];
    if (!katanaLocName) {
      skipped++;
      continue;
    }

    // Amazon USA stock is at Amazon's fulfillment center — not our facility
    if (katanaLocName === 'Amazon USA') {
      skipped++;
      continue;
    }

    // Resolve item type
    var pInfo = productMap[vInfo.productId] || { type: 'unknown' };
    var itemType = pInfo.type;

    // Look up WASP destination via SYNC_LOCATION_MAP
    var lookupKey = katanaLocName + '|' + itemType;
    var dest = SYNC_LOCATION_MAP[lookupKey];
    if (!dest) {
      Logger.log('  WARN: No SYNC_LOCATION_MAP entry for key "' + lookupKey + '" (SKU=' + sku + ') — skipping');
      skipped++;
      continue;
    }

    var waspSite     = dest.site;
    var waspLocation = dest.location;

    // Accumulate quantity under composite key
    var mapKey = sku + '|' + waspSite + '|' + waspLocation;
    if (map[mapKey]) {
      map[mapKey] += qty;
    } else {
      map[mapKey] = qty;
    }
    totalQty += qty;
  }

  var totalKeys = Object.keys(map).length;
  Logger.log('syncBuildKatanaMap: keys=' + totalKeys + ', totalQty=' + totalQty + ', skipped=' + skipped);
  return map;
}

/**
 * Build a WASP inventory map keyed by 'SKU|site|location'.
 *
 * First attempts the bulk advancedinventorysearch endpoint.
 * If that returns no data, falls back to per-SKU inventorysearch
 * calls against each site configured in SYNC_LOCATION_MAP.
 *
 * @param {Array} katanaSkus  Array of unique SKU strings from Katana
 * @return {Object}           Map of { 'SKU|site|location': qty }
 */
function syncBuildWaspMap(katanaSkus) {
  Logger.log('--- syncBuildWaspMap: attempting advancedinventorysearch ---');

  var map = {};
  var url = CONFIG.WASP_BASE_URL + '/public-api/ic/item/advancedinventorysearch';
  var payload = [{}];

  var result = waspApiCall(url, payload);

  if (result.success) {
    try {
      var parsed = JSON.parse(result.response);
      var data = parsed.Data || parsed.data || [];

      if (Array.isArray(data) && data.length > 0) {
        for (var i = 0; i < data.length; i++) {
          var rec = data[i];
          var key = rec.ItemNumber + '|' + rec.SiteName + '|' + rec.LocationCode;
          var qty = rec.QuantityOnHand || rec.Quantity || 0;
          if (map[key]) {
            map[key] += qty;
          } else {
            map[key] = qty;
          }
        }
        Logger.log('syncBuildWaspMap: advancedinventorysearch returned ' + data.length + ' records');
        return map;
      }
    } catch (e) {
      Logger.log('syncBuildWaspMap: advancedinventorysearch parse error — ' + e.message);
    }
  }

  // ---- Fallback: per-SKU inventorysearch ----
  Logger.log('syncBuildWaspMap: Falling back to per-item inventorysearch');

  var searchUrl = CONFIG.WASP_BASE_URL + '/public-api/ic/item/inventorysearch';

  // Collect the unique site names we need to query from SYNC_LOCATION_MAP
  var sitesNeeded = {};
  for (var mapKey in SYNC_LOCATION_MAP) {
    if (SYNC_LOCATION_MAP.hasOwnProperty(mapKey)) {
      var entry = SYNC_LOCATION_MAP[mapKey];
      sitesNeeded[entry.site] = true;
    }
  }
  var sites = Object.keys(sitesNeeded);

  for (var s = 0; s < katanaSkus.length; s++) {
    var sku = katanaSkus[s];

    for (var siteIdx = 0; siteIdx < sites.length; siteIdx++) {
      var siteName = sites[siteIdx];
      var searchPayload = { SearchPattern: sku, SiteName: siteName };
      var searchResult = waspApiCall(searchUrl, searchPayload);

      if (searchResult.success) {
        try {
          var searchParsed = JSON.parse(searchResult.response);
          var searchData = searchParsed.Data || searchParsed.data || [];
          for (var d = 0; d < searchData.length; d++) {
            var item = searchData[d];
            var itemKey = item.ItemNumber + '|' + item.SiteName + '|' + item.LocationCode;
            var itemQty = item.QuantityOnHand || item.Quantity || 0;
            if (map[itemKey]) {
              map[itemKey] += itemQty;
            } else {
              map[itemKey] = itemQty;
            }
          }
        } catch (parseErr) {
          Logger.log('  WARN: parse error for SKU=' + sku + ' site=' + siteName + ' — ' + parseErr.message);
        }
      }

      Utilities.sleep(SYNC_CONFIG.RATE_LIMIT_MS);
    }

    // Progress logging every 50 SKUs
    if ((s + 1) % 50 === 0) {
      Logger.log('  syncBuildWaspMap fallback: processed ' + (s + 1) + ' / ' + katanaSkus.length + ' SKUs');
    }
  }

  Logger.log('syncBuildWaspMap: fallback complete, total keys=' + Object.keys(map).length);
  return map;
}

// ============================================
// DELTA CALCULATION
// ============================================

/**
 * Calculate inventory deltas between Katana (source of truth) and WASP.
 *
 * Builds the union of all keys from both maps and computes
 * delta = katanaQty - waspQty for each key.
 * Only keys where delta !== 0 are included in the result.
 * Results are sorted by |delta| descending (largest discrepancies first).
 *
 * @param {Object} katanaMap  { 'SKU|site|location': qty }
 * @param {Object} waspMap    { 'SKU|site|location': qty }
 * @return {Array} Array of { sku, site, location, katanaQty, waspQty, delta }
 */
function syncCalcDeltas(katanaMap, waspMap) {
  Logger.log('--- syncCalcDeltas: computing deltas ---');

  // Build union of all keys
  var allKeys = {};
  var key;

  for (key in katanaMap) {
    if (katanaMap.hasOwnProperty(key)) {
      allKeys[key] = true;
    }
  }
  for (key in waspMap) {
    if (waspMap.hasOwnProperty(key)) {
      allKeys[key] = true;
    }
  }

  var results = [];
  var totalAdd = 0;
  var totalRemove = 0;

  for (key in allKeys) {
    if (!allKeys.hasOwnProperty(key)) continue;

    var katanaQty = katanaMap[key] || 0;
    var waspQty   = waspMap[key]   || 0;
    var delta     = katanaQty - waspQty;

    if (delta === 0) continue;

    // Split composite key: 'SKU|site|location'
    var parts    = key.split('|');
    var sku      = parts[0];
    var site     = parts[1];
    var location = parts[2];

    results.push({
      sku: sku,
      site: site,
      location: location,
      katanaQty: katanaQty,
      waspQty: waspQty,
      delta: delta
    });

    if (delta > 0) {
      totalAdd += delta;
    } else {
      totalRemove += Math.abs(delta);
    }
  }

  // Sort by absolute delta descending (biggest discrepancies first)
  results.sort(function(a, b) {
    return Math.abs(b.delta) - Math.abs(a.delta);
  });

  Logger.log('syncCalcDeltas: Total deltas=' + results.length +
    ', Total add=' + totalAdd +
    ', Total remove=' + totalRemove);

  return results;
}

// ============================================
// ADJUSTMENT APPLICATION
// ============================================

/**
 * Apply inventory adjustments to WASP based on calculated deltas.
 *
 * Iterates through deltas up to SYNC_CONFIG.MAX_EXECUTE_ITEMS.
 * Each adjustment is posted to the WASP /adjust endpoint via
 * syncWaspCallWithRetry_ for transient-error resilience.
 *
 * When SYNC_CONFIG.DRY_RUN is true the payloads are logged but
 * no API calls are made (callers should check DRY_RUN before
 * calling this function, but the guard is also applied here).
 *
 * @param {Array} deltas  Output of syncCalcDeltas()
 * @return {Array}        Array of { sku, site, location, delta, status, error }
 */
function syncApplyAdjustments(deltas) {
  Logger.log('--- syncApplyAdjustments: applying ' + deltas.length + ' deltas ---');

  var url     = CONFIG.WASP_BASE_URL + '/public-api/transactions/item/adjust';
  var today   = new Date().toISOString().split('T')[0];
  var results = [];
  var applied = 0;
  var okCount = 0;
  var errCount = 0;

  var limit = Math.min(deltas.length, SYNC_CONFIG.MAX_EXECUTE_ITEMS);

  for (var i = 0; i < limit; i++) {
    var delta = deltas[i];

    var payload = [{
      ItemNumber:   delta.sku,
      Quantity:     delta.delta,
      SiteName:     delta.site,
      LocationCode: delta.location,
      Notes:        'Katana sync ' + today
    }];

    if (SYNC_CONFIG.DRY_RUN) {
      Logger.log('  [DRY RUN] ' + delta.sku + ' delta=' + delta.delta +
        ' @ ' + delta.site + '/' + delta.location);
      results.push({
        sku:      delta.sku,
        site:     delta.site,
        location: delta.location,
        delta:    delta.delta,
        status:   'DRY_RUN',
        error:    ''
      });
      applied++;
      continue;
    }

    var callResult = syncWaspCallWithRetry_(url, payload, 2);

    var status;
    var error;
    if (callResult.success) {
      status = 'OK';
      error  = '';
      okCount++;
    } else {
      status = 'ERROR';
      error  = (callResult.response || '').substring(0, 200);
      errCount++;
    }

    results.push({
      sku:      delta.sku,
      site:     delta.site,
      location: delta.location,
      delta:    delta.delta,
      status:   status,
      error:    error
    });

    applied++;

    Utilities.sleep(SYNC_CONFIG.RATE_LIMIT_MS);

    // Progress log every 25 items
    if (applied % 25 === 0) {
      Logger.log('  syncApplyAdjustments: ' + applied + ' / ' + limit +
        ' applied (OK=' + okCount + ', ERROR=' + errCount + ')');
    }
  }

  Logger.log('syncApplyAdjustments: Applied ' + applied +
    ' adjustments, OK=' + okCount + ', ERROR=' + errCount);

  return results;
}

// ============================================
// PRIVATE HELPERS
// ============================================

/**
 * Call waspApiCall with exponential-backoff retry for transient errors.
 *
 * Retries on HTTP 429 (rate limit) and 5xx (server error) responses.
 * 4xx errors other than 429 are treated as permanent and returned
 * immediately without further retries.
 *
 * Back-off schedule (seconds): 3^attempt  →  3s, 9s, 27s …
 *
 * @param {string} url        WASP API endpoint URL
 * @param {Array}  payload    Request body (will be JSON-encoded)
 * @param {number} maxRetries Maximum number of retry attempts
 * @return {Object}           Last { success, code, response } from waspApiCall
 */
function syncWaspCallWithRetry_(url, payload, maxRetries) {
  var attempt = 0;
  var result;

  while (true) {
    result = waspApiCall(url, payload);

    if (result.success) {
      return result;
    }

    var code = result.code || 0;

    // Transient: rate-limited or server error
    if (code === 429 || code >= 500) {
      attempt++;
      if (attempt > maxRetries) {
        Logger.log('  syncWaspCallWithRetry_: giving up after ' + attempt + ' attempts (code=' + code + ')');
        return result;
      }
      var sleepMs = Math.pow(3, attempt) * 1000;
      Logger.log('  syncWaspCallWithRetry_: transient error code=' + code +
        ', retry ' + attempt + '/' + maxRetries +
        ' after ' + sleepMs + 'ms');
      Utilities.sleep(sleepMs);
      continue;
    }

    // Permanent error (4xx excluding 429) — do not retry
    Logger.log('  syncWaspCallWithRetry_: permanent error code=' + code + ', aborting retries');
    return result;
  }
}
