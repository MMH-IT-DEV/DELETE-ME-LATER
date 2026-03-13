// ============================================
// 07_F3_AutoPoll.gs - F3 STOCK TRANSFER AUTO-POLLING
// ============================================
// FIXED: Changed const to var for Google Apps Script compatibility
// FIXED: Removed verbose logging - only errors logged now
// ============================================

var F3_CONFIG = {
  POLL_INTERVAL_MINUTES: 5,
  SHEET_ID: CONFIG.DEBUG_SHEET_ID,
  TAB_NAME: 'StockTransfers',
  KATANA_ENDPOINT: 'stock_transfers',
  CACHE_KEY_LAST_SYNC: 'F3_LAST_SYNC_TIMESTAMP',

  // Sync when ST is in these statuses (partial or full receive)
  SYNC_ON_STATUS: ['completed', 'done', 'received', 'partial', 'in_transit'],

  // Only sync STs created AFTER this date (set to now to ignore old ones)
  MIN_CREATED_DATE: '2026-02-04T00:00:00Z',

  // Only sync STs created within the last X days
  MAX_AGE_DAYS: 30,

  COLS: {
    ST_ID: 1, ST_NUMBER: 2, STATUS: 3, FROM_LOC: 4, TO_LOC: 5,
    ITEM_COUNT: 6, CREATED: 7, UPDATED: 8, WASP_STATUS: 9,
    WASP_SYNCED: 10, NOTES: 11
  },

  ITEM_COLS: {
    ST_ID: 1, ST_NUMBER: 2, SKU: 3, QTY: 4, LOT: 5,
    EXPIRY: 6, FROM_LOC: 7, TO_LOC: 8, SYNC_STATUS: 9, NOTES: 10
  }
};

/**
 * SITE MAPPING: Katana Location → WASP Site + Location
 */
var F3_SITE_MAP = {
  'mmh kelowna': {
    site: 'MMH Kelowna',
    location: 'RECEIVING-DOCK'
  },
  'mmh mayfair': {
    site: 'MMH Mayfair',
    location: 'QA-Hold-1'
  },
  'storage warehouse': {
    site: 'Storage Warehouse',
    location: 'SW-STORAGE'
  }
};

/**
 * Virtual/external locations that have no WASP equivalent.
 * Transfers involving these locations are skipped immediately.
 */
var F3_SKIP_LOCATIONS = ['amazon usa', 'shopify'];

// ============================================
// ITEM TRACKING - REMOVED
// ============================================
// STItems tracking removed - ST-level status only

// ============================================
// ERROR CLASSIFICATION — PERMANENT vs TRANSIENT
// ============================================

/**
 * Permanent WASP errors — will never self-resolve.
 * These fail immediately on first attempt (no retries).
 */
var PERMANENT_ERROR_PATTERNS = [
  'does not exist',          // Location/site not found in WASP
  'not found',               // Item/lot not found
  'Remove failed',           // Insufficient qty at source (data issue)
  'Insufficient',            // Explicit insufficient
  'not enough',              // Variant of insufficient
  'Invalid SKU',             // Bad data
  'Unknown source',          // Unmapped Katana location
  'Unknown dest'             // Unmapped Katana location
];

/**
 * Check if an error message indicates a permanent (non-retryable) failure.
 */
function isPermanentError(errorMsg) {
  if (!errorMsg) return false;
  var lower = String(errorMsg).toLowerCase();
  for (var i = 0; i < PERMANENT_ERROR_PATTERNS.length; i++) {
    if (lower.indexOf(PERMANENT_ERROR_PATTERNS[i].toLowerCase()) >= 0) return true;
  }
  return false;
}

/**
 * Parse retry count from Notes field. Format: "[R2] error details..."
 * Returns 0 if no retry marker found.
 */
function parseRetryCount(notes) {
  if (!notes) return 0;
  var match = String(notes).match(/^\[R(\d+)\]\s*/);
  return match ? parseInt(match[1], 10) : 0;
}

/**
 * Prepend retry count to notes. "[R2] error details..."
 */
function prependRetryCount(count, notes) {
  // Strip existing [Rx] prefix if present
  var clean = String(notes || '').replace(/^\[R\d+\]\s*/, '');
  return '[R' + count + '] ' + clean;
}

// Max transient retries before escalating to permanent "Failed"
var MAX_TRANSIENT_RETRIES = 3;

// Backoff schedule (hours): 1h, 6h, 24h
var RETRY_BACKOFF_HOURS = [1, 6, 24];

// Katana ST statuses that mean the transfer was cancelled — trigger WASP reversal
var F3_REVERSE_STATUS = ['cancelled', 'voided'];

// ============================================
// ERROR MESSAGE HELPERS
// ============================================

function parseWaspError(response, operation, sku, location) {
  try {
    var parsed = typeof response === 'string' ? JSON.parse(response) : response;

    if (parsed.Data && parsed.Data.ResultList && parsed.Data.ResultList.length > 0) {
      var result = parsed.Data.ResultList[0];
      var msg = result.Message || '';
      var msgLower = msg.toLowerCase();

      if (msgLower.indexOf('date code is missing') >= 0) {
        return 'Lot required for ' + (sku || 'item');
      }
      if (msgLower.indexOf('insufficient') >= 0 || msgLower.indexOf('not enough') >= 0) {
        return 'No stock at ' + (location || 'source');
      }
      if (msgLower.indexOf('does not exist') >= 0) {
        return (location || 'Location') + ' not in WASP';
      }
      if (msgLower.indexOf('item') >= 0 && msgLower.indexOf('not found') >= 0) {
        return (sku || 'Item') + ' not in WASP';
      }
      if (msgLower.indexOf('lot') >= 0 && msgLower.indexOf('not found') >= 0) {
        return 'Lot not found at ' + (location || 'location');
      }

      return msg.replace(/ItemNumber:\s*\S+\s*is\s*fail,?\s*message\s*:/gi, '').trim().substring(0, 60);
    }

    if (parsed.Message) return parsed.Message.substring(0, 60);
    return operation + ' failed';
  } catch (e) {
    return operation + ' failed';
  }
}

/**
 * Format item details for notes - CLEAN FORMAT
 */
function formatItemSummary(items) {
  if (!items || items.length === 0) return '';

  var parts = [];
  for (var i = 0; i < items.length; i++) {
    var item = items[i];
    var part = item.qty + 'x ' + item.sku;
    if (item.lot) part += ' [' + item.lot + ']';
    parts.push(part);
  }
  return parts.join(', ');
}

// ============================================
// WASP API FUNCTIONS
// ============================================

function waspAddInventoryToSite(siteName, itemNumber, quantity, locationCode, lotNumber, expiryDate, notes) {
  var url = CONFIG.WASP_BASE_URL + '/public-api/transactions/item/add';
  var normalizedExpiry = normalizeBusinessDate_(expiryDate);

  var item = {
    ItemNumber: itemNumber,
    Quantity: quantity,
    SiteName: siteName,
    LocationCode: locationCode,
    Notes: notes || ''
  };

  // Only include Lot/DateCode when we have actual values — empty fields cause WASP errors
  if (lotNumber) {
    item.Lot = lotNumber;
    item.DateCode = normalizedExpiry || '';
  }

  return waspApiCall(url, [item]);
}

function waspRemoveInventoryFromSite(siteName, itemNumber, quantity, locationCode, lotNumber, dateCode, notes) {
  var url = CONFIG.WASP_BASE_URL + '/public-api/transactions/item/remove';
  var normalizedDateCode = normalizeBusinessDate_(dateCode);

  var item = {
    ItemNumber: itemNumber,
    Quantity: quantity,
    SiteName: siteName,
    LocationCode: locationCode,
    Notes: notes || ''
  };

  // Only include Lot/DateCode when we have actual values — empty fields cause WASP errors
  if (lotNumber) {
    item.Lot = lotNumber;
    item.DateCode = normalizedDateCode || '';
  }

  return waspApiCall(url, [item]);
}

// ============================================
// LOCATION/SITE MAPPING
// ============================================

function mapKatanaToWaspSite(katanaLocation) {
  if (!katanaLocation) return null;

  var lower = katanaLocation.toLowerCase().trim();
  var useConfigDrivenMap = !!(typeof HOTFIX_FLAGS !== 'undefined' && HOTFIX_FLAGS.F3_USE_CONFIG_DRIVEN_LOCATION_MAP);

  if (useConfigDrivenMap) {
    var configuredMap = {};
    var configSources = [typeof F3_LOCATION_OVERRIDES !== 'undefined' ? F3_LOCATION_OVERRIDES : null, KATANA_LOCATION_TO_WASP];
    for (var srcIdx = 0; srcIdx < configSources.length; srcIdx++) {
      var srcMap = configSources[srcIdx];
      if (!srcMap) continue;
      for (var srcKey in srcMap) {
        if (!srcMap.hasOwnProperty(srcKey)) continue;
        configuredMap[String(srcKey || '').toLowerCase().trim()] = srcMap[srcKey];
      }
    }

    if (configuredMap[lower]) return configuredMap[lower];
    for (var cfgKey in configuredMap) {
      if (lower.indexOf(cfgKey) >= 0 || cfgKey.indexOf(lower) >= 0) {
        return configuredMap[cfgKey];
      }
    }
  }

  if (F3_SITE_MAP[lower]) return F3_SITE_MAP[lower];

  for (var key in F3_SITE_MAP) {
    if (lower.indexOf(key) >= 0 || key.indexOf(lower) >= 0) {
      return F3_SITE_MAP[key];
    }
  }

  return null;
}

function getF3LocationRank_(siteName, locationCode, preferredLocation) {
  var site = String(siteName || '').trim();
  var loc = String(locationCode || '').trim();
  var preferred = String(preferredLocation || '').trim();

  if (preferred && loc === preferred) return -1000;

  var priorities = {
    'MMH Kelowna': ['SHIPPING-DOCK', 'PRODUCTION', 'RECEIVING-DOCK', 'PROD-RECEIVING', 'UNSORTED', 'SHOPIFY'],
    'MMH Mayfair': ['QA-Hold-1', 'QA-Hold-2', 'QA-Hold-4', 'QA-Hold-5'],
    'Storage Warehouse': ['SW-STORAGE']
  };

  var siteList = priorities[site] || [];
  for (var i = 0; i < siteList.length; i++) {
    if (siteList[i] === loc) return i;
  }

  return 999;
}

function fetchF3SourceCandidates_(itemNumber, siteName, lotNumber, expectedDateCode) {
  if (!itemNumber || !siteName) return [];

  var url = CONFIG.WASP_BASE_URL + '/public-api/ic/item/advancedinventorysearch';
  var maxPages = 20;
  var expectedDateSet = buildComparableDateSet_(expectedDateCode);
  var candidates = [];

  try {
    for (var page = 1; page <= maxPages; page++) {
      var payload = {
        SearchPattern: itemNumber,
        PageSize: 100,
        PageNumber: page
      };

      var result = waspApiCall(url, payload);
      if (!result.success) return [];

      var response = JSON.parse(result.response);
      var rows = getWaspResultRows_(response);
      if (rows.length === 0) break;

      for (var i = 0; i < rows.length; i++) {
        var row = rows[i];
        if (!row) continue;

        var rowItem = String(row.ItemNumber || row.itemNumber || '').trim();
        if (rowItem !== itemNumber) continue;

        var rowSite = String(row.SiteName || row.siteName || '').trim();
        if (rowSite !== siteName) continue;

        var rowLoc = String(row.LocationCode || row.locationCode || '').trim();
        if (!rowLoc) continue;

        var rowLot = String(row.Lot || row.lot || row.LotNumber || row.lotNumber || '').trim();
        if (lotNumber && rowLot !== String(lotNumber).trim()) continue;

        var rowDateRaw = row.DateCode || row.dateCode || '';
        var rowDateCode = normalizeExactLotDate_(rowDateRaw);
        var rowDateSet = buildComparableDateSet_(rowDateRaw);
        if (expectedDateSet.length && !dateSetsOverlap_(expectedDateSet, rowDateSet)) continue;

        var qtyAvailable = parseFloat(
          row.QtyAvailable || row.TotalInHouse || row.TotalAvailable || row.QuantityAvailable || row.Quantity || 0
        ) || 0;
        if (qtyAvailable <= 0) continue;

        candidates.push({
          siteName: rowSite,
          locationCode: rowLoc,
          lot: rowLot,
          dateCode: rowDateCode,
          qtyAvailable: qtyAvailable
        });
      }

      if (rows.length < 100) break;
    }
  } catch (e) {
    return [];
  }

  // WASP often returns aggregate no-lot rows alongside the real per-lot rows.
  // Ignore the aggregate row when a location already has lot-specific detail.
  var locHasLot = {};
  for (var j = 0; j < candidates.length; j++) {
    if (candidates[j].lot) locHasLot[candidates[j].locationCode] = true;
  }

  var filtered = [];
  for (var k = 0; k < candidates.length; k++) {
    if (locHasLot[candidates[k].locationCode] && !candidates[k].lot) continue;
    filtered.push(candidates[k]);
  }

  return filtered;
}

function resolveF3SourceRow_(itemNumber, qty, lotNumber, expectedDateCode, fromWasp) {
  var normalizedLot = String(lotNumber || '').trim();
  var normalizedDate = normalizeExactLotDate_(expectedDateCode);
  var candidates = fetchF3SourceCandidates_(itemNumber, fromWasp.site, normalizedLot, normalizedDate);

  if (candidates.length === 0) {
    return {
      success: false,
      error: normalizedLot
        ? ('Lot ' + normalizedLot + ' not found in ' + fromWasp.site)
        : ('No source stock found in ' + fromWasp.site + ' for ' + itemNumber)
    };
  }

  candidates.sort(function(a, b) {
    var aEnough = a.qtyAvailable >= qty ? 0 : 1;
    var bEnough = b.qtyAvailable >= qty ? 0 : 1;
    if (aEnough !== bEnough) return aEnough - bEnough;

    var aRank = getF3LocationRank_(a.siteName, a.locationCode, fromWasp.location);
    var bRank = getF3LocationRank_(b.siteName, b.locationCode, fromWasp.location);
    if (aRank !== bRank) return aRank - bRank;

    return b.qtyAvailable - a.qtyAvailable;
  });

  var siteTotal = 0;
  for (var i = 0; i < candidates.length; i++) {
    siteTotal += candidates[i].qtyAvailable;
    if (candidates[i].qtyAvailable >= qty) {
      return {
        success: true,
        site: candidates[i].siteName || fromWasp.site,
        location: candidates[i].locationCode || fromWasp.location,
        lot: candidates[i].lot || normalizedLot,
        dateCode: candidates[i].dateCode || normalizedDate,
        note: candidates[i].locationCode !== fromWasp.location
          ? ('resolved source ' + candidates[i].locationCode)
          : ''
      };
    }
  }

  if (siteTotal >= qty) {
    return {
      success: false,
      error: 'Stock split across multiple source rows in ' + fromWasp.site + ' for ' + itemNumber
    };
  }

  return {
    success: false,
    error: 'No source row with ' + qty + ' available in ' + fromWasp.site + ' for ' + itemNumber
  };
}

function buildF3HeaderLocations_(successItems, errors, defaultFrom, defaultTo) {
  var fromLocation = defaultFrom.location;
  var toLocation = defaultTo.location;
  var mixedFrom = false;
  var mixedTo = false;
  var seenAny = false;

  function scanItem(item) {
    if (!item) return;
    var itemFrom = item.fromWasp || defaultFrom;
    var itemTo = item.toWasp || defaultTo;

    if (!seenAny) {
      fromLocation = itemFrom.location || defaultFrom.location;
      toLocation = itemTo.location || defaultTo.location;
      seenAny = true;
      return;
    }

    if ((itemFrom.location || defaultFrom.location) !== fromLocation) mixedFrom = true;
    if ((itemTo.location || defaultTo.location) !== toLocation) mixedTo = true;
  }

  for (var i = 0; i < successItems.length; i++) scanItem(successItems[i]);
  for (var j = 0; j < errors.length; j++) scanItem(errors[j]);

  return {
    fromLocation: mixedFrom ? 'Mixed source' : fromLocation,
    toLocation: mixedTo ? 'Mixed dest' : toLocation,
    fromSite: defaultFrom.site,
    toSite: defaultTo.site
  };
}

// ============================================
// MAIN POLLING FUNCTION
// ============================================

function pollKatanaStockTransfers() {
  try {
    var cache = CacheService.getScriptCache();
    var lastSync = cache.get(F3_CONFIG.CACHE_KEY_LAST_SYNC);

    if (!lastSync) {
      lastSync = F3_CONFIG.MIN_CREATED_DATE;
    }

    var transfers = fetchKatanaStockTransfers(lastSync);

    if (!transfers || transfers.length === 0) {
      cache.put(F3_CONFIG.CACHE_KEY_LAST_SYNC, new Date().toISOString(), 21600);
      updateF3Summary();
      return { status: 'ok', message: 'No new transfers', count: 0 };
    }

    var synced = 0, errors = 0, skipped = 0;
    var minDate = new Date(F3_CONFIG.MIN_CREATED_DATE);

    for (var i = 0; i < transfers.length; i++) {
      var transfer = transfers[i];

      // Skip old STs before our start date
      var createdAt = new Date(transfer.created_at);
      if (createdAt < minDate) {
        skipped++;
        continue;
      }

      var result = processStockTransfer(transfer);
      if (result.synced) synced++;
      if (result.error) errors++;
    }

    cache.put(F3_CONFIG.CACHE_KEY_LAST_SYNC, new Date().toISOString(), 21600);
    updateF3Summary();

    return { status: 'ok', total: transfers.length, synced: synced, errors: errors, skipped: skipped };

  } catch (error) {
    logToSheet('F3_POLL_ERROR', { error: error.message }, {});
    return { status: 'error', message: error.message };
  }
}

// ============================================
// KATANA API FUNCTIONS
// ============================================

function fetchKatanaStockTransfers(since) {
  var endpoint = F3_CONFIG.KATANA_ENDPOINT;
  var params = [];

  if (since) params.push('updated_at[gte]=' + encodeURIComponent(since));
  params.push('per_page=100');
  params.push('sort=-updated_at');

  if (params.length > 0) endpoint += '?' + params.join('&');

  var result = katanaApiCall(endpoint);

  if (!result || !result.data) return [];
  return result.data;
}

function fetchKatanaLocation(locationId) {
  if (!locationId) return null;

  var cache = CacheService.getScriptCache();
  var cacheKey = 'KATANA_LOC_' + locationId;
  var cached = cache.get(cacheKey);

  if (cached) return JSON.parse(cached);

  var result = katanaApiCall('locations/' + locationId);

  if (result) {
    cache.put(cacheKey, JSON.stringify(result), 3600);
    return result;
  }

  return null;
}

function fetchKatanaVariantF3(variantId) {
  if (!variantId) return null;

  var cache = CacheService.getScriptCache();
  var cacheKey = 'KATANA_VAR_' + variantId;
  var cached = cache.get(cacheKey);

  if (cached) return JSON.parse(cached);

  var result = katanaApiCall('variants/' + variantId);

  if (result) {
    // Enrich variant with uom from parent product/material so resolveVariantUom()
    // hits the fast path (no extra API calls during ST processing loop).
    if (!result.uom) {
      var parentId = result.material_id || result.product_id;
      var parentEndpoint = result.material_id ? 'materials/' : 'products/';
      if (parentId) {
        var parentData = katanaApiCall(parentEndpoint + parentId);
        if (parentData) {
          var parent = parentData.data || parentData;
          result.uom = parent.uom || '';
        }
      }
    }
    cache.put(cacheKey, JSON.stringify(result), 3600);
    return result;
  }

  return null;
}

function fetchKatanaBatchForF3(batchId) {
  if (!batchId) return null;

  var cache = CacheService.getScriptCache();
  var cacheKey = 'KATANA_BATCH_' + batchId;
  var cached = cache.get(cacheKey);

  if (cached) return JSON.parse(cached);

  var result = katanaApiCall('batch_stocks?batch_id=' + batchId);

  if (result && result.data && result.data.length > 0) {
    for (var i = 0; i < result.data.length; i++) {
      var batch = result.data[i];
      var foundId = batch.batch_id || batch.id;
      if (String(foundId) !== String(batchId)) continue;
      cache.put(cacheKey, JSON.stringify(batch), 3600);
      return batch;
    }
  }

  return null;
}

// ============================================
// TRANSFER PROCESSING - SUPPORTS PARTIAL RECEIVE
// ============================================

function processStockTransfer(transfer) {
  var stId = transfer.id;
  var stNumber = transfer.stock_transfer_number || ('ST-' + stId);
  var status = (transfer.status || '').toLowerCase();

  var fromLocId = transfer.source_location_id;
  var toLocId = transfer.target_location_id;

  var fromLoc = fetchKatanaLocation(fromLocId);
  var toLoc = fetchKatanaLocation(toLocId);

  var fromLocName = fromLoc ? fromLoc.name : 'Unknown';
  var toLocName = toLoc ? toLoc.name : 'Unknown';

  var rows = transfer.stock_transfer_rows || [];

  // Update or create ST in sheet
  var existingRow = findSTInSheet(stId);

  if (existingRow) {
    updateSTInSheet(existingRow, {
      status: status,
      updated: transfer.updated_at,
      itemCount: rows.length,
      fromLoc: fromLocName,
      toLoc: toLocName
    });
  } else {
    addSTToSheet({
      stId: stId, stNumber: stNumber, status: status,
      fromLoc: fromLocName, toLoc: toLocName, itemCount: rows.length,
      created: transfer.created_at, updated: transfer.updated_at
    });
    existingRow = findSTInSheet(stId);
  }

  // Check if this transfer was cancelled after being synced — reverse the WASP moves
  var shouldReverse = F3_REVERSE_STATUS.indexOf(status) >= 0;
  if (shouldReverse) {
    if (existingRow) {
      var rvWaspSt = getSTWaspStatus(existingRow);
      // Reverse if fully or partially synced — partial means some items moved and need undoing
      if (rvWaspSt === 'Synced' || rvWaspSt === 'Partial') {
        return reverseSTFromWasp(stId, stNumber, fromLocName, toLocName, rows, existingRow);
      }
      // Already handled (Reversed/Skip/Failed/Error/Pending) — don't overwrite on repeated polls
      if (rvWaspSt !== 'Reversed' && rvWaspSt !== 'Skip' && rvWaspSt !== 'Failed') {
        updateSTWaspStatus(existingRow, 'Skip', 'Cancelled — not previously synced');
      }
    }
    return { processed: true, synced: false };
  }

  // Check if we should sync (status indicates items are being/have been received)
  var shouldSync = F3_CONFIG.SYNC_ON_STATUS.indexOf(status) >= 0;

  if (shouldSync) {
    if (existingRow) {
      var currentWaspStatus = getSTWaspStatus(existingRow);

      // Skip permanently resolved statuses.
      // 'Partial' is included — retrying partial syncs re-moves already-synced items (double-deduction).
      // Use forceSyncST() to manually reset and retry after fixing the source inventory.
      if (currentWaspStatus === 'Synced' || currentWaspStatus === 'Failed'
          || currentWaspStatus === 'Skip' || currentWaspStatus === 'Reversed'
          || currentWaspStatus === 'Partial') {
        return { processed: true, synced: false };
      }

      // Throttle retries for Error STs — escalating backoff (1h, 6h, 24h)
      if (currentWaspStatus === 'Error') {
        var currentNotes = getF3Sheet().getRange(existingRow, F3_CONFIG.COLS.NOTES).getValue();
        var retryCount = parseRetryCount(currentNotes);

        // Already hit max retries → escalate to permanent "Failed"
        if (retryCount >= MAX_TRANSIENT_RETRIES) {
          updateSTWaspStatus(existingRow, 'Failed', currentNotes.replace(/^\[R\d+\]\s*/, '[MAXRETRY] '));
          return { processed: true, synced: false };
        }

        // Check backoff timing
        var lastAttemptVal = getSTLastAttempt(existingRow);
        if (lastAttemptVal) {
          var hoursSince = (new Date() - new Date(lastAttemptVal)) / (1000 * 60 * 60);
          var backoffHours = RETRY_BACKOFF_HOURS[retryCount] || 24;
          if (hoursSince < backoffHours) {
            return { processed: true, synced: false }; // Too soon to retry
          }
        }
      }
    }

    var isRetry = existingRow && getSTWaspStatus(existingRow) === 'Error';
    return syncSTToWasp(stId, stNumber, fromLocName, toLocName, rows, existingRow, status, isRetry);
  }

  return { processed: true, synced: false };
}

function syncSTToWasp(stId, stNumber, fromLocKatana, toLocKatana, rows, sheetRow, katanaStatus, isRetry) {
  // Skip virtual/external locations (Amazon FBA, Shopify channel) — no WASP equivalent
  var fromLower = (fromLocKatana || '').toLowerCase().trim();
  var toLower = (toLocKatana || '').toLowerCase().trim();
  for (var sl = 0; sl < F3_SKIP_LOCATIONS.length; sl++) {
    if (fromLower.indexOf(F3_SKIP_LOCATIONS[sl]) >= 0 || toLower.indexOf(F3_SKIP_LOCATIONS[sl]) >= 0) {
      updateSTWaspStatus(sheetRow, 'Skip', 'Virtual location — no WASP sync needed');
      return { synced: false, error: null };
    }
  }

  var fromWasp = mapKatanaToWaspSite(fromLocKatana);
  var toWasp = mapKatanaToWaspSite(toLocKatana);

  if (!fromWasp) {
    var error = 'Unknown source: ' + fromLocKatana;
    updateSTWaspStatus(sheetRow, 'Failed', error);
    return { synced: false, error: error };
  }

  if (!toWasp) {
    var error2 = 'Unknown dest: ' + toLocKatana;
    updateSTWaspStatus(sheetRow, 'Failed', error2);
    return { synced: false, error: error2 };
  }

  var errors = [], successItems = [];

  for (var i = 0; i < rows.length; i++) {
    var row = rows[i];

    var variantId = row.variant_id;
    var variant = fetchKatanaVariantF3(variantId);
    var sku = variant ? (variant.sku || variant.name || '') : '';
    var variantUom = resolveVariantUom(variant);
    var qty = parseFloat(row.quantity) || 0;

    var bt = row.batch_transactions || [];

    if (bt.length > 1) {
      // MULTI-BATCH: Process each batch_transaction separately
      markSyncedToWasp(sku, fromWasp.location, 'remove');
      markSyncedToWasp(sku, toWasp.location, 'add');

      for (var btIdx = 0; btIdx < bt.length; btIdx++) {
        var btEntry = bt[btIdx];
        var btQty = parseFloat(btEntry.quantity) || 0;
        if (btQty <= 0) continue;

        var btLot = '';
        var btExpiry = '';
        var btBatchId = btEntry.batch_id || null;
        if (btBatchId) {
          var btBatch = fetchKatanaBatchForF3(btBatchId);
          if (btBatch) {
            btLot = btBatch.batch_number || btBatch.nr || '';
            btExpiry = normalizeBusinessDate_(btBatch.expiration_date || btBatch.expiry_date || '');
          }
        }

        var btResult = syncSTLineItem(stId, stNumber, sku, btQty, btLot, btExpiry, fromWasp, toWasp);

        if (btResult.success) {
          successItems.push({
            sku: sku, qty: btQty, lot: btResult.lot || btLot, expiry: btResult.expiry || btExpiry,
            multiBatch: true, multiBatchTotal: qty,
            multiBatchFirst: (btIdx === 0), multiBatchLast: (btIdx === bt.length - 1),
            uom: variantUom,
            fromWasp: btResult.fromWasp || fromWasp,
            toWasp: btResult.toWasp || toWasp
          });
        } else {
          errors.push({
            sku: sku, qty: btQty, lot: btResult.lot || btLot, expiry: btResult.expiry || btExpiry,
            uom: variantUom, error: btResult.error, permanent: isPermanentError(btResult.error),
            fromWasp: btResult.fromWasp || fromWasp,
            toWasp: btResult.toWasp || toWasp
          });
        }
      }
    } else {
      // SINGLE-BATCH or NO-BATCH: existing behavior
      var lot = '';
      var expiry = '';

      if (bt.length > 0) {
        var batchId = bt[0].batch_id;
        var batch = fetchKatanaBatchForF3(batchId);
        if (batch) {
          lot = batch.batch_number || batch.nr || '';
          expiry = normalizeBusinessDate_(batch.expiration_date || batch.expiry_date || '');
        }
      }

      var result = syncSTLineItem(stId, stNumber, sku, qty, lot, expiry, fromWasp, toWasp);

      if (result.success) {
        successItems.push({
          sku: sku, qty: qty, lot: result.lot || lot, expiry: result.expiry || expiry, uom: variantUom,
          fromWasp: result.fromWasp || fromWasp,
          toWasp: result.toWasp || toWasp
        });
      } else {
        errors.push({
          sku: sku, qty: qty, lot: result.lot || lot, expiry: result.expiry || expiry, uom: variantUom,
          error: result.error, permanent: isPermanentError(result.error),
          fromWasp: result.fromWasp || fromWasp,
          toWasp: result.toWasp || toWasp
        });
      }
    }
  }

  // Determine final status
  var totalItems = rows.length;
  var syncedNow = successItems.length;
  var totalSynced = syncedNow;
  var failedCount = errors.length;

  // Check if ALL errors are permanent
  var allPermanent = failedCount > 0;
  for (var ep = 0; ep < errors.length; ep++) {
    if (!errors[ep].permanent) { allPermanent = false; break; }
  }

  // Build notes
  var notes = '';

  if (syncedNow > 0) {
    notes = formatItemSummary(successItems);
  }

  if (failedCount > 0) {
    var failNotes = errors.map(function(e) { return e.sku + ': ' + e.error; }).join(', ');
    notes += (notes ? ' | ' : '') + 'FAILED: ' + failNotes;
  }

  // Determine WASP status
  var waspStatus;
  if (failedCount > 0 && totalSynced === 0) {
    if (allPermanent) {
      // All errors are permanent — mark as Failed, never retry
      waspStatus = 'Failed';
    } else {
      // Has transient errors — mark as Error, will retry with backoff
      waspStatus = 'Error';
      // Increment retry counter
      var prevNotes = sheetRow ? getF3Sheet().getRange(sheetRow, F3_CONFIG.COLS.NOTES).getValue() : '';
      var prevRetry = parseRetryCount(prevNotes);
      notes = prependRetryCount(prevRetry + 1, notes);
    }
  } else if (totalSynced === totalItems && failedCount === 0) {
    waspStatus = 'Synced';
    if (!notes) notes = 'All ' + totalItems + ' items synced';
  } else if (totalSynced > 0) {
    waspStatus = 'Partial';
    notes = totalSynced + '/' + totalItems + ' synced' + (notes ? ': ' + notes : '');
  } else {
    waspStatus = 'Pending';
  }

  updateSTWaspStatus(sheetRow, waspStatus, notes);

  // Activity Log — F3 Transfer
  // Skip Activity log on retry failures (don't flood the log)
  // Log on: first failure, any success, or retry that finally succeeds
  var shouldLogActivity = (syncedNow > 0) || (failedCount > 0 && !isRetry);
  if (shouldLogActivity) {
    var f3HeaderLocs = buildF3HeaderLocations_(successItems, errors, fromWasp, toWasp);
    var f3pSubItems = [];
    for (var s = 0; s < successItems.length; s++) {
      var sItem = successItems[s];
      var sFrom = sItem.fromWasp || fromWasp;
      var sTo = sItem.toWasp || toWasp;

      if (sItem.multiBatch && sItem.multiBatchFirst) {
        // Multi-batch parent row
        var f3BatchCount = 0;
        for (var f3bc = s; f3bc < successItems.length; f3bc++) {
          if (successItems[f3bc].sku === sItem.sku && successItems[f3bc].multiBatch) f3BatchCount++;
          else if (f3bc > s) break;
        }
        f3pSubItems.push({
          sku: sItem.sku,
          qty: sItem.multiBatchTotal,
          uom: sItem.uom || '',
          success: true,
          status: '',
          action: sFrom.location + ' → ' + sTo.location,
          qtyColor: 'grey',
          isParent: true,
          batchCount: f3BatchCount
        });
        f3pSubItems[f3pSubItems.length - 1].action = getActivityDisplayLocation_(sFrom.location) + ' -> ' + getActivityDisplayLocation_(sTo.location);
      }

      if (sItem.multiBatch) {
        // Nested batch sub-row
        var f3NestedAction = '';
        if (sItem.lot) f3NestedAction += 'lot:' + sItem.lot;
        if (sItem.expiry) f3NestedAction += (f3NestedAction ? '  ' : '') + 'exp:' + sItem.expiry;
        f3pSubItems.push({
          sku: '',
          qty: sItem.qty,
          uom: sItem.uom || '',
          success: true,
          status: 'Synced',
          action: f3NestedAction,
          qtyColor: 'green',
          nested: true
        });
      } else {
        f3pSubItems.push({
          sku: sItem.sku,
          qty: sItem.qty,
          uom: sItem.uom || '',
          success: true,
          status: 'Synced',
          action: sFrom.location + ' → ' + sTo.location
            + (sItem.lot ? '  lot:' + sItem.lot : '')
            + (sItem.expiry ? '  exp:' + sItem.expiry : ''),
          qtyColor: 'green'
        });
        f3pSubItems[f3pSubItems.length - 1].action = buildActivityActionText_(
          'move ' + getActivityDisplayLocation_(sFrom.location) + ' -> ' + getActivityDisplayLocation_(sTo.location),
          sItem.lot,
          sItem.expiry
        );
      }
    }
    for (var e = 0; e < errors.length; e++) {
      var eItem = errors[e];
      var eFrom = eItem.fromWasp || fromWasp;
      var eTo = eItem.toWasp || toWasp;
      f3pSubItems.push({
        sku: eItem.sku,
        qty: eItem.qty || '',
        uom: eItem.uom || '',
        success: false,
        status: 'Failed',
        action: eFrom.location + ' → ' + eTo.location
          + (eItem.lot    ? '  lot:' + eItem.lot    : '')
          + (eItem.expiry ? '  exp:' + eItem.expiry : ''),
        error: eItem.error
      });
      f3pSubItems[f3pSubItems.length - 1].action = buildActivityActionText_(
        'move ' + getActivityDisplayLocation_(eFrom.location) + ' -> ' + getActivityDisplayLocation_(eTo.location),
        eItem.lot,
        eItem.expiry
      );
    }

    var f3pStatus = failedCount === 0 ? 'success' : syncedNow === 0 ? 'failed' : 'partial';
    var f3pDetail = stNumber + '  ' + totalItems + ' item' + (totalItems !== 1 ? 's' : '');
    if (failedCount > 0) f3pDetail += '  ' + failedCount + ' error' + (failedCount > 1 ? 's' : '');
    stNumber = extractCanonicalActivityRef_(stNumber, 'ST-', stId);
    f3pDetail = joinActivitySegments_([stNumber, buildActivityCountSummary_(totalItems, 'line', 'lines', 'moved')]);

    var f3pContext = f3HeaderLocs.fromLocation + ' @ ' + f3HeaderLocs.fromSite
      + ' → ' + f3HeaderLocs.toLocation + ' @ ' + f3HeaderLocs.toSite;
    var f3FromLabel = f3HeaderLocs.fromLocation === 'Mixed source' ? 'mixed-source' : getActivityDisplayLocation_(f3HeaderLocs.fromLocation);
    var f3ToLabel = f3HeaderLocs.toLocation === 'Mixed dest' ? 'mixed-dest' : getActivityDisplayLocation_(f3HeaderLocs.toLocation);
    f3pContext = f3FromLabel + ' -> ' + f3ToLabel;
    var f3pExecId = logActivity('F3', f3pDetail, f3pStatus, f3pContext, f3pSubItems, {
      text: stNumber,
      url: getKatanaWebUrl('st', stId)
    });
    // Log to F3 detail tab
    var f3FlowItems = [];
    for (var si = 0; si < successItems.length; si++) {
      var siItem = successItems[si];
      var siFrom = siItem.fromWasp || fromWasp;
      var siTo = siItem.toWasp || toWasp;

      if (siItem.multiBatch && siItem.multiBatchFirst) {
        // Multi-batch parent row
        var f3fBatchCount = 0;
        for (var f3fbc = si; f3fbc < successItems.length; f3fbc++) {
          if (successItems[f3fbc].sku === siItem.sku && successItems[f3fbc].multiBatch) f3fBatchCount++;
          else if (f3fbc > si) break;
        }
        f3FlowItems.push({
          sku: siItem.sku,
          qty: siItem.multiBatchTotal,
          uom: siItem.uom || '',
          detail: siFrom.location + ' → ' + siTo.location,
          status: '',
          error: '',
          qtyColor: 'grey',
          isParent: true,
          batchCount: f3fBatchCount
        });
      }

      if (siItem.multiBatch) {
        // Nested batch sub-row
        var f3fNestedDetail = '';
        if (siItem.lot) f3fNestedDetail += 'lot:' + siItem.lot;
        if (siItem.expiry) f3fNestedDetail += (f3fNestedDetail ? '  ' : '') + 'exp:' + siItem.expiry;
        f3FlowItems.push({
          sku: '',
          qty: siItem.qty,
          uom: siItem.uom || '',
          detail: f3fNestedDetail,
          status: 'Synced',
          error: '',
          qtyColor: 'green',
          nested: true
        });
      } else {
        f3FlowItems.push({
          sku: siItem.sku,
          qty: siItem.qty,
          uom: siItem.uom || '',
          detail: siFrom.location + ' → ' + siTo.location + (siItem.lot ? '  lot:' + siItem.lot : '')
            + (siItem.expiry ? '  exp:' + siItem.expiry : ''),
          status: 'Synced',
          error: '',
          qtyColor: 'green'
        });
      }
    }
    for (var ei = 0; ei < errors.length; ei++) {
      var flowError = errors[ei];
      var flowFrom = flowError.fromWasp || fromWasp;
      var flowTo = flowError.toWasp || toWasp;
      f3FlowItems.push({
        sku: flowError.sku,
        qty: flowError.qty || '',
        uom: flowError.uom || '',
        detail: flowFrom.location + ' → ' + flowTo.location
          + (flowError.lot    ? '  lot:' + flowError.lot    : '')
          + (flowError.expiry ? '  exp:' + flowError.expiry : ''),
        status: 'Failed',
        error: flowError.error || ''
      });
    }
    logFlowDetail('F3', f3pExecId, {
      ref: stNumber,
      detail: totalItems + ' items  ' + f3HeaderLocs.fromLocation + ' → ' + f3HeaderLocs.toLocation,
      status: f3pStatus === 'success' ? 'Synced' : f3pStatus === 'failed' ? 'Failed' : 'Partial',
      linkText: stNumber,
      linkUrl: getKatanaWebUrl('st', stId)
    }, f3FlowItems);
  }

  return {
    synced: syncedNow > 0,
    itemCount: syncedNow,
    errors: errors.length > 0 ? errors : null
  };
}

// ============================================
// F3 REVERSE — UNDO A PREVIOUSLY SYNCED TRANSFER
// ============================================
// Called when a Katana ST transitions to 'cancelled'/'voided' after WASP status = 'Synced'.
// Removes from original destination, adds back to original source.
// ============================================

function reverseSTFromWasp(stId, stNumber, fromLocKatana, toLocKatana, rows, sheetRow) {
  // Skip virtual/external locations
  var fromLower = (fromLocKatana || '').toLowerCase().trim();
  var toLower   = (toLocKatana   || '').toLowerCase().trim();
  for (var sl = 0; sl < F3_SKIP_LOCATIONS.length; sl++) {
    if (fromLower.indexOf(F3_SKIP_LOCATIONS[sl]) >= 0 || toLower.indexOf(F3_SKIP_LOCATIONS[sl]) >= 0) {
      updateSTWaspStatus(sheetRow, 'Skip', 'Virtual location — reverse not needed');
      return { synced: false, error: null };
    }
  }

  var fromWasp = mapKatanaToWaspSite(fromLocKatana); // original source
  var toWasp   = mapKatanaToWaspSite(toLocKatana);   // original destination

  if (!fromWasp) {
    updateSTWaspStatus(sheetRow, 'Skip', 'Reverse skipped — unknown source: ' + fromLocKatana);
    return { synced: false, error: null };
  }
  if (!toWasp) {
    updateSTWaspStatus(sheetRow, 'Skip', 'Reverse skipped — unknown dest: ' + toLocKatana);
    return { synced: false, error: null };
  }

  var errors = [], successItems = [];

  for (var i = 0; i < rows.length; i++) {
    var row = rows[i];
    var variantId = row.variant_id;
    var variant = fetchKatanaVariantF3(variantId);
    var sku = variant ? (variant.sku || variant.name || '') : '';
    var variantUom = resolveVariantUom(variant);
    var qty = parseFloat(row.quantity) || 0;
    var bt = row.batch_transactions || [];

    if (bt.length > 1) {
      // Multi-batch: pre-mark then reverse each entry
      markSyncedToWasp(sku, toWasp.location, 'remove');
      markSyncedToWasp(sku, fromWasp.location, 'add');

      for (var btIdx = 0; btIdx < bt.length; btIdx++) {
        var btEntry = bt[btIdx];
        var btQty = parseFloat(btEntry.quantity) || 0;
        if (btQty <= 0) continue;

        var btLot = '', btExpiry = '';
        var btBatchId = btEntry.batch_id || null;
        if (btBatchId) {
          var btBatch = fetchKatanaBatchForF3(btBatchId);
          if (btBatch) {
            btLot   = btBatch.batch_number || btBatch.nr || '';
            btExpiry = normalizeBusinessDate_(btBatch.expiration_date || btBatch.expiry_date || '');
          }
        }

        // REVERSED: remove from toWasp (original dest), add to fromWasp (original source)
        var btResult = syncSTLineItem(stId, stNumber, sku, btQty, btLot, btExpiry, toWasp, fromWasp);
        if (btResult.success) {
          successItems.push({ sku: sku, qty: btQty, lot: btLot, expiry: btExpiry, multiBatch: true,
            multiBatchTotal: qty, multiBatchFirst: (btIdx === 0), multiBatchLast: (btIdx === bt.length - 1),
            uom: variantUom });
        } else {
          errors.push({ sku: sku, qty: btQty, lot: btLot, expiry: btExpiry, uom: variantUom, error: btResult.error, permanent: isPermanentError(btResult.error) });
        }
      }
    } else {
      var lot = '', expiry = '';
      if (bt.length > 0) {
        var batchId = bt[0].batch_id;
        var batch = fetchKatanaBatchForF3(batchId);
        if (batch) {
          lot    = batch.batch_number || batch.nr || '';
          expiry = normalizeBusinessDate_(batch.expiration_date || batch.expiry_date || '');
        }
      }

      // REVERSED: remove from toWasp (original dest), add to fromWasp (original source)
      var result = syncSTLineItem(stId, stNumber, sku, qty, lot, expiry, toWasp, fromWasp);
      if (result.success) {
        successItems.push({ sku: sku, qty: qty, lot: lot, expiry: expiry, uom: variantUom });
      } else {
        errors.push({ sku: sku, qty: qty, lot: lot, expiry: expiry, uom: variantUom, error: result.error, permanent: isPermanentError(result.error) });
      }
    }
  }

  var totalItems  = rows.length;
  var syncedNow   = successItems.length;
  var failedCount = errors.length;

  var notes = '';
  if (syncedNow > 0) notes = formatItemSummary(successItems);
  if (failedCount > 0) {
    var failNotes = errors.map(function(e) { return e.sku + ': ' + e.error; }).join(', ');
    notes += (notes ? ' | ' : '') + 'FAILED: ' + failNotes;
  }

  var waspStatus;
  if (failedCount > 0 && syncedNow === 0) {
    waspStatus = 'Failed';
  } else if (syncedNow === totalItems && failedCount === 0) {
    waspStatus = 'Reversed';
  } else {
    waspStatus = 'Partial';
    notes = syncedNow + '/' + totalItems + ' reversed' + (notes ? ': ' + notes : '');
  }

  updateSTWaspStatus(sheetRow, waspStatus, notes);

  // Activity Log — show reversal direction (toWasp → fromWasp)
  if (syncedNow > 0 || failedCount > 0) {
    var f3SubItems = [];
    for (var s = 0; s < successItems.length; s++) {
      var sItem = successItems[s];
      if (sItem.multiBatch && sItem.multiBatchFirst) {
        var f3BatchCount = 0;
        for (var f3bc = s; f3bc < successItems.length; f3bc++) {
          if (successItems[f3bc].sku === sItem.sku && successItems[f3bc].multiBatch) f3BatchCount++;
          else if (f3bc > s) break;
        }
        f3SubItems.push({ sku: sItem.sku, qty: sItem.multiBatchTotal, uom: sItem.uom || '',
          success: true, status: '', action: toWasp.location + ' \u2192 ' + fromWasp.location,
          qtyColor: 'grey', isParent: true, batchCount: f3BatchCount });
        f3SubItems[f3SubItems.length - 1].action = getActivityDisplayLocation_(toWasp.location) + ' -> ' + getActivityDisplayLocation_(fromWasp.location);
      }
      if (sItem.multiBatch) {
        var nestedAction = '';
        if (sItem.lot)    nestedAction += 'lot:' + sItem.lot;
        if (sItem.expiry) nestedAction += (nestedAction ? '  ' : '') + 'exp:' + sItem.expiry;
        f3SubItems.push({ sku: '', qty: sItem.qty, uom: sItem.uom || '', success: true, status: 'Reversed',
          action: nestedAction, qtyColor: 'green', nested: true });
      } else {
        f3SubItems.push({ sku: sItem.sku, qty: sItem.qty, uom: sItem.uom || '',
          success: true, status: 'Reversed',
          action: toWasp.location + ' \u2192 ' + fromWasp.location
            + (sItem.lot    ? '  lot:' + sItem.lot    : '')
            + (sItem.expiry ? '  exp:' + sItem.expiry : ''),
          qtyColor: 'green' });
        f3SubItems[f3SubItems.length - 1].action = buildActivityActionText_(
          'move ' + getActivityDisplayLocation_(toWasp.location) + ' -> ' + getActivityDisplayLocation_(fromWasp.location),
          sItem.lot,
          sItem.expiry
        );
      }
    }
    for (var e = 0; e < errors.length; e++) {
      f3SubItems.push({ sku: errors[e].sku, qty: errors[e].qty || '', uom: errors[e].uom || '', success: false, status: 'Failed',
        action: toWasp.location + ' \u2192 ' + fromWasp.location, error: errors[e].error });
      f3SubItems[f3SubItems.length - 1].action = buildActivityActionText_(
        'move ' + getActivityDisplayLocation_(toWasp.location) + ' -> ' + getActivityDisplayLocation_(fromWasp.location),
        errors[e].lot,
        errors[e].expiry
      );
    }

    var f3pStatus = failedCount === 0 ? 'success' : syncedNow === 0 ? 'failed' : 'partial';
    var f3pDetail = stNumber + ' (reversed)  ' + totalItems + ' item' + (totalItems !== 1 ? 's' : '');
    if (failedCount > 0) f3pDetail += '  ' + failedCount + ' error' + (failedCount > 1 ? 's' : '');
    var f3pContext = toWasp.location + ' @ ' + toWasp.site + ' \u2192 ' + fromWasp.location + ' @ ' + fromWasp.site;
    stNumber = extractCanonicalActivityRef_(stNumber, 'ST-', stId);
    f3pDetail = joinActivitySegments_([stNumber, buildActivityCountSummary_(totalItems, 'line', 'lines', 'reversed')]);
    f3pContext = getActivityDisplayLocation_(toWasp.location) + ' -> ' + getActivityDisplayLocation_(fromWasp.location);
    logActivity('F3', f3pDetail, f3pStatus, f3pContext, f3SubItems, {
      text: stNumber, url: getKatanaWebUrl('st', stId)
    });
  }

  return { synced: syncedNow > 0, itemCount: syncedNow, errors: errors.length > 0 ? errors : null };
}

function syncSTLineItem(stId, stNumber, sku, qty, lot, expiry, fromWasp, toWasp) {
  var notes = 'F3:' + stNumber;

  if (!sku || qty <= 0) {
    return { success: false, error: 'Invalid SKU/qty' };
  }

  var resolvedSource = resolveF3SourceRow_(sku, qty, lot, expiry, fromWasp);
  if (!resolvedSource.success) {
    logToSheet('F3_SYNC_ERROR', { sku: sku, error: resolvedSource.error, stNumber: stNumber }, {});
    return {
      success: false,
      error: resolvedSource.error,
      fromWasp: fromWasp,
      toWasp: toWasp,
      lot: lot || '',
      expiry: expiry || ''
    };
  }

  var actualFromWasp = {
    site: resolvedSource.site || fromWasp.site,
    location: resolvedSource.location || fromWasp.location
  };
  var actualLot = resolvedSource.lot || lot || '';
  var actualExpiry = resolvedSource.dateCode || expiry || '';

  // Mark sync BEFORE WASP calls — prevents F2 echo when WASP callouts fire
  markSyncedToWasp(sku, actualFromWasp.location, 'remove');
  markSyncedToWasp(sku, toWasp.location, 'add');

  // Step 1: REMOVE from source
  var removeResult = waspRemoveInventoryFromSite(
    actualFromWasp.site, sku, qty, actualFromWasp.location, actualLot, actualExpiry, notes
  );

  if (!removeResult.success) {
    var removeError = parseWaspError(removeResult.response, 'Remove', sku, actualFromWasp.location);
    logToSheet('F3_SYNC_ERROR', { sku: sku, error: removeError, stNumber: stNumber }, {});
    return {
      success: false,
      error: removeError,
      fromWasp: actualFromWasp,
      toWasp: toWasp,
      lot: actualLot,
      expiry: actualExpiry
    };
  }

  // Step 2: ADD to destination
  var addResult = waspAddInventoryToSite(
    toWasp.site, sku, qty, toWasp.location, actualLot, actualExpiry, notes
  );

  if (!addResult.success) {
    var addError = parseWaspError(addResult.response, 'Add', sku, toWasp.location);

    // ROLLBACK
    waspAddInventoryToSite(actualFromWasp.site, sku, qty, actualFromWasp.location, actualLot, actualExpiry, 'ROLLBACK:' + notes);

    logToSheet('F3_SYNC_ERROR', { sku: sku, error: addError, stNumber: stNumber }, {});
    return {
      success: false,
      error: addError,
      fromWasp: actualFromWasp,
      toWasp: toWasp,
      lot: actualLot,
      expiry: actualExpiry
    };
  }

  // SUCCESS
  return {
    success: true,
    fromWasp: actualFromWasp,
    toWasp: toWasp,
    lot: actualLot,
    expiry: actualExpiry,
    note: resolvedSource.note || ''
  };
}

// ============================================
// SHEET FUNCTIONS
// ============================================

function getF3Sheet() {
  var ss = SpreadsheetApp.openById(F3_CONFIG.SHEET_ID);
  var sheet = ss.getSheetByName(F3_CONFIG.TAB_NAME);
  if (!sheet) { createF3Sheet(); sheet = ss.getSheetByName(F3_CONFIG.TAB_NAME); }
  return sheet;
}

function findSTInSheet(stId) {
  var sheet = getF3Sheet();
  var data = sheet.getDataRange().getValues();
  for (var i = 2; i < data.length; i++) {
    if (data[i][F3_CONFIG.COLS.ST_ID - 1] == stId) return i + 1;
  }
  return null;
}

function addSTToSheet(data) {
  var sheet = getF3Sheet();
  sheet.appendRow([
    data.stId, data.stNumber, data.status, data.fromLoc, data.toLoc,
    data.itemCount + ' items', data.created, data.updated, 'Pending', '', ''
  ]);
}

function updateSTInSheet(rowNum, data) {
  var sheet = getF3Sheet();
  if (data.status !== undefined) sheet.getRange(rowNum, F3_CONFIG.COLS.STATUS).setValue(data.status);
  if (data.updated !== undefined) sheet.getRange(rowNum, F3_CONFIG.COLS.UPDATED).setValue(data.updated);
  if (data.itemCount !== undefined) sheet.getRange(rowNum, F3_CONFIG.COLS.ITEM_COUNT).setValue(data.itemCount + ' items');
  if (data.fromLoc !== undefined) sheet.getRange(rowNum, F3_CONFIG.COLS.FROM_LOC).setValue(data.fromLoc);
  if (data.toLoc !== undefined) sheet.getRange(rowNum, F3_CONFIG.COLS.TO_LOC).setValue(data.toLoc);
}

function getSTWaspStatus(rowNum) {
  if (!rowNum) return null;
  return getF3Sheet().getRange(rowNum, F3_CONFIG.COLS.WASP_STATUS).getValue();
}

function getSTLastAttempt(rowNum) {
  if (!rowNum) return null;
  return getF3Sheet().getRange(rowNum, F3_CONFIG.COLS.WASP_SYNCED).getValue();
}

function updateSTWaspStatus(rowNum, status, notes) {
  if (!rowNum) return;
  var sheet = getF3Sheet();
  sheet.getRange(rowNum, F3_CONFIG.COLS.WASP_STATUS).setValue(status);
  sheet.getRange(rowNum, F3_CONFIG.COLS.NOTES).setValue(notes);
  // Always set timestamp — used for retry throttling on errors
  sheet.getRange(rowNum, F3_CONFIG.COLS.WASP_SYNCED).setValue(new Date().toISOString());
}

function updateF3Summary() {
  var sheet = getF3Sheet();
  var data = sheet.getDataRange().getValues();
  var counts = { Pending: 0, Syncing: 0, Synced: 0, Error: 0, Partial: 0, Failed: 0, Skip: 0, Reversed: 0 };

  for (var i = 2; i < data.length; i++) {
    var status = data[i][F3_CONFIG.COLS.WASP_STATUS - 1] || '';
    if (counts[status] !== undefined) counts[status]++;
  }

  var summary = 'Pending:' + counts.Pending + ' | Syncing:' + counts.Syncing + ' | Synced:' + counts.Synced + ' | Partial:' + counts.Partial + ' | Error:' + counts.Error + ' | Failed:' + counts.Failed + ' | Skip:' + counts.Skip + ' | Reversed:' + counts.Reversed;
  sheet.getRange(2, 1).setValue('SUMMARY');
  sheet.getRange(2, F3_CONFIG.COLS.WASP_STATUS).setValue(summary);
  sheet.getRange(2, 1, 1, 11).setBackground('#e8e8e8').setFontWeight('bold');
}

// ============================================
// SHEET CREATION
// ============================================

function createF3Sheet() {
  var ss = SpreadsheetApp.openById(F3_CONFIG.SHEET_ID);
  if (ss.getSheetByName(F3_CONFIG.TAB_NAME)) return;

  var sheet = ss.insertSheet(F3_CONFIG.TAB_NAME);
  var headers = ['Katana ID', 'ST Number', 'Status', 'From', 'To', 'Items', 'Created', 'Updated', 'WASP Status', 'Synced At', 'Notes'];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#4285f4').setFontColor('white');
  sheet.getRange(2, 1).setValue('SUMMARY');
  sheet.getRange(2, 1, 1, headers.length).setBackground('#e8e8e8').setFontWeight('bold');
  sheet.setFrozenRows(2);
}

// createF3ItemsSheet removed - no longer tracking individual items
// createF3ArchiveSheet removed - no longer archiving transfers

function setupF3Formatting() {
  var sheet = getF3Sheet();
  var range = sheet.getRange('I3:I500');
  var rules = [
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Pending').setBackground('#fff3cd').setRanges([range]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Syncing').setBackground('#cce5ff').setRanges([range]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Synced').setBackground('#d4edda').setRanges([range]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Error').setBackground('#f8d7da').setRanges([range]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Partial').setBackground('#ffe5b4').setRanges([range]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Failed').setBackground('#d6d6d6').setRanges([range]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Skip').setBackground('#e0e0e0').setRanges([range]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Reversed').setBackground('#b3e5fc').setRanges([range]).build()
  ];
  sheet.setConditionalFormatRules(rules);
}

// ============================================
// TRIGGERS
// ============================================

function setupF3Trigger() {
  var triggers = ScriptApp.getProjectTriggers();
  var f3Function = 'pollKatanaStockTransfers';

  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === f3Function) {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }

  ScriptApp.newTrigger('pollKatanaStockTransfers').timeBased().everyMinutes(F3_CONFIG.POLL_INTERVAL_MINUTES).create();
}

function removeF3Trigger() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'pollKatanaStockTransfers') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}

// ============================================
// TEST & UTILITY FUNCTIONS
// ============================================

function setupF3Complete() {
  createF3Sheet();
  setupF3Formatting();
  setupF3Trigger();
}

function testF3Poll() {
  var result = pollKatanaStockTransfers();
  Logger.log(JSON.stringify(result, null, 2));
}

function resetF3SyncTimestamp() {
  CacheService.getScriptCache().remove(F3_CONFIG.CACHE_KEY_LAST_SYNC);
}

/**
 * Clear all F3 data and start fresh (keeps sheet structure)
 */
function resetF3Fresh() {
  var ss = SpreadsheetApp.openById(F3_CONFIG.SHEET_ID);

  // Clear StockTransfers (keep headers + summary row)
  var stSheet = ss.getSheetByName(F3_CONFIG.TAB_NAME);
  if (stSheet && stSheet.getLastRow() > 2) {
    stSheet.deleteRows(3, stSheet.getLastRow() - 2);
  }

  // Reset cache
  CacheService.getScriptCache().remove(F3_CONFIG.CACHE_KEY_LAST_SYNC);
}

function forceSyncST(stNumber) {
  stNumber = stNumber || 'ST-221 TEST DELETE ME';

  var sheet = getF3Sheet();
  var data = sheet.getDataRange().getValues();

  for (var i = 2; i < data.length; i++) {
    var num = data[i][F3_CONFIG.COLS.ST_NUMBER - 1];
    if (num === stNumber) {
      var rowNum = i + 1;
      sheet.getRange(rowNum, F3_CONFIG.COLS.WASP_STATUS).setValue('Pending');
      sheet.getRange(rowNum, F3_CONFIG.COLS.NOTES).setValue('');
      sheet.getRange(rowNum, F3_CONFIG.COLS.WASP_SYNCED).setValue('');
      break;
    }
  }

  CacheService.getScriptCache().remove(F3_CONFIG.CACHE_KEY_LAST_SYNC);
  var result = pollKatanaStockTransfers();
  Logger.log('Result: ' + JSON.stringify(result, null, 2));
}

function forceSyncST221() {
  forceSyncST('ST-221 TEST DELETE ME');
}

function forceSyncST248() {
  forceSyncST('ST-248');
}
