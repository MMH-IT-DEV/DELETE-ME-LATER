// ============================================
// 03_WaspCallouts.gs - WASP → KATANA HANDLERS
// ============================================
// Handles WASP Callout events that sync back to Katana
// UPDATED: Added batch grouping with 10-second window
// UPDATED: Multiple items removed together = one Katana SA
// UPDATED: Supports 20+ items per batch
// ============================================

// Batch window in milliseconds (10 seconds - captures rapid item removals)
// WASP fires callouts sequentially; 10s catches bulk operations of 5-8 items
var BATCH_WINDOW_MS = 10000;

// Notes prefixes written by our own scripts when they mutate WASP.
// If a WASP callout carries one of these notes, it is an internal sync echo
// and must not create a Katana stock adjustment.
var INTERNAL_WASP_NOTE_PREFIXES = [
  'Sheet push ',
  'Sheet DELETE ',
  'Batch push ',
  'Batch DELETE ',
  'Katana push ',
  'Katana adj push ',
  'Manual adjustment ',
  'Sync zero ',
  'Sync add ',
  'Katana sync ',
  'ROLLBACK:'
];

var F5_MANAGED_WASP_NOTE_PREFIXES = [
  'ShipStation shipment ',
  'Voided label',
  'SO Cancelled: '
];

function isInternalWaspAdjustmentNote_(notes) {
  var text = String(notes || '').trim();
  if (!text) return false;

  for (var i = 0; i < INTERNAL_WASP_NOTE_PREFIXES.length; i++) {
    if (text.indexOf(INTERNAL_WASP_NOTE_PREFIXES[i]) === 0) {
      return true;
    }
  }

  return false;
}

function isF5ManagedWaspAdjustmentEcho_(notes) {
  var text = String(notes || '').trim();
  if (!text) return false;

  for (var i = 0; i < F5_MANAGED_WASP_NOTE_PREFIXES.length; i++) {
    if (text.indexOf(F5_MANAGED_WASP_NOTE_PREFIXES[i]) === 0) {
      return true;
    }
  }

  return false;
}

/**
 * Handle WASP quantity_added callout
 * Adds to batch queue for grouped processing
 */
function handleWaspQuantityAdded(payload) {
  var sku = payload.ItemNumber || payload.AssetTag;
  var location = payload.LocationCode;
  var siteName = payload.SiteName || CONFIG.WASP_SITE;
  var quantity = parseFloat(payload.Quantity) || 0;
  var userNotes = payload.AssetDescription || payload.Notes || '';
  var payloadLot = payload.Lot || payload.LotNumber || payload.BatchNumber || '';
  var payloadExpiry = payload.DateCode || payload.ExpiryDate || payload.ExpDate || '';

  // Filter out unresolved WASP template variables (e.g., "{trans.Lot}", "{trans.DateCode}")
  if (payloadLot && payloadLot.indexOf('{') >= 0) payloadLot = '';
  if (payloadExpiry && payloadExpiry.indexOf('{') >= 0) payloadExpiry = '';

  var userName = payload.Assignee || payload.UserName || payload.User || payload.ModifiedBy || payload.PerformedBy || '';

  // Only suppress actual F5-managed echoes from SHOPIFY.
  // Manual WASP adjustments at SHOPIFY must still create Katana SAs.
  if (isF5ManagedWaspAdjustmentEcho_(userNotes)) {
    return { status: 'skipped', reason: 'Location ' + location + ' managed by F5' };
  }

  // Suppress internal sheet/script echoes even if the cache pre-mark expired
  // before WASP delivered the callout.
  if (isInternalWaspAdjustmentNote_(userNotes)) {
    return { status: 'skipped', reason: 'Internal sheet/script adjustment' };
  }

  // Check if recently synced — covers both Katana→WASP loops and engin sheet pre-marks
  if (wasRecentlySyncedToWasp(sku, location, 'add', siteName)) {
    return { status: 'skipped', reason: 'Recently synced by Katana handler or engin sheet' };
  }

  if (quantity === 0) {
    return { status: 'skipped', reason: 'zero quantity' };
  }

  // Add to batch queue
  var batchId = addToBatchQueue('add', {
    sku: sku,
    quantity: quantity,
    location: location,
    siteName: siteName,
    userNotes: userNotes,
    payloadLot: payloadLot,
    payloadExpiry: payloadExpiry,
    userName: userName
  });

  // Try to process batch after window closes
  var result = processBatchIfReady('add', batchId);

  return {
    status: result ? 'processed' : 'queued',
    batchId: batchId,
    sku: sku,
    quantity: quantity
  };
}

/**
 * Handle WASP quantity_removed callout
 * Adds to batch queue for grouped processing
 */
function handleWaspQuantityRemoved(payload) {
  var sku = payload.ItemNumber || payload.AssetTag;
  var quantity = parseFloat(payload.Quantity) || 0;
  var location = payload.LocationCode;
  var siteName = payload.SiteName || CONFIG.WASP_SITE;
  var userNotes = payload.AssetDescription || payload.Notes || '';
  var pickOrderNumber = payload.PickOrderNumber || payload.OrderNumber || '';
  var payloadLot = payload.Lot || payload.LotNumber || payload.BatchNumber || '';
  var payloadExpiry = payload.DateCode || payload.ExpiryDate || payload.ExpDate || '';

  // Filter out unresolved WASP template variables (e.g., "{trans.Lot}", "{trans.DateCode}")
  if (payloadLot && payloadLot.indexOf('{') >= 0) payloadLot = '';
  if (payloadExpiry && payloadExpiry.indexOf('{') >= 0) payloadExpiry = '';

  var userName = payload.Assignee || payload.UserName || payload.User || payload.ModifiedBy || payload.PerformedBy || '';

  // Only suppress actual F5-managed echoes from SHOPIFY.
  // Manual WASP adjustments at SHOPIFY must still create Katana SAs.
  if (isF5ManagedWaspAdjustmentEcho_(userNotes)) {
    return { status: 'skipped', reason: 'Location ' + location + ' managed by F5' };
  }

  // Suppress internal sheet/script echoes even if the cache pre-mark expired
  // before WASP delivered the callout.
  if (isInternalWaspAdjustmentNote_(userNotes)) {
    return { status: 'skipped', reason: 'Internal sheet/script adjustment' };
  }

  // Route pick orders separately (they have different handling)
  if (pickOrderNumber) {
    return handlePickRemoval(payload, sku, quantity, location, userNotes, pickOrderNumber);
  }

  // DEBUG: capture cache state at callout-arrival time so engin-src can read it back
  (function() {
    try {
      var _c = CacheService.getScriptCache();
      var _k1 = 'wasp_sync_remove_' + sku + '_' + location;
      var _k2 = 'wasp_sync_remove_' + sku;
      PropertiesService.getScriptProperties().setProperty('DEBUG_LAST_CALLOUT', JSON.stringify({
        t: new Date().toISOString(),
        sku: sku, loc: location,
        k1: _k1, k1val: _c.get(_k1),
        k2: _k2, k2val: _c.get(_k2),
        rawPayload: JSON.stringify(payload).substring(0, 600)
      }));
    } catch(e) {}
  })();

  // Check if recently synced — covers both Katana→WASP loops and engin sheet pre-marks
  if (wasRecentlySyncedToWasp(sku, location, 'remove', siteName)) {
    return { status: 'skipped', reason: 'Recently synced by Katana handler or engin sheet' };
  }

  // Add to batch queue
  var batchId = addToBatchQueue('remove', {
    sku: sku,
    quantity: quantity,
    location: location,
    siteName: siteName,
    userNotes: userNotes,
    payloadLot: payloadLot,
    payloadExpiry: payloadExpiry,
    userName: userName
  });

  // Try to process batch after window closes
  var result = processBatchIfReady('remove', batchId);

  return {
    status: result ? 'processed' : 'queued',
    batchId: batchId,
    sku: sku,
    quantity: quantity
  };
}

// ============================================
// KATANA ERROR PARSING
// ============================================

/**
 * Parse Katana API error response into a short readable message.
 * Katana returns JSON like: {"errors":[{"message":"..."}]} or {"error":"..."}
 */
function parseKatanaError(responseText) {
  if (!responseText) return '';
  // Normalise to string so we always have something to show
  var raw = typeof responseText === 'string' ? responseText : JSON.stringify(responseText);
  try {
    var parsed = JSON.parse(raw);
    // Katana v1 format: {errors: [{message: "..."}]}
    if (parsed.errors && parsed.errors.length > 0) {
      var msg = parsed.errors[0].message || parsed.errors[0].detail || '';
      if (msg.toLowerCase().indexOf('batch') >= 0) return 'Batch tracking required';
      if (msg.toLowerCase().indexOf('variant') >= 0 && msg.toLowerCase().indexOf('not found') >= 0) return 'Item not in Katana';
      if (msg.toLowerCase().indexOf('location') >= 0) return 'Location error';
      return msg.substring(0, 100);
    }
    // Katana v2 format: {error: "..."} or {error: {message: "..."}}
    if (parsed.error) {
      var errVal = parsed.error;
      if (typeof errVal === 'string') return errVal.substring(0, 100);
      return String(errVal.message || errVal.name || JSON.stringify(errVal)).substring(0, 100);
    }
    // Katana v3 format: {message: "..."}
    if (parsed.message) return String(parsed.message).substring(0, 100);
    // Unknown format — show raw so we can diagnose
    return raw.substring(0, 120);
  } catch (e) {
    // Not JSON — show raw text
    return raw.substring(0, 120);
  }
}

function buildF2BatchNote_(lot, expiry) {
  var parts = [];
  var lotText = String(lot || '').trim();
  var expiryText = formatExpiryDate(expiry || '');

  if (lotText) parts.push('lot:' + lotText);
  if (expiryText) parts.push('exp:' + expiryText);

  return parts.join('  ');
}

function buildF2ActivityDetails_(sku, quantity, lot, expiry) {
  return sku + ' x' + Math.abs(quantity);
}

// ============================================
// BATCH QUEUE FUNCTIONS
// ============================================

/**
 * Get or create batch ID using sliding window
 * Items join existing batch if within window, otherwise new batch created
 */
function getOrCreateBatchId(action) {
  var cache = CacheService.getScriptCache();
  var activeBatchKey = 'active_batch_' + action;
  var activeBatch = cache.get(activeBatchKey);

  if (activeBatch) {
    var batchData = JSON.parse(activeBatch);
    var elapsed = new Date().getTime() - batchData.created;

    // If within window, reuse batch
    if (elapsed < BATCH_WINDOW_MS) {
      return batchData.batchId;
    }
  }

  // Create new batch
  var newBatchId = 'BATCH-' + new Date().getTime();
  cache.put(activeBatchKey, JSON.stringify({
    batchId: newBatchId,
    created: new Date().getTime()
  }), 120);

  return newBatchId;
}

/**
 * Add item to batch queue
 * Uses sliding window - items arriving within 15 seconds of first item join same batch
 * Returns batch ID
 */
function addToBatchQueue(action, itemData) {
  var cache = CacheService.getScriptCache();
  var batchId = getOrCreateBatchId(action);
  var key = 'batch_' + action + '_' + batchId;
  var timeKey = key + '_time';

  // Get existing items or create new array
  var existingData = cache.get(key);
  var items = existingData ? JSON.parse(existingData) : [];

  // Add timestamp to item
  itemData.addedAt = new Date().toISOString();

  // Add new item
  items.push(itemData);

  // Store batch with 120 second expiry (enough time to process)
  cache.put(key, JSON.stringify(items), 120);

  // Store batch creation time (only first item sets this)
  if (!cache.get(timeKey)) {
    cache.put(timeKey, new Date().getTime().toString(), 120);
  }

  // Skip noisy per-item logging - only log when batch is processed

  return batchId;
}

/**
 * Process batch if window has closed.
 * Uses LockService.getScriptLock() to guarantee only ONE execution processes
 * each batch — prevents duplicate SAs when concurrent callouts arrive.
 */
function processBatchIfReady(action, batchId) {
  var cache = CacheService.getScriptCache();
  var key = 'batch_' + action + '_' + batchId;
  var timeKey = key + '_time';
  var processedKey = key + '_processed';

  // Fast-exit: already processed (no lock needed for this check)
  if (cache.get(processedKey)) {
    return false;
  }

  // Wait for batch window to close
  var batchTime = parseInt(cache.get(timeKey) || '0');
  var elapsed = new Date().getTime() - batchTime;
  if (elapsed < BATCH_WINDOW_MS) {
    Utilities.sleep(Math.min(BATCH_WINDOW_MS - elapsed + 500, 5000));
  }

  // Acquire script lock — only held for cache reads/writes, NOT during API calls.
  // Releasing before API calls prevents concurrent F4/F6 webhook handlers from
  // timing out on acquireExecutionGuard_ while F2 is making slow Katana API calls.
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000); // wait up to 10s; throws if unavailable
  } catch (e) {
    // Another execution holds the lock and presumably is processing this batch
    return false;
  }

  var items;
  try {
    // Re-check after acquiring lock — another execution may have just finished
    if (cache.get(processedKey)) {
      return false;
    }

    var itemsData = cache.get(key);
    if (!itemsData) {
      return false;
    }

    items = JSON.parse(itemsData);
    if (items.length === 0) {
      return false;
    }

    // Mark processed BEFORE releasing lock — prevents any re-entry
    cache.put(processedKey, 'true', 120);

  } catch (e) {
    logToSheet('BATCH_PROCESS_ERROR', { batchId: batchId, error: e.message }, {});
    return false;
  } finally {
    // Release lock here — before any Katana/WASP API calls.
    // All remaining work (SA creation, logActivity) runs outside the lock.
    lock.releaseLock();
  }

  // Outside the lock — slow Katana API calls and sheet writes happen here.
  // Deduplication is already guaranteed by processedKey being set above.
  try {
    if (action === 'add') {
      processBatchAdd(items, batchId);
    } else if (action === 'remove') {
      processBatchRemove(items, batchId);
    }
    return true;
  } catch (e) {
    logToSheet('BATCH_PROCESS_ERROR', { batchId: batchId, error: e.message }, {});
    return false;
  }
}

/**
 * Process a batch of stock additions
 * Creates one Katana SA with multiple rows
 */
function processBatchAdd(items, batchId) {
  // Group items by site (in case items come from different sites)
  var itemsBySite = {};
  for (var i = 0; i < items.length; i++) {
    var item = items[i];
    var site = item.siteName || CONFIG.WASP_SITE;
    if (!itemsBySite[site]) {
      itemsBySite[site] = [];
    }
    itemsBySite[site].push(item);
  }

  // Process each site separately
  for (var siteName in itemsBySite) {
    var siteItems = itemsBySite[siteName];
    processSiteBatchAdd(siteItems, siteName, batchId);
  }
}

/**
 * Process batch additions for a single site
 * Each item gets its own Katana SA — failures are isolated per item.
 */
function processSiteBatchAdd(items, siteName, batchId) {
  var katanaLocationName = getKatanaLocationForSite(siteName);
  if (!katanaLocationName) {
    logToSheet('F2_SITE_NOT_MAPPED', { site: siteName }, {});
    return;
  }

  var katanaLocation = getKatanaLocationByName(katanaLocationName);
  if (!katanaLocation) {
    logToSheet('F2_LOCATION_NOT_FOUND', { location: katanaLocationName }, {});
    return;
  }

  // One SA per item — a failure on one item does not affect others
  for (var i = 0; i < items.length; i++) {
    var item = items[i];

    var variant = getKatanaVariantBySku(item.sku);
    if (!variant) {
      logToSheet('F2_SKU_NOT_FOUND', { sku: item.sku }, {});
      logAdjustment('WASP', 'Add', item.userName || '', item.sku, '', siteName, item.location || '', item.payloadLot || '', item.payloadExpiry || '', item.quantity, '', null, 'ERROR');
      logActivity(
        'F2',
        joinActivitySegments_(['WASP Adjustment', buildF2ActivityDetails_(item.sku, item.quantity, item.payloadLot || '', item.payloadExpiry || '')]),
        'skipped',
        buildActivitySourceActionContext_('Wasp', 'adjust', siteName),
        [{
          sku: item.sku,
          qty: item.quantity,
          action: buildActivityCompactMeta_(siteName, item.location || '', item.payloadLot || '', item.payloadExpiry || '', ['not in Katana']),
          status: 'Skipped',
          success: false,
          qtyColor: 'grey'
        }],
        null
      );
      continue;
    }

    var lot = item.payloadLot;
    var expiry = item.payloadExpiry;

    if (!lot) {
      var lotInfo = getWaspLotInfo(siteName, item.sku, item.location);
      lot = lotInfo.lot || '';
      expiry = lotInfo.expiry || expiry;
    }

    var costPerUnit = parseFloat(variant.purchase_price || variant.default_purchase_price || variant.cost || 0);
    if (costPerUnit === 0) {
      // purchase_price not set — fall back to the item's current average cost in Katana
      costPerUnit = getKatanaAverageCost_(variant.id, katanaLocation.id);
    }

    var row = {
      variant_id: variant.id,
      quantity: item.quantity
    };

    if (costPerUnit > 0) {
      row.cost_per_unit = costPerUnit;
    }

    var batchIdKatana = null;
    if (lot) {
      var batchResultAdd = findBatchIdByNumber(variant.id, lot);
      batchIdKatana = batchResultAdd ? batchResultAdd.id : null;
      if (batchResultAdd && batchResultAdd.expiry && !expiry) {
        expiry = batchResultAdd.expiry;
      }
      if (batchIdKatana) {
        row.batch_transactions = [{
          batch_id: batchIdKatana,
          quantity: item.quantity
        }];
      }
    }

    var execId = getNextExecId(null);
    var saNumber = 'WASP Adjustment';
    var saReason = item.sku + ' +' + Math.abs(item.quantity);

    var result = createKatanaBatchStockAdjustment(
      [row],
      katanaLocation.id,
      '',
      execId,
      saNumber,
      saReason
    );

    var adjStatus = result.success ? 'OK' : 'ERROR';
    logAdjustment('WASP', 'Add', item.userName || '', item.sku, '', siteName, item.location || '', lot, expiry || '', item.quantity, result.success ? saNumber : '', result.success ? (result.saId || null) : null, adjStatus);

    // Log to Activity tab — link text "WASP Adjustment" links directly to the Katana SA page
    var saRef = result.saId ? ('SA-' + result.saId) : 'WASP Adjustment';
    var f2AddUrl = getKatanaWebUrl('sa', result.saId);
    var addActionNote = buildActivityCompactMeta_(siteName, item.location || '', lot, expiry);
    var itemUom = resolveVariantUom(variant);
    logActivity(
      'F2',
      joinActivitySegments_([saRef, buildF2ActivityDetails_(item.sku, item.quantity, lot, expiry)]),
      result.success ? 'success' : 'failed',
      buildActivitySourceActionContext_('Wasp', 'adjust', siteName),
      [{ sku: item.sku, qty: item.quantity, uom: itemUom, action: addActionNote, status: result.success ? 'Synced' : 'Failed', success: result.success, error: result.success ? '' : (result.response ? parseKatanaError(result.response) : 'Adjustment failed'), qtyColor: 'green' }],
      f2AddUrl ? { text: saRef, url: f2AddUrl } : null
    );

    // Save snapshot for SA revert — if this SA is deleted in Katana, handleSADeleted
    // will look up this snapshot and reverse the WASP quantity change.
    if (result.success && result.saId) {
      PropertiesService.getScriptProperties().setProperty(
        'sa_snapshot_' + result.saId,
        JSON.stringify({
          saId: result.saId,
          action: 'add',
          sku: item.sku,
          qty: item.quantity,
          uom: itemUom,
          location: item.location || '',
          site: siteName,
          lot: lot,
          expiry: expiry || ''
        })
      );
    }
  }

}

/**
 * Process a batch of stock removals
 * Creates one Katana SA with multiple rows
 */
function processBatchRemove(items, batchId) {
  // Group items by site
  var itemsBySite = {};
  for (var i = 0; i < items.length; i++) {
    var item = items[i];
    var site = item.siteName || CONFIG.WASP_SITE;
    if (!itemsBySite[site]) {
      itemsBySite[site] = [];
    }
    itemsBySite[site].push(item);
  }

  // Process each site separately
  for (var siteName in itemsBySite) {
    var siteItems = itemsBySite[siteName];
    processSiteBatchRemove(siteItems, siteName, batchId);
  }
}

/**
 * Process batch removals for a single site
 * Each item gets its own Katana SA — failures are isolated per item.
 */
function processSiteBatchRemove(items, siteName, batchId) {
  var katanaLocationName = getKatanaLocationForSite(siteName);
  if (!katanaLocationName) {
    logToSheet('F2_SITE_NOT_MAPPED', { site: siteName }, {});
    return;
  }

  var katanaLocation = getKatanaLocationByName(katanaLocationName);
  if (!katanaLocation) {
    logToSheet('F2_LOCATION_NOT_FOUND', { location: katanaLocationName }, {});
    return;
  }

  // One SA per item — a failure on one item does not affect others
  for (var i = 0; i < items.length; i++) {
    var item = items[i];

    var variant = getKatanaVariantBySku(item.sku);
    if (!variant) {
      logToSheet('F2_SKU_NOT_FOUND', { sku: item.sku }, {});
      logAdjustment('WASP', 'Remove', item.userName || '', item.sku, '', siteName, item.location || '', item.payloadLot || '', item.payloadExpiry || '', -Math.abs(item.quantity), '', null, 'ERROR');
      logActivity(
        'F2',
        joinActivitySegments_(['WASP Adjustment', buildF2ActivityDetails_(item.sku, Math.abs(item.quantity), item.payloadLot || '', item.payloadExpiry || '')]),
        'skipped',
        buildActivitySourceActionContext_('Wasp', 'adjust', siteName),
        [{
          sku: item.sku,
          qty: Math.abs(item.quantity),
          action: buildActivityCompactMeta_(siteName, item.location || '', item.payloadLot || '', item.payloadExpiry || '', ['not in Katana']),
          status: 'Skipped',
          success: false,
          qtyColor: 'grey'
        }],
        null
      );
      continue;
    }

    var lot = item.payloadLot;
    var lotFromPayload = !!lot;  // true = user explicitly sent a lot in the callout
    var expiry = item.payloadExpiry;

    if (!lot) {
      var lotInfo = getWaspLotInfo(siteName, item.sku, item.location);
      lot = lotInfo.lot || '';
      expiry = lotInfo.expiry || expiry;
    }

    var row = {
      variant_id: variant.id,
      quantity: -Math.abs(item.quantity)
      // cost_per_unit intentionally omitted for negative quantities —
      // Katana rejects cost_per_unit on removals (uses average cost automatically)
    };

    // Add batch_transactions if lot found.
    // Only skip the item if the lot was explicitly in the payload and the batch
    // doesn't exist in Katana — creating a SA without the correct batch would be wrong.
    // If the lot came from the fallback WASP lookup and isn't in Katana (item is not
    // lot-tracked in Katana), proceed without batch tracking.
    var batchIdKatana = null;
    if (lot) {
      var batchResultRem = findBatchIdByNumber(variant.id, lot);
      if (!batchResultRem) {
        if (lotFromPayload) {
          logToSheet('F2_BATCH_SKIP_NOT_FOUND', { sku: item.sku, lot: lot }, {});
          logAdjustment('WASP', 'Remove', item.userName || '', item.sku, '', siteName, item.location || '', lot, expiry || '', -Math.abs(item.quantity), '', null, 'ERROR');
          logActivity(
            'F2',
            joinActivitySegments_(['WASP Adjustment', buildF2ActivityDetails_(item.sku, Math.abs(item.quantity), lot, expiry || '')]),
            'skipped',
            buildActivitySourceActionContext_('Wasp', 'adjust', siteName),
            [{
              sku: item.sku,
              qty: Math.abs(item.quantity),
              action: buildActivityCompactMeta_(siteName, item.location || '', lot, expiry || '', ['lot not in Katana']),
              status: 'Skipped',
              success: false,
              qtyColor: 'grey'
            }],
            null
          );
          continue;
        }
        // Fallback lot not in Katana — not lot-tracked, proceed without batch
        logToSheet('F2_BATCH_FALLBACK_NO_KATANA', { sku: item.sku, lot: lot }, {});
        lot = '';
      } else {
        // Batch found — proceed. Katana will reject with a real error if stock is
        // truly insufficient; we don't pre-validate qty here.
        batchIdKatana = batchResultRem.id;
        if (batchResultRem.expiry && !expiry) {
          expiry = batchResultRem.expiry;
        }
        row.batch_transactions = [{
          batch_id: batchIdKatana,
          quantity: -Math.abs(item.quantity)
        }];
      }
    }

    var execId = getNextExecId(null);
    var saNumberR = 'WASP Adjustment';
    var saReasonR = item.sku + ' -' + Math.abs(item.quantity);

    var result = createKatanaBatchStockAdjustment(
      [row],
      katanaLocation.id,
      '',
      execId,
      saNumberR,
      saReasonR
    );

    var adjRemStatus = result.success ? 'OK' : 'ERROR';
    logAdjustment('WASP', 'Remove', item.userName || '', item.sku, '', siteName, item.location || '', lot, expiry || '', -Math.abs(item.quantity), result.success ? saNumberR : '', result.success ? (result.saId || null) : null, adjRemStatus);

    // Log to Activity tab — link text "WASP Adjustment" links directly to the Katana SA page
    var saRefR = result.saId ? ('SA-' + result.saId) : 'WASP Adjustment';
    var f2RemUrl = getKatanaWebUrl('sa', result.saId);
    var removeActionNote = buildActivityCompactMeta_(siteName, item.location || '', lot, expiry);
    var removeUom = resolveVariantUom(variant);
    logActivity(
      'F2',
      joinActivitySegments_([saRefR, buildF2ActivityDetails_(item.sku, Math.abs(item.quantity), lot, expiry)]),
      result.success ? 'success' : 'failed',
      buildActivitySourceActionContext_('Wasp', 'adjust', siteName),
      [{ sku: item.sku, qty: Math.abs(item.quantity), uom: removeUom, action: removeActionNote, status: result.success ? 'Synced' : 'Failed', success: result.success, error: result.success ? '' : (result.response ? parseKatanaError(result.response) : 'Adjustment failed'), qtyColor: 'red' }],
      f2RemUrl ? { text: saRefR, url: f2RemUrl } : null
    );

    // Save snapshot for SA revert — if this SA is deleted in Katana, handleSADeleted
    // will look up this snapshot and reverse the WASP quantity change.
    if (result.success && result.saId) {
      PropertiesService.getScriptProperties().setProperty(
        'sa_snapshot_' + result.saId,
        JSON.stringify({
          saId: result.saId,
          action: 'remove',
          sku: item.sku,
          qty: Math.abs(item.quantity),
          uom: removeUom,
          location: item.location || '',
          site: siteName,
          lot: lot,
          expiry: expiry || ''
        })
      );
    }
  }

}

// ============================================
// WASP MOVE HANDLER (location change — log only, no Katana SA)
// ============================================

/**
 * Handle WASP item_moved callout
 * Records location change in Adjustments Log — does NOT create a Katana SA.
 * Wire this up in doPost: event === 'item_moved' or 'quantity_moved'
 */
function handleWaspItemMoved(payload) {
  var sku = payload.ItemNumber || payload.AssetTag;
  var fromLocation = payload.FromLocationCode || payload.FromLocation || '';
  var toLocation = payload.ToLocationCode || payload.ToLocation || payload.LocationCode || '';
  var siteName = payload.SiteName || CONFIG.WASP_SITE;
  var quantity = parseFloat(payload.Quantity) || 0;
  var userName = payload.Assignee || payload.UserName || payload.User || payload.ModifiedBy || '';
  var payloadLot = payload.Lot || payload.LotNumber || payload.BatchNumber || '';
  if (payloadLot && payloadLot.indexOf('{') >= 0) payloadLot = '';

  if (!sku) {
    return { status: 'skipped', reason: 'no SKU' };
  }

  var moveLabel = (fromLocation && toLocation) ? fromLocation + '\u2192' + toLocation : (toLocation || fromLocation || siteName);

  logAdjustment('WASP', 'Move', userName, sku, '', siteName, moveLabel, payloadLot, '', quantity, '', null, 'OK');

  return { status: 'logged', sku: sku, move: moveLabel };
}

// ============================================
// WASP LOT LOOKUP
// ============================================

/**
 * Query WASP for lot/expiry info for an item
 * Returns { lot, expiry } or empty strings if not found
 */
function getWaspLotInfo(siteName, itemNumber, locationCode) {
  try {
    var url = CONFIG.WASP_BASE_URL + '/public-api/ic/item/inventorysearch';

    var payload = {
      ItemNumber: itemNumber,
      SiteName: siteName
    };

    if (locationCode) {
      payload.LocationCode = locationCode;
    }

    var result = waspApiCall(url, payload);

    if (!result.success) {
      return { lot: '', expiry: '' };
    }

    var response = JSON.parse(result.response);
    var data = response.Data || response.data || [];

    if (data.length === 0) {
      return { lot: '', expiry: '' };
    }

    // First try to find exact location match
    for (var i = 0; i < data.length; i++) {
      var rec = data[i];
      var recLocation = rec.LocationCode || rec.Location || '';

      // Match location if specified
      if (locationCode && recLocation !== locationCode) {
        continue;
      }

      var lot = rec.Lot || rec.LotNumber || rec.BatchNumber || '';
      var expiry = rec.DateCode || rec.ExpiryDate || rec.ExpDate || '';

      if (lot || expiry) {
        expiry = formatExpiryDate(expiry);
        return { lot: lot, expiry: expiry, quantity: rec.Quantity || 0 };
      }
    }

    // If no location match, try any record with lot
    for (var j = 0; j < data.length; j++) {
      var rec2 = data[j];
      var lot2 = rec2.Lot || rec2.LotNumber || rec2.BatchNumber || '';
      var expiry2 = rec2.DateCode || rec2.ExpiryDate || rec2.ExpDate || '';

      if (lot2) {
        expiry2 = formatExpiryDate(expiry2);
        return { lot: lot2, expiry: expiry2, quantity: rec2.Quantity || 0 };
      }
    }

    return { lot: '', expiry: '' };

  } catch (e) {
    // Silent fail - lot lookup is optional
    return { lot: '', expiry: '' };
  }
}

/**
 * Format expiry date to YYYY-MM-DD
 */
function formatExpiryDate(expiry) {
  return normalizeBusinessDate_(expiry);
}

// ============================================
// PICK ORDER HANDLING (Not batched)
// ============================================

/**
 * Handle removal from pick order
 * Marks Katana SO as delivered and triggers ShipStation
 */
function handlePickRemoval(payload, sku, quantity, location, notes, pickOrderNumber) {
  logToSheet('PICK_REMOVAL_DETECTED', {
    sku: sku,
    qty: quantity,
    location: location,
    pickOrderNumber: pickOrderNumber
  }, {});

  var katanaSoId = null;
  var orderNumber = pickOrderNumber;

  if (!orderNumber && notes) {
    var match = notes.match(/#(\d+)/);
    if (match) {
      orderNumber = '#' + match[1];
    }
  }

  if (orderNumber) {
    katanaSoId = getKatanaSoIdFromPickOrder(orderNumber);
    logToSheet('PICK_MAPPING_LOOKUP', {
      orderNumber: orderNumber,
      katanaSoId: katanaSoId
    }, {});

    if (pickOrderNumber) {
      var shipstationResult = handlePickCompleteShipStation(pickOrderNumber);

      if (shipstationResult.success) {
        logToSheet('PICK_COMPLETE_SUCCESS', {
          pickOrderNumber: pickOrderNumber
        }, {
          trackingNumber: shipstationResult.trackingNumber || 'N/A'
        });
      } else {
        logToSheet('PICK_COMPLETE_FAILED', {
          pickOrderNumber: pickOrderNumber
        }, shipstationResult);
      }
    }
  }

  if (katanaSoId) {
    var deliverResult = markKatanaSalesOrderDelivered(katanaSoId);

    if (deliverResult.success) {
      logToSheet('KATANA_SO_DELIVERED', {
        soId: katanaSoId,
        orderNumber: orderNumber
      }, deliverResult);

      clearPickOrderMapping(orderNumber);

      sendSlackNotification(
        'Order Complete!\n' +
        'Pick: ' + orderNumber + '\n' +
        'Katana SO delivered'
      );

      return { status: 'processed', katanaSoId: katanaSoId, delivered: true };
    } else {
      logToSheet('KATANA_DELIVER_FAILED', { soId: katanaSoId }, deliverResult);
      return { status: 'failed', error: 'Failed to deliver SO' };
    }
  }

  return { status: 'partial', message: 'WASP updated but Katana SO not found' };
}

// ============================================
// SYNC CACHE FUNCTIONS
// ============================================

/**
 * Pre-mark an engin-sheet WASP adjustment so F2 skips SA creation.
 * Called by the engin sheet BEFORE its WASP API call via HTTP POST to this webhook.
 * Uses the same cache as markSyncedToWasp() so wasRecentlySyncedToWasp() picks it up.
 *
 * Payload: { action: 'engin_mark', sku, location, op: 'add'|'remove' }
 */
function handleEnginMark(payload) {
  var sku = payload.sku || '';
  var location = payload.location || '';
  var siteName = payload.site || payload.siteName || payload.SiteName || '';
  var op = payload.op || 'add';
  if (!sku || !location) {
    return { status: 'error', reason: 'missing sku or location' };
  }
  markSyncedToWasp(sku, location, op, siteName);
  // DEBUG: record that pre-mark was accepted and keys were set
  try {
    var siteKey = buildWaspSyncSiteKey_(sku, siteName, op);
    PropertiesService.getScriptProperties().setProperty('DEBUG_LAST_PREMARK', JSON.stringify({
      t: new Date().toISOString(), sku: sku, loc: location, site: siteName, op: op,
      k1: 'wasp_sync_' + op + '_' + sku + '_' + location,
      k2: 'wasp_sync_' + op + '_' + sku,
      k3: siteKey || ''
    }));
  } catch(e) {}
  return { status: 'ok', sku: sku, location: location, site: siteName, op: op };
}

function normalizeWaspSyncScopeToken_(value) {
  return String(value || '').trim().toLowerCase().replace(/[^a-z0-9]+/g, '_');
}

function buildWaspSyncSiteKey_(sku, siteName, action) {
  var siteToken = normalizeWaspSyncScopeToken_(siteName);
  if (!siteToken) return '';
  return 'wasp_sync_' + action + '_' + sku + '_site_' + siteToken;
}

function markSyncedToWasp(sku, location, action, siteName) {
  var cache = CacheService.getScriptCache();
  // Location-specific key (exact match when location is known)
  var key = 'wasp_sync_' + action + '_' + sku + '_' + location;
  cache.put(key, new Date().toISOString(), SYNC_CACHE_SECONDS);
  var siteKey = buildWaspSyncSiteKey_(sku, siteName, action);
  if (siteKey) {
    cache.put(siteKey, new Date().toISOString(), SYNC_CACHE_SECONDS);
  }
  // Location-agnostic fallback key — covers cases where the push engine targets
  // one location code but WASP fires the callout with a slightly different one.
  // Shorter TTL (60s) minimises the window for false-blocking a real WASP UI adjustment.
  var anyKey = 'wasp_sync_' + action + '_' + sku;
  cache.put(anyKey, new Date().toISOString(), 20); // Shorter — broader key, higher false-positive risk
}

function wasRecentlySyncedToWasp(sku, location, action, siteName) {
  var cache = CacheService.getScriptCache();
  var key    = 'wasp_sync_' + action + '_' + sku + '_' + location;
  var siteKey = buildWaspSyncSiteKey_(sku, siteName, action);
  var anyKey = 'wasp_sync_' + action + '_' + sku;
  return cache.get(key) !== null || (siteKey && cache.get(siteKey) !== null) || cache.get(anyKey) !== null;
}
