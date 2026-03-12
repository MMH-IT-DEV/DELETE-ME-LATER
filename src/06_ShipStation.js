// ============================================
// 06_ShipStation.gs - SHIPSTATION API
// ============================================
// ShipStation webhook processing and inventory synchronization
// Handles SHIP_NOTIFY webhooks and voided label polling
// ============================================

/**
 * Bundle SKU → component mapping
 * Auto-assembly products in Katana that don't exist as standalone items in WASP.
 * When shipped, deduct components instead of the bundle SKU.
 */
var BUNDLE_MAP = {
  'VB-UFC-DUO': [
    { sku: 'UFC-4OZ', qty: 2 },
    { sku: 'TTFC-1OZ', qty: 1 }
  ],
  'VB-CPT-DUO-4OZ': [
    { sku: 'B-PURPLE-GREEN-DUO', qty: 1 },
    { sku: 'TTFC-1OZ', qty: 1 }
  ],
  'VB-UHP': [
    { sku: 'UFC-4OZ', qty: 3 },
    { sku: 'TTFC-1OZ', qty: 1 },
    { sku: 'WH-8OZ-W', qty: 1 }
  ],
  'VB-UFC-4OZ-SUB': [
    { sku: 'UFC-4OZ', qty: 1 }
  ]
};

/**
 * Generic ShipStation API call
 */
function callShipStationAPI(endpoint, method, payload) {
  var url = CONFIG.SHIPSTATION_BASE_URL + endpoint;
  var credentials = Utilities.base64Encode(
    CONFIG.SHIPSTATION_API_KEY + ':' + CONFIG.SHIPSTATION_API_SECRET
  );

  var options = {
    method: method || 'GET',
    headers: {
      'Authorization': 'Basic ' + credentials,
      'Content-Type': 'application/json'
    },
    muteHttpExceptions: true
  };

  if (payload) {
    options.payload = JSON.stringify(payload);
  }

  var response = UrlFetchApp.fetch(url, options);
  return {
    code: response.getResponseCode(),
    data: JSON.parse(response.getContentText())
  };
}

// ============================================
// SHIP_NOTIFY WEBHOOK HANDLER
// ============================================

/**
 * Handle SHIP_NOTIFY webhook from ShipStation
 * @param {string} resourceUrl - Full ShipStation resource URL
 * @returns {Object} Processing summary
 */
function handleShipNotify(resourceUrl) {
  var execId = Utilities.getUuid();

  logToSheet('SS_SHIP_NOTIFY_RECEIVED', {
    execId: execId,
    resourceUrl: resourceUrl
  }, { timestamp: new Date().toISOString() });

  // Fetch shipment details from resource URL using Basic Auth
  var credentials = Utilities.base64Encode(
    CONFIG.SHIPSTATION_API_KEY + ':' + CONFIG.SHIPSTATION_API_SECRET
  );

  var options = {
    method: 'GET',
    headers: {
      'Authorization': 'Basic ' + credentials,
      'Content-Type': 'application/json'
    },
    muteHttpExceptions: true
  };

  // Append includeShipmentItems=true — ShipStation resource URL omits items by default
  var fetchUrl = String(resourceUrl).trim();
  // Remove any existing includeShipmentItems param (might be =false)
  fetchUrl = fetchUrl.replace(/[?&]includeShipmentItems=[^&]*/gi, '');
  // Clean up double-? if replace removed the first param
  fetchUrl = fetchUrl.replace(/\?&/, '?').replace(/\?$/, '');
  // Append fresh
  fetchUrl += (fetchUrl.indexOf('?') >= 0 ? '&' : '?') + 'includeShipmentItems=true';

  var response = UrlFetchApp.fetch(fetchUrl, options);
  var responseCode = response.getResponseCode();

  if (responseCode !== 200) {
    logToSheet('SS_SHIP_NOTIFY_FETCH_FAILED', {
      execId: execId,
      statusCode: responseCode
    }, { error: 'Failed to fetch resource URL' });
    return { success: false, error: 'Resource fetch failed with code ' + responseCode };
  }

  var responseText = response.getContentText();
  var responseData = JSON.parse(responseText);
  var shipments = responseData.shipments || [];

  // DEBUG: Log response structure to diagnose itemCount:0 issue
  if (shipments.length > 0) {
    var firstShip = shipments[0];
    var itemField = 'none';
    if (Array.isArray(firstShip.shipmentItems) && firstShip.shipmentItems.length > 0) itemField = 'shipmentItems[' + firstShip.shipmentItems.length + ']';
    else if (Array.isArray(firstShip.items) && firstShip.items.length > 0) itemField = 'items[' + firstShip.items.length + ']';
    else if (firstShip.shipmentItems === null) itemField = 'shipmentItems=null';
    else if (firstShip.shipmentItems === undefined) itemField = 'shipmentItems=undefined';
    logToSheet('SS_DEBUG_RESPONSE', {
      execId: execId
    }, {
      fetchUrl: fetchUrl.substring(0, 200),
      shipmentKeys: Object.keys(firstShip).sort().join(','),
      itemField: itemField,
      totalShipments: shipments.length
    });
  }

  if (shipments.length === 0) {
    logToSheet('SS_SHIP_NOTIFY_NO_SHIPMENTS', {
      execId: execId
    }, { warning: 'No shipments in resource response' });
    return { success: true, processed: 0 };
  }

  var results = {
    total: shipments.length,
    processed: 0,
    skipped: 0,
    failed: 0,
    errors: []
  };

  // Process each shipment
  for (var i = 0; i < shipments.length; i++) {
    var shipment = shipments[i];
    var shipmentResult = processShipment(shipment, execId);

    if (shipmentResult.skipped) {
      results.skipped++;
    } else if (shipmentResult.success) {
      results.processed++;
    } else {
      results.failed++;
      results.errors.push({
        shipmentId: shipment.shipmentId,
        error: shipmentResult.error
      });
    }
  }

  logToSheet('SS_SHIP_NOTIFY_COMPLETE', {
    execId: execId
  }, results);

  return results;
}

// ============================================
// SHIPMENT PROCESSING
// ============================================

/**
 * Process a single shipment
 * @param {Object} shipment - ShipStation shipment object
 * @param {string} execId - Execution ID for logging
 * @returns {Object} Processing result
 */
function processShipment(shipment, execId) {
  var shipmentId = shipment.shipmentId;
  var orderNumber = shipment.orderNumber;
  var shipTo = shipment.shipTo || {};
  var trackingNumber = shipment.trackingNumber || 'N/A';
  var shipmentItems = shipment.shipmentItems || shipment.items || shipment.orderItems || [];
  var shipDateText = typeof formatF5RecoveryDate_ === 'function'
    ? formatF5RecoveryDate_(shipment.shipDate || new Date())
    : Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  var ledgerLines = typeof buildF5ShipmentLedgerLines_ === 'function'
    ? buildF5ShipmentLedgerLines_(shipment, shipDateText)
    : [];

  // Deduplication check using CacheService
  var cache = CacheService.getScriptCache();
  var cacheKey = 'ss_shipped_' + shipmentId;
  var runningKey = cacheKey + '_running';
  var useLedgerRetryGuard = !!(typeof HOTFIX_FLAGS !== 'undefined' && HOTFIX_FLAGS.F5_LEDGER_RETRY_GUARD);

  if (cache.get(cacheKey)) {
    logToSheet('SS_SHIPMENT_DUPLICATE', {
      execId: execId,
      shipmentId: shipmentId,
      orderNumber: orderNumber
    }, { reason: 'Already processed' });
    return { success: true, skipped: true };
  }

  if (useLedgerRetryGuard && cache.get(runningKey)) {
    logToSheet('SS_SHIPMENT_IN_PROGRESS_SKIP', {
      execId: execId,
      shipmentId: shipmentId,
      orderNumber: orderNumber
    }, { reason: 'Shipment already processing' });
    return { success: true, skipped: true, inProgress: true };
  }

  if (useLedgerRetryGuard &&
      typeof getF5ShipmentLedgerEntryByShipmentId_ === 'function' &&
      typeof classifyF5ShipmentLedgerEntry_ === 'function') {
    var existingLedgerEntry = getF5ShipmentLedgerEntryByShipmentId_(shipmentId);
    if (existingLedgerEntry) {
      var ledgerClass = classifyF5ShipmentLedgerEntry_(existingLedgerEntry);
      if (ledgerClass.mode === 'handled') {
        cache.put(cacheKey, 'processed', 86400);
        logToSheet('SS_SHIPMENT_LEDGER_HANDLED_SKIP', {
          execId: execId,
          shipmentId: shipmentId,
          orderNumber: orderNumber
        }, { reason: 'Ledger already shows shipment handled' });
        return { success: true, skipped: true, handled: true };
      }
      if (ledgerClass.mode === 'manual') {
        cache.put(cacheKey, 'ledger_manual', 86400);
        logToSheet('SS_SHIPMENT_LEDGER_MANUAL_SKIP', {
          execId: execId,
          shipmentId: shipmentId,
          orderNumber: orderNumber
        }, {
          reason: 'Ledger already has failed lines; use F5 manual recovery',
          items: (typeof summarizeF5ShipmentLedgerLines_ === 'function')
            ? summarizeF5ShipmentLedgerLines_(ledgerClass.lines || [], 220)
            : ''
        });
        return { success: true, skipped: true, manual: true };
      }
      if (ledgerClass.mode === 'review') {
        logToSheet('SS_SHIPMENT_LEDGER_REVIEW_SKIP', {
          execId: execId,
          shipmentId: shipmentId,
          orderNumber: orderNumber
        }, { reason: ledgerClass.reason || 'Ledger requires review' });
        return { success: true, skipped: true, review: true };
      }
    }
  }

  if (typeof findF5ManualRecoveryRecord_ === 'function') {
    var recovery = findF5ManualRecoveryRecord_(null, shipmentId, orderNumber, shipDateText);
    if (recovery) {
      cache.put(cacheKey, 'manual_recovery', 86400);
      logToSheet('SS_SHIPMENT_MANUAL_RECOVERY_SKIP', {
        execId: execId,
        shipmentId: shipmentId,
        orderNumber: orderNumber
      }, { reason: 'Manually recovered on ' + (recovery.markedAt || recovery.reportDate || shipDateText) });
      return { success: true, skipped: true, manualRecovered: true };
    }
  }

  if (useLedgerRetryGuard) {
    cache.put(runningKey, '1', 900);
  }

  if (ledgerLines.length > 0 && typeof registerF5ShipmentLedgerPending_ === 'function') {
    registerF5ShipmentLedgerPending_(shipment, execId, ledgerLines);
  }

  if (!useLedgerRetryGuard) {
    cache.put(cacheKey, 'processed', 86400);
  }

  logToSheet('SS_SHIPMENT_PROCESSING', {
    execId: execId,
    shipmentId: shipmentId,
    orderNumber: orderNumber
  }, {
    shipTo: shipTo.name || 'Unknown',
    trackingNumber: trackingNumber,
    itemCount: shipmentItems.length
  });

  // Extract items and process inventory removal
  var itemResults = [];
  var ledgerResults = [];
  var shipmentUomCache = {};
  var successCount = 0;
  var failCount = 0;
  var ledgerIndex = 0;

  for (var i = 0; i < shipmentItems.length; i++) {
    var item = shipmentItems[i];
    var sku = item.sku || item.SKU || 'UNKNOWN';
    var quantity = item.quantity || item.Quantity || 0;

    if (!sku || sku === 'UNKNOWN' || quantity === 0) {
      itemResults.push({
        sku: sku,
        qty: quantity,
        success: false,
        action: 'SHOPIFY',
        status: 'Failed',
        error: 'Invalid SKU or quantity',
        qtyColor: 'red'
      });
      failCount++;
      continue;
    }

    // Expand bundle SKUs to components
    var itemsToProcess = [];
    if (BUNDLE_MAP[sku]) {
      var components = BUNDLE_MAP[sku];
      for (var b = 0; b < components.length; b++) {
        itemsToProcess.push({
          sku: components[b].sku,
          qty: components[b].qty * quantity,
          bundleSku: sku
        });
      }
    } else {
      itemsToProcess.push({ sku: sku, qty: quantity, bundleSku: null });
    }

    // Remove inventory from SHOPIFY location for each item/component
    for (var p = 0; p < itemsToProcess.length; p++) {
      var proc = itemsToProcess[p];
      var ledgerLine = ledgerLines[ledgerIndex] || null;
      ledgerIndex++;
      var previousLedgerStatus = ledgerLine ? String(ledgerLine.previousLedgerStatus || ledgerLine.ledgerStatus || '').trim() : '';
      var procUom = resolveSkuUom(proc.sku, shipmentUomCache);
      var notes = 'ShipStation shipment ' + shipmentId;
      if (proc.bundleSku) notes += ' (bundle: ' + proc.bundleSku + ')';

      if (useLedgerRetryGuard && ledgerLine && isF5LedgerHandledStatus_(previousLedgerStatus)) {
        itemResults.push({
          sku: proc.sku,
          qty: proc.qty,
          uom: procUom,
          ledgerKey: ledgerLine ? ledgerLine.lineKey : '',
          success: true,
          preserveLedgerStatus: true,
          action: (proc.bundleSku ? 'SHOPIFY  (bundle: ' + proc.bundleSku + ')' : 'SHOPIFY'),
          status: previousLedgerStatus === 'Voided' ? 'Returned' : 'Deducted',
          error: 'Already handled in ledger',
          qtyColor: 'grey'
        });
        ledgerResults.push(itemResults[itemResults.length - 1]);
        successCount++;
        continue;
      }

      var removeResult = waspRemoveInventory(
        proc.sku,
        proc.qty,
        'SHOPIFY',
        notes,
        'MMH Kelowna'
      );

      // Lot fallback: if -57041 (lot required), look up lot from WASP and retry
      var itemLot = '';
      var itemDateCode = '';
      if (!removeResult.success && removeResult.response && String(removeResult.response).indexOf('-57041') >= 0) {
        var lotData = waspLookupItemLotAndDate(proc.sku, 'SHOPIFY', 'MMH Kelowna');
        if (lotData && lotData.lot) {
          itemLot = lotData.lot;
          itemDateCode = normalizeBusinessDate_(lotData.dateCode || '');
          removeResult = waspRemoveInventoryWithLot(proc.sku, proc.qty, 'SHOPIFY',
            lotData.lot, notes, 'MMH Kelowna', lotData.dateCode || '');
        }
      }

      var parsedRemoveError = '';
      if (!removeResult.success && typeof parseWaspError === 'function') {
        parsedRemoveError = parseWaspError(removeResult.response, 'Remove', proc.sku, 'SHOPIFY');
      }
      var removeError = removeResult.error || parsedRemoveError || 'No stock at SHOPIFY';
      if (parsedRemoveError && (!removeResult.error || /^remove failed$/i.test(String(removeResult.error)))) {
        removeError = parsedRemoveError;
      }

      itemResults.push({
        sku: proc.sku,
        qty: proc.qty,
        uom: procUom,
        ledgerKey: ledgerLine ? ledgerLine.lineKey : '',
        success: removeResult.success || false,
        action: (proc.bundleSku ? 'SHOPIFY  (bundle: ' + proc.bundleSku + ')' : 'SHOPIFY')
          + (itemLot ? '  lot:' + itemLot : '')
          + (itemDateCode ? '  exp:' + itemDateCode : ''),
        status: removeResult.success ? 'Deducted' : 'Failed',
        error: removeResult.success ? '' : removeError,
        qtyColor: 'red'
      });
      ledgerResults.push(itemResults[itemResults.length - 1]);

      if (removeResult.success) {
        successCount++;
      } else {
        failCount++;
      }
    }
  }

  // Determine overall status
  var status = 'success';
  if (failCount > 0 && successCount === 0) {
    status = 'failed';
  } else if (failCount > 0) {
    status = 'partial';
  }

  if (ledgerLines.length > 0 && typeof applyF5ShipmentLedgerResults_ === 'function') {
    applyF5ShipmentLedgerResults_(ledgerLines, ledgerResults, execId);
  }

  if (useLedgerRetryGuard) {
    if (failCount === 0) {
      cache.put(cacheKey, 'processed', 86400);
    } else {
      cache.remove(cacheKey);
    }
    cache.remove(runningKey);
  }

  // Prepare Activity log header
  var itemCount = shipmentItems.length;
  var headerDetails;
  if (itemCount === 1 && itemResults.length === 1) {
    headerDetails = '#' + orderNumber + '  ' + itemResults[0].sku + ' x' + itemResults[0].qty;
  } else {
    headerDetails = '#' + orderNumber + '  ' + itemCount + ' items';
  }
  if (failCount > 0) headerDetails += '  ' + failCount + ' error' + (failCount > 1 ? 's' : '');

  // Build Shopify search link
  var shopifyLink = 'https://admin.shopify.com/store/mymagichealer/orders?query=' + encodeURIComponent(orderNumber);

  // Log to Activity tab — capture exec ID for flow detail
  var ssExecId = logActivity('F5', headerDetails, status, '→ SHOPIFY @ MMH Kelowna', itemResults, {
    text: '#' + orderNumber,
    url: shopifyLink
  });

  // Log to F5 Detail tab — match Activity header format
  var f5FlowItems = [];
  for (var j = 0; j < itemResults.length; j++) {
    var itemResult = itemResults[j];
    f5FlowItems.push({
      sku: itemResult.sku,
      qty: itemResult.qty,
      uom: itemResult.uom || '',
      detail: itemResult.action || '',
      status: itemResult.success ? 'Deducted' : 'Failed',
      error: itemResult.error || '',
      qtyColor: 'red'
    });
  }
  logFlowDetail('F5', ssExecId, {
    ref: '#' + orderNumber,
    detail: headerDetails + '  → SHOPIFY @ MMH Kelowna',
    status: status === 'success' ? 'Shipped' : status === 'failed' ? 'Failed' : 'Partial',
    linkText: '#' + orderNumber,
    linkUrl: shopifyLink
  }, f5FlowItems);

  logToSheet('SS_SHIPMENT_COMPLETE', {
    execId: execId,
    shipmentId: shipmentId,
    orderNumber: orderNumber
  }, {
    status: status,
    successCount: successCount,
    failCount: failCount
  });

  return {
    success: status !== 'failed',
    shipmentId: shipmentId,
    orderNumber: orderNumber,
    itemResults: itemResults
  };
}

function fetchShipStationShipmentsPaged_(endpointBase, pageSize) {
  pageSize = Number(pageSize || 500) || 500;
  var page = 1;
  var allShipments = [];

  while (true) {
    var endpoint = endpointBase +
      (endpointBase.indexOf('?') >= 0 ? '&' : '?') +
      'pageSize=' + pageSize + '&page=' + page;
    var result = callShipStationAPI(endpoint, 'GET');

    if (result.code !== 200) {
      return {
        success: false,
        code: result.code,
        shipments: allShipments
      };
    }

    var data = result.data || {};
    var batch = data.shipments || [];
    for (var i = 0; i < batch.length; i++) {
      allShipments.push(batch[i]);
    }

    var pages = Number(data.pages || 1) || 1;
    if (page >= pages || batch.length < pageSize) {
      break;
    }
    page++;
  }

  return {
    success: true,
    code: 200,
    shipments: allShipments
  };
}

// ============================================
// VOIDED LABEL POLLING
// ============================================

/**
 * Poll ShipStation for voided labels
 * Runs on 5-minute timer trigger
 */
function pollVoidedLabels() {
  var manualRecoveryPause = typeof getF5ManualRecoveryModeInfo_ === 'function'
    ? getF5ManualRecoveryModeInfo_()
    : { active: false };
  if (manualRecoveryPause.active) {
    logToSheet('SS_VOID_POLL_PAUSED', {
      reportDate: manualRecoveryPause.reportDate || '',
      startedAt: manualRecoveryPause.startedAt || ''
    }, { reason: manualRecoveryPause.reason || 'F5 manual recovery active' });
    return { success: true, skipped: true, paused: true };
  }

  var execId = Utilities.getUuid();
  var scriptProps = PropertiesService.getScriptProperties();

  // Get last poll timestamp
  var lastPollStr = scriptProps.getProperty('SS_VOID_POLL_LAST');
  var lastPoll;

  if (lastPollStr) {
    lastPoll = new Date(lastPollStr);
  } else {
    // Default to 1 hour ago if never polled
    lastPoll = new Date(Date.now() - 3600000);
  }

  var lastPollISO = lastPoll.toISOString();

  logToSheet('SS_VOID_POLL_START', {
    execId: execId,
    lastPoll: lastPollISO
  }, { timestamp: new Date().toISOString() });

  // Call ShipStation API for voided shipments
  var result = fetchShipStationShipmentsPaged_(
    '/shipments?voidDateStart=' + encodeURIComponent(lastPollISO) +
    '&includeShipmentItems=true'
  );

  if (!result.success) {
    logToSheet('SS_VOID_POLL_FAILED', {
      execId: execId,
      statusCode: result.code
    }, { error: 'API call failed' });
    return { success: false, error: 'API call failed with code ' + result.code };
  }

  var shipments = result.shipments || [];
  var voidedShipments = [];

  // Filter for shipments with voidDate
  for (var i = 0; i < shipments.length; i++) {
    if (shipments[i].voidDate) {
      voidedShipments.push(shipments[i]);
    }
  }

  if (voidedShipments.length === 0) {
    scriptProps.setProperty('SS_VOID_POLL_LAST', new Date().toISOString());
    logToSheet('SS_VOID_POLL_COMPLETE', {
      execId: execId
    }, { voidedCount: 0 });
    return { success: true, processed: 0 };
  }

  // Process each voided shipment
  var processedCount = 0;
  var skippedCount = 0;
  var failedCount = 0;

  for (var j = 0; j < voidedShipments.length; j++) {
    manualRecoveryPause = typeof getF5ManualRecoveryModeInfo_ === 'function'
      ? getF5ManualRecoveryModeInfo_()
      : { active: false };
    if (manualRecoveryPause.active) {
      logToSheet('SS_VOID_POLL_PAUSED', {
        execId: execId,
        reportDate: manualRecoveryPause.reportDate || '',
        startedAt: manualRecoveryPause.startedAt || ''
      }, {
        reason: (manualRecoveryPause.reason || 'F5 manual recovery active') +
          ' | processed=' + processedCount + ' skipped=' + skippedCount + ' failed=' + failedCount
      });
      return {
        success: true,
        paused: true,
        processed: processedCount,
        skipped: skippedCount,
        failed: failedCount
      };
    }

    var voidResult = processVoid(voidedShipments[j], execId);

    if (voidResult.skipped) {
      skippedCount++;
    } else if (voidResult.success) {
      processedCount++;
    } else {
      failedCount++;
    }
  }

  // Update last poll timestamp
  scriptProps.setProperty('SS_VOID_POLL_LAST', new Date().toISOString());

  var summary = {
    voidedCount: voidedShipments.length,
    processed: processedCount,
    skipped: skippedCount,
    failed: failedCount
  };

  logToSheet('SS_VOID_POLL_COMPLETE', {
    execId: execId
  }, summary);

  // Log to Debug tab if any voids found
  if (voidedShipments.length > 0) {
    logToSheet('DEBUG', {
      execId: execId,
      type: 'SS_VOID_SUMMARY'
    }, summary);
  }

  return { success: true, summary: summary };
}

// ============================================
// VOID PROCESSING
// ============================================

/**
 * Process a voided shipment
 * @param {Object} shipment - ShipStation shipment object with voidDate
 * @param {string} execId - Execution ID for logging
 * @returns {Object} Processing result
 */
function processVoid(shipment, execId) {
  var shipmentId = shipment.shipmentId;
  var orderNumber = shipment.orderNumber;
  var shipmentItems = shipment.shipmentItems || shipment.items || shipment.orderItems || [];
  var shipDateText = typeof formatF5RecoveryDate_ === 'function'
    ? formatF5RecoveryDate_(shipment.shipDate || shipment.createDate || shipment.shipmentDate || new Date())
    : Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  var ledgerLines = typeof buildF5ShipmentLedgerLines_ === 'function'
    ? buildF5ShipmentLedgerLines_(shipment, shipDateText)
    : [];

  // Deduplication check
  var cache = CacheService.getScriptCache();
  var cacheKey = 'ss_voided_' + shipmentId;
  var runningKey = cacheKey + '_running';
  var useLedgerRetryGuard = !!(typeof HOTFIX_FLAGS !== 'undefined' && HOTFIX_FLAGS.F5_LEDGER_RETRY_GUARD);

  if (cache.get(cacheKey)) {
    logToSheet('SS_VOID_DUPLICATE', {
      execId: execId,
      shipmentId: shipmentId,
      orderNumber: orderNumber
    }, { reason: 'Already processed' });
    return { success: true, skipped: true };
  }

  if (useLedgerRetryGuard && cache.get(runningKey)) {
    logToSheet('SS_VOID_IN_PROGRESS_SKIP', {
      execId: execId,
      shipmentId: shipmentId,
      orderNumber: orderNumber
    }, { reason: 'Void already processing' });
    return { success: true, skipped: true, inProgress: true };
  }

  if (useLedgerRetryGuard && typeof hydrateF5ShipmentLedgerLines_ === 'function') {
    hydrateF5ShipmentLedgerLines_(ledgerLines);
    cache.put(runningKey, '1', 900);
  }

  if (!useLedgerRetryGuard) {
    cache.put(cacheKey, 'processed', 86400);
  }

  logToSheet('SS_VOID_PROCESSING', {
    execId: execId,
    shipmentId: shipmentId,
    orderNumber: orderNumber
  }, {
    voidDate: shipment.voidDate,
    itemCount: shipmentItems.length
  });

  // Process each item - add inventory back
  var voidedItems = [];
  var reviewOnlyVoid = false;
  var voidFailures = 0;
  var ledgerIndex = 0;

  for (var i = 0; i < shipmentItems.length; i++) {
    var item = shipmentItems[i];
    var sku = item.sku || item.SKU || 'UNKNOWN';
    var quantity = item.quantity || item.Quantity || 0;

    if (!sku || sku === 'UNKNOWN' || quantity === 0) {
      continue;
    }

    // Expand bundle SKUs to components (same as ship flow)
    var returnItems = [];
    if (BUNDLE_MAP[sku]) {
      var components = BUNDLE_MAP[sku];
      for (var b = 0; b < components.length; b++) {
        returnItems.push({ sku: components[b].sku, qty: components[b].qty * quantity });
      }
    } else {
      returnItems.push({ sku: sku, qty: quantity });
    }

    for (var r = 0; r < returnItems.length; r++) {
      var ret = returnItems[r];
      var ledgerLine = ledgerLines[ledgerIndex] || null;
      ledgerIndex++;
      var previousLedgerStatus = ledgerLine ? String(ledgerLine.previousLedgerStatus || ledgerLine.ledgerStatus || '').trim() : '';

      if (useLedgerRetryGuard && ledgerLine && previousLedgerStatus === 'Voided') {
        voidedItems.push({
          sku: ret.sku,
          qty: ret.qty,
          success: true,
          preserveLedgerStatus: true,
          error: 'Already voided in ledger',
          action: 'SHOPIFY',
          status: 'Voided',
          qtyColor: 'grey'
        });
        continue;
      }

      if (useLedgerRetryGuard && ledgerLine && previousLedgerStatus === 'Pending') {
        reviewOnlyVoid = true;
        voidedItems.push({
          sku: ret.sku,
          qty: ret.qty,
          success: false,
          preserveLedgerStatus: true,
          error: 'Ledger pending — manual review required',
          action: 'SHOPIFY',
          status: 'Review',
          qtyColor: 'grey'
        });
        continue;
      }

      if (useLedgerRetryGuard && ledgerLine && previousLedgerStatus && !isF5LedgerVoidEligibleStatus_(previousLedgerStatus)) {
        voidedItems.push({
          sku: ret.sku,
          qty: ret.qty,
          success: true,
          preserveLedgerStatus: true,
          error: 'Not previously deducted',
          action: 'SHOPIFY',
          status: 'Skipped',
          qtyColor: 'grey'
        });
        continue;
      }

      var addResult = waspAddInventory(
        ret.sku,
        ret.qty,
        'SHOPIFY',
        'Voided label — returned to stock. ShipmentId: ' + shipmentId,
        'MMH Kelowna'
      );

      voidedItems.push({
        sku: ret.sku,
        qty: ret.qty,
        preserveLedgerStatus: false,
        success: addResult.success || false,
        error: addResult.error || null,
        action: 'SHOPIFY',
        status: 'Voided',
        qtyColor: 'green'
      });
      if (!addResult.success) voidFailures++;
    }
  }

  if (useLedgerRetryGuard && ledgerLines.length > 0 && typeof applyF5ShipmentLedgerVoidResults_ === 'function') {
    applyF5ShipmentLedgerVoidResults_(ledgerLines, voidedItems, execId);
  }

  if (!reviewOnlyVoid && voidFailures === 0) {
    // Update original Activity row in-place (status → Voided, qty colors → green)
    var activityResult = updateActivityRow(orderNumber, voidedItems);
    if (!activityResult.success) {
      logToSheet('SS_VOID_ACTIVITY_UPDATE_MISS', {
        execId: execId, orderNumber: orderNumber
      }, { error: activityResult.error });
    }

    // Update original F5 Detail row in-place (status → Voided, sub-items → Returned, qty → green)
    var f5Result = updateFlowDetailRow('F5', orderNumber, 'Voided', 'Returned', 'green');
    if (!f5Result.success) {
      logToSheet('SS_VOID_F5_UPDATE_MISS', {
        execId: execId, orderNumber: orderNumber
      }, { error: f5Result.error });
    }
    if (useLedgerRetryGuard) {
      cache.put(cacheKey, 'processed', 86400);
    }
  } else {
    if (useLedgerRetryGuard) {
      logToSheet('SS_VOID_LEDGER_REVIEW', {
        execId: execId,
        shipmentId: shipmentId,
        orderNumber: orderNumber
      }, {
        reviewOnly: reviewOnlyVoid,
        failedAdds: voidFailures
      });
      cache.remove(cacheKey);
    }
  }
  if (useLedgerRetryGuard) cache.remove(runningKey);

  logToSheet('SS_VOID_COMPLETE', {
    execId: execId,
    shipmentId: shipmentId,
    orderNumber: orderNumber
  }, {
    itemsReturned: voidedItems.length
  });

  return {
    success: !reviewOnlyVoid && voidFailures === 0,
    shipmentId: shipmentId,
    orderNumber: orderNumber,
    voidedItems: voidedItems
  };
}

// ============================================
// TRIGGER SETUP
// ============================================

/**
 * Setup time-based trigger for void polling
 * Run this once to enable automatic void polling
 */
function setupVoidPollTrigger() {
  // Delete any existing triggers for pollVoidedLabels
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'pollVoidedLabels') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }

  // Create new 5-minute trigger
  ScriptApp.newTrigger('pollVoidedLabels')
    .timeBased()
    .everyMinutes(5)
    .create();

  logToSheet('SS_TRIGGER_SETUP', {}, {
    function: 'pollVoidedLabels',
    interval: '5 minutes',
    timestamp: new Date().toISOString()
  });

  return { success: true, message: 'Void poll trigger created' };
}

// ============================================
// SHIPMENT POLLING — processWebhookQueue
// ============================================
// Polls ShipStation every minute for new shipments.
// Called by time-driven trigger (every 1 minute).
// Uses SS_SHIP_POLL_LAST script property to track last processed time.
// Skips voided shipments — those are handled by pollVoidedLabels().
// Holds script lock to prevent concurrent runs.
// ============================================

/**
 * Poll ShipStation for new shipments since last run.
 * Triggered every minute by time-driven trigger.
 */
function processWebhookQueue() {
  var manualRecoveryPause = typeof getF5ManualRecoveryModeInfo_ === 'function'
    ? getF5ManualRecoveryModeInfo_()
    : { active: false };
  if (manualRecoveryPause.active) {
    logToSheet('SS_SHIP_POLL_PAUSED', {
      reportDate: manualRecoveryPause.reportDate || '',
      startedAt: manualRecoveryPause.startedAt || ''
    }, { reason: manualRecoveryPause.reason || 'F5 manual recovery active' });
    return { success: true, skipped: true, paused: true };
  }
  // Use CacheService flag (not LockService) so this time-driven trigger
  // does NOT block LockService.getScriptLock() used by processBatchIfReady
  // in 03_WaspCallouts.js — previously this caused WASP callout batches to
  // get stuck when both ran simultaneously.
  var pollCache = CacheService.getScriptCache();
  if (pollCache.get('ss_poll_running')) {
    Logger.log('processWebhookQueue: already running, skipping');
    return;
  }
  pollCache.put('ss_poll_running', '1', 60); // 60s safety TTL

  try {
    var execId = Utilities.getUuid();
    var scriptProps = PropertiesService.getScriptProperties();

    // Get last poll timestamp — default to 2 hours ago on first run
    var lastPollStr = scriptProps.getProperty('SS_SHIP_POLL_LAST');
    var lastPoll = lastPollStr ? new Date(lastPollStr) : new Date(Date.now() - 7200000);
    var lastPollISO = lastPoll.toISOString();

    logToSheet('SS_SHIP_POLL_START', { execId: execId, lastPoll: lastPollISO }, {});

    // ShipStation shipDateStart only accepts YYYY-MM-DD — full ISO datetime returns 0 results
    var shipDateParam = lastPollISO.substring(0, 10);
    var result = fetchShipStationShipmentsPaged_(
      '/shipments?shipDateStart=' + encodeURIComponent(shipDateParam) +
      '&includeShipmentItems=true&sortBy=ShipDate&sortDir=ASC'
    );

    if (!result.success) {
      logToSheet('SS_SHIP_POLL_FAILED', { execId: execId, statusCode: result.code }, { error: 'API call failed' });
      return;
    }

    // Advance timestamp immediately so a crash mid-loop doesn't reprocess on next run
    scriptProps.setProperty('SS_SHIP_POLL_LAST', new Date().toISOString());

    var shipments = result.shipments || [];

    // Skip voided shipments — pollVoidedLabels handles those
    var activeShipments = [];
    for (var i = 0; i < shipments.length; i++) {
      if (!shipments[i].voidDate) {
        activeShipments.push(shipments[i]);
      }
    }

    if (activeShipments.length === 0) {
      return;
    }

    var processed = 0;
    var skipped = 0;
    var failed = 0;

    for (var j = 0; j < activeShipments.length; j++) {
      manualRecoveryPause = typeof getF5ManualRecoveryModeInfo_ === 'function'
        ? getF5ManualRecoveryModeInfo_()
        : { active: false };
      if (manualRecoveryPause.active) {
        scriptProps.setProperty('SS_SHIP_POLL_LAST', lastPollISO);
        logToSheet('SS_SHIP_POLL_PAUSED', {
          execId: execId,
          reportDate: manualRecoveryPause.reportDate || '',
          startedAt: manualRecoveryPause.startedAt || ''
        }, {
          reason: (manualRecoveryPause.reason || 'F5 manual recovery active') +
            ' | processed=' + processed + ' skipped=' + skipped + ' failed=' + failed
        });
        return {
          success: true,
          paused: true,
          processed: processed,
          skipped: skipped,
          failed: failed
        };
      }

      var r = processShipment(activeShipments[j], execId);
      if (r.skipped) {
        skipped++;
      } else if (r.success) {
        processed++;
      } else {
        failed++;
      }
    }

    logToSheet('SS_SHIP_POLL_COMPLETE', { execId: execId }, {
      total: activeShipments.length,
      processed: processed,
      skipped: skipped,
      failed: failed
    });

  } finally {
    pollCache.remove('ss_poll_running');
  }
}

/**
 * TEST: Run processWebhookQueue against shipments from the last 2 hours.
 * Run this manually from the GAS editor — does NOT require a new deployment.
 * Check View → Logs and the Activity tab for results.
 */
function testProcessWebhookQueue() {
  var scriptProps = PropertiesService.getScriptProperties();
  var original = scriptProps.getProperty('SS_SHIP_POLL_LAST');

  Logger.log('=== testProcessWebhookQueue START ===');
  Logger.log('Original SS_SHIP_POLL_LAST: ' + (original || 'not set'));

  // Set last poll to 2 hours ago so recent shipments are included
  var testFrom = new Date(Date.now() - 7200000).toISOString();
  scriptProps.setProperty('SS_SHIP_POLL_LAST', testFrom);
  Logger.log('Polling shipments since: ' + testFrom);

  processWebhookQueue();

  Logger.log('=== testProcessWebhookQueue DONE ===');
  Logger.log('Check Activity tab for new F5 entries.');
}

/**
 * CHECK: Fetch all of today's shipments from ShipStation (all pages) and report
 * which ones have/haven't been processed by F5 (via dedup cache).
 * Read-only — no WASP calls, no inventory changes.
 * Run from GAS editor → View → Executions → Logs.
 */
function testF5TodaysCoverage() {
  Logger.log('=== F5 TODAY\'S COVERAGE CHECK ===');

  var credentials = Utilities.base64Encode(
    CONFIG.SHIPSTATION_API_KEY + ':' + CONFIG.SHIPSTATION_API_SECRET
  );
  var today = new Date().toISOString().substring(0, 10);
  var cache = CacheService.getScriptCache();

  var allShipments = [];
  var page = 1;
  var pageSize = 500;

  // Fetch all pages
  while (true) {
    var url = CONFIG.SHIPSTATION_BASE_URL +
      '/shipments?shipDateStart=' + today + '&pageSize=' + pageSize + '&page=' + page;
    var resp = UrlFetchApp.fetch(url, {
      method: 'GET',
      headers: { 'Authorization': 'Basic ' + credentials },
      muteHttpExceptions: true
    });
    if (resp.getResponseCode() !== 200) {
      Logger.log('[API] Error on page ' + page + ': HTTP ' + resp.getResponseCode());
      break;
    }
    var data = JSON.parse(resp.getContentText());
    var batch = data.shipments || [];
    for (var i = 0; i < batch.length; i++) {
      allShipments.push(batch[i]);
    }
    Logger.log('[API] Page ' + page + ' → ' + batch.length + ' shipments (total so far: ' + allShipments.length + ')');
    if (batch.length < pageSize) break; // last page
    page++;
  }

  Logger.log('[SS TOTAL] ' + allShipments.length + ' shipments on ' + today);

  var processed = [], missing = [], voided = [];

  for (var j = 0; j < allShipments.length; j++) {
    var s = allShipments[j];
    if (s.voidDate) {
      voided.push(s.orderNumber);
      continue;
    }
    var key = 'ss_shipped_' + s.shipmentId;
    if (cache.get(key)) {
      processed.push(s.orderNumber);
    } else {
      missing.push('#' + s.orderNumber + ' (shipmentId=' + s.shipmentId + ')');
    }
  }

  Logger.log('[PROCESSED] ' + processed.length + ' orders already in dedup cache ✓');
  Logger.log('[VOIDED]    ' + voided.length + ' voided shipments (skipped)');
  Logger.log('[MISSING]   ' + missing.length + ' orders NOT in cache:');
  if (missing.length === 0) {
    Logger.log('  ✓ All non-voided shipments have been processed!');
  } else {
    for (var k = 0; k < missing.length; k++) {
      Logger.log('  ✗ ' + missing[k]);
    }
  }

  Logger.log('=== COVERAGE CHECK COMPLETE ===');
}

/**
 * DIAGNOSTIC: Check F5 plumbing without touching any inventory.
 * Verifies: triggers running, SS_SHIP_POLL_LAST timestamp, webhook URL reachable.
 * Run from GAS editor → check View → Executions → Logs.
 */
function testF5Diagnostic() {
  Logger.log('=== F5 DIAGNOSTIC ===');

  // ── 1. Script properties ───────────────────────────────────────────────────
  var props = PropertiesService.getScriptProperties();
  var lastPollStr = props.getProperty('SS_SHIP_POLL_LAST');
  Logger.log('[PROPS] SS_SHIP_POLL_LAST = ' + (lastPollStr || 'NOT SET'));
  if (lastPollStr) {
    var lastPoll = new Date(lastPollStr);
    var minsAgo = Math.round((Date.now() - lastPoll.getTime()) / 60000);
    Logger.log('[PROPS] That was ' + minsAgo + ' minutes ago');
    if (minsAgo > 10) {
      Logger.log('[PROPS] ⚠ Last poll was >10 minutes ago — polling trigger may not be running');
    } else {
      Logger.log('[PROPS] ✓ Poll timestamp is recent');
    }
  } else {
    Logger.log('[PROPS] ⚠ SS_SHIP_POLL_LAST never set — processWebhookQueue has never run');
  }

  // ── 2. Triggers ────────────────────────────────────────────────────────────
  Logger.log('[TRIGGERS] Checking all project triggers...');
  var triggers = ScriptApp.getProjectTriggers();
  Logger.log('[TRIGGERS] Total triggers found: ' + triggers.length);
  var foundPoll = false;
  for (var i = 0; i < triggers.length; i++) {
    var t = triggers[i];
    var fn = t.getHandlerFunction();
    var type = t.getEventType();
    Logger.log('[TRIGGERS] ' + fn + ' | type=' + type);
    if (fn === 'processWebhookQueue') foundPoll = true;
  }
  if (foundPoll) {
    Logger.log('[TRIGGERS] ✓ processWebhookQueue trigger exists');
  } else {
    Logger.log('[TRIGGERS] ✗ processWebhookQueue trigger NOT FOUND — polling is not running');
  }

  // ── 3. ShipStation API — confirm order 93651 shipDate vs last poll ─────────
  var credentials = Utilities.base64Encode(
    CONFIG.SHIPSTATION_API_KEY + ':' + CONFIG.SHIPSTATION_API_SECRET
  );
  var response = UrlFetchApp.fetch(
    CONFIG.SHIPSTATION_BASE_URL + '/shipments?orderNumber=93651&includeShipmentItems=false',
    { method: 'GET', headers: { 'Authorization': 'Basic ' + credentials }, muteHttpExceptions: true }
  );
  if (response.getResponseCode() === 200) {
    var shipments = JSON.parse(response.getContentText()).shipments || [];
    if (shipments.length > 0) {
      var s = shipments[0];
      Logger.log('[SS] order 93651 → shipmentId=' + s.shipmentId + ' shipDate=' + s.shipDate + ' voided=' + s.voided);
      if (lastPollStr) {
        var shipTime = new Date(s.shipDate);
        var pollTime = new Date(lastPollStr);
        if (shipTime < pollTime) {
          Logger.log('[SS] ✓ shipDate is BEFORE last poll — polling should have caught it');
        } else {
          Logger.log('[SS] ⚠ shipDate is AFTER last poll — polling window missed this shipment');
        }
      }
    }
  } else {
    Logger.log('[SS] API error: HTTP ' + response.getResponseCode());
  }

  // ── 4. Poll query format test ──────────────────────────────────────────────
  // processWebhookQueue passes a full ISO datetime to shipDateStart,
  // but ShipStation's API requires YYYY-MM-DD format.
  // Test both and compare results.
  var isoTimestamp = lastPollStr || new Date(Date.now() - 7200000).toISOString();
  var dateOnly = isoTimestamp.substring(0, 10); // e.g. "2026-03-05"

  Logger.log('[QUERY TEST] Testing shipDateStart with ISO datetime: ' + isoTimestamp);
  var rIso = UrlFetchApp.fetch(
    CONFIG.SHIPSTATION_BASE_URL + '/shipments?shipDateStart=' + encodeURIComponent(isoTimestamp) + '&pageSize=5',
    { method: 'GET', headers: { 'Authorization': 'Basic ' + credentials }, muteHttpExceptions: true }
  );
  var isoCount = (JSON.parse(rIso.getContentText()).shipments || []).length;
  Logger.log('[QUERY TEST] ISO datetime → shipments returned: ' + isoCount);

  Logger.log('[QUERY TEST] Testing shipDateStart with date only: ' + dateOnly);
  var rDate = UrlFetchApp.fetch(
    CONFIG.SHIPSTATION_BASE_URL + '/shipments?shipDateStart=' + encodeURIComponent(dateOnly) + '&pageSize=5',
    { method: 'GET', headers: { 'Authorization': 'Basic ' + credentials }, muteHttpExceptions: true }
  );
  var dateCount = (JSON.parse(rDate.getContentText()).shipments || []).length;
  Logger.log('[QUERY TEST] Date only → shipments returned: ' + dateCount);

  if (isoCount === 0 && dateCount > 0) {
    Logger.log('[QUERY TEST] ✗ BUG CONFIRMED — ISO datetime format returns 0 results, date-only works');
    Logger.log('[QUERY TEST] processWebhookQueue is silently missing all shipments every poll');
  } else if (isoCount > 0) {
    Logger.log('[QUERY TEST] ✓ ISO datetime format works fine');
  } else {
    Logger.log('[QUERY TEST] Both returned 0 — no shipments today yet, or different issue');
  }

  // ── 5. Deployed web app URL ────────────────────────────────────────────────
  var webAppUrl = ScriptApp.getService().getUrl();
  Logger.log('[WEBAPP] Deployed URL: ' + (webAppUrl || 'NOT DEPLOYED'));

  Logger.log('=== DIAGNOSTIC COMPLETE ===');
}

/**
 * DRY-RUN TEST: Fetch a specific ShipStation shipment and log what F5 would do.
 * Does NOT call WASP — completely read-only.
 *
 * Usage:
 *   1. Set SHIPMENT_ID below to the shipment you want to inspect.
 *   2. Run testF5DryRun() from the GAS editor.
 *   3. Check View → Executions → Logs for full output.
 *
 * Output includes:
 *   - Whether the dedup cache already has this shipment (i.e. was it already processed?)
 *   - Raw item list from ShipStation
 *   - Bundle expansion (if any VB-* SKUs)
 *   - WASP actions that WOULD be called (SKU, qty, location) — but NOT executed
 */
function testF5DryRun() {
  var ORDER_NUMBER = '93651'; // ← Shopify/ShipStation order number

  Logger.log('=== F5 DRY RUN — orderNumber=' + ORDER_NUMBER + ' ===');

  // ── 1. Fetch shipment from ShipStation by order number ────────────────────
  var credentials = Utilities.base64Encode(
    CONFIG.SHIPSTATION_API_KEY + ':' + CONFIG.SHIPSTATION_API_SECRET
  );
  var fetchUrl = CONFIG.SHIPSTATION_BASE_URL +
    '/shipments?orderNumber=' + encodeURIComponent(ORDER_NUMBER) + '&includeShipmentItems=true';

  Logger.log('[API] Fetching: ' + fetchUrl);
  var response = UrlFetchApp.fetch(fetchUrl, {
    method: 'GET',
    headers: { 'Authorization': 'Basic ' + credentials },
    muteHttpExceptions: true
  });

  var code = response.getResponseCode();
  Logger.log('[API] HTTP ' + code);
  if (code !== 200) {
    Logger.log('[API] ERROR: ' + response.getContentText());
    return;
  }

  var data = JSON.parse(response.getContentText());
  var shipments = data.shipments || [];
  Logger.log('[API] shipments returned: ' + shipments.length);

  if (shipments.length === 0) {
    Logger.log('[API] No shipment found for orderNumber=' + ORDER_NUMBER);
    return;
  }

  // ── 2. Dedup cache check (uses real API shipmentId) ───────────────────────
  var cache = CacheService.getScriptCache();
  var shipment = shipments[0];
  var cacheKey = 'ss_shipped_' + shipment.shipmentId;
  var cached = cache.get(cacheKey);
  Logger.log('[CACHE] key=' + cacheKey + ' value=' + (cached || 'NOT FOUND'));
  if (cached) {
    Logger.log('[CACHE] ✓ Dedup cache HIT — F5 already processed this shipment within 24h.');
  } else {
    Logger.log('[CACHE] ✗ Not in cache — not yet processed (or cache expired >24h ago).');
  }

  var shipment = shipments[0];
  Logger.log('[SHIPMENT] orderNumber=' + shipment.orderNumber +
    ' | shipmentId=' + shipment.shipmentId +
    ' | shipDate=' + shipment.shipDate +
    ' | voided=' + (shipment.voided || false));

  var shipmentItems = shipment.shipmentItems || shipment.items || shipment.orderItems || [];
  Logger.log('[SHIPMENT] raw item count=' + shipmentItems.length);

  // ── 3. Log raw items + bundle expansion ────────────────────────────────────
  var wouldProcess = [];

  for (var i = 0; i < shipmentItems.length; i++) {
    var item = shipmentItems[i];
    var sku = item.sku || item.SKU || 'UNKNOWN';
    var qty = item.quantity || item.Quantity || 0;
    Logger.log('[ITEM ' + i + '] sku=' + sku + ' qty=' + qty);

    if (!sku || sku === 'UNKNOWN' || qty === 0) {
      Logger.log('[ITEM ' + i + '] SKIP — invalid sku or qty=0');
      continue;
    }

    if (BUNDLE_MAP[sku]) {
      Logger.log('[ITEM ' + i + '] BUNDLE — expanding ' + sku + ' x' + qty);
      var components = BUNDLE_MAP[sku];
      for (var b = 0; b < components.length; b++) {
        var expandedQty = components[b].qty * qty;
        Logger.log('  → component: ' + components[b].sku + ' x' + expandedQty);
        wouldProcess.push({ sku: components[b].sku, qty: expandedQty, bundleSku: sku });
      }
    } else {
      wouldProcess.push({ sku: sku, qty: qty, bundleSku: null });
    }
  }

  // ── 4. Log what WASP calls would be made ──────────────────────────────────
  Logger.log('[DRY-RUN] WASP actions that WOULD be called (not executed):');
  if (wouldProcess.length === 0) {
    Logger.log('  (none — no valid items found)');
  }
  for (var p = 0; p < wouldProcess.length; p++) {
    var proc = wouldProcess[p];
    var note = proc.bundleSku ? ' (bundle: ' + proc.bundleSku + ')' : '';
    Logger.log('  waspRemoveInventory(' + proc.sku + ', ' + proc.qty +
      ', "SHOPIFY", "ShipStation shipment ' + shipment.shipmentId + note + '", "MMH Kelowna")');
  }

  Logger.log('=== DRY RUN COMPLETE ===');
}

/**
 * Setup trigger for processWebhookQueue (every 1 minute).
 * Only needed if the trigger was accidentally deleted.
 * The existing trigger created before is still in place — skip if it shows in Triggers panel.
 */
function setupShipmentPollTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'processWebhookQueue') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  ScriptApp.newTrigger('processWebhookQueue')
    .timeBased()
    .everyMinutes(1)
    .create();
  Logger.log('processWebhookQueue trigger created (every 1 minute)');
  return { success: true };
}
