// ============================================
// 09_Utils.gs - UTILITIES & TESTS
// ============================================
// Slack notifications, test functions, debugging
// FIXED: Changed const/let to var for Google Apps Script compatibility
// ============================================

// ============================================
// UOM HELPER
// ============================================

/**
 * Normalise a unit-of-measure string for consistent display.
 * "ea", "each", "PC", "pcs" → "each"  (matches WASP EA label)
 * "g", "gram", "grams"      → "grams" (matches WASP grams label)
 * Everything else is returned as-is (trimmed).
 */
function normalizeUom(uom) {
  if (!uom) return '';
  var u = (uom + '').trim();
  var lc = u.toLowerCase();
  if (lc === 'ea' || lc === 'each' || lc === 'pc' || lc === 'pcs') return 'each';
  if (lc === 'g' || lc === 'gram' || lc === 'grams') return 'grams';
  return u;
}

/**
 * Extract UOM from an already-fetched variant object (no API calls).
 * Katana variants do not carry uom directly — this is a fast-path
 * check only; use resolveVariantUom() for reliable results.
 */
function getVariantUom(variant) {
  if (!variant) return '';
  return variant.uom || (variant.product ? (variant.product.uom || '') : '') || '';
}

/**
 * Resolve UOM for a Katana variant by fetching the parent product or material.
 * Katana stores UOM on products/materials, NOT on the variant itself.
 *   variant.product_id → GET products/{id}   → .uom
 *   variant.material_id → GET materials/{id} → .uom
 * Falls back to getVariantUom() (embedded fields) before making any API call.
 * Returns a normalised string ("pcs", "g", "kg", "mL", …) or "".
 */
function resolveVariantUom(variant) {
  if (!variant) return '';
  // Fast path — check embedded fields first (avoids extra API call)
  var fast = getVariantUom(variant);
  if (fast) return normalizeUom(fast);
  // Manufactured products
  if (variant.product_id) {
    var pData = fetchKatanaProduct(variant.product_id);
    var p = pData && pData.data ? pData.data : pData;
    return normalizeUom(p ? (p.uom || '') : '');
  }
  // Raw materials / purchased items
  if (variant.material_id) {
    var mData = fetchKatanaMaterial(variant.material_id);
    var m = mData && mData.data ? mData.data : mData;
    return normalizeUom(m ? (m.uom || '') : '');
  }
  return '';
}

/**
 * Resolve Katana UOM by SKU with optional per-call cache.
 * Returns a normalized display UOM or "" if the SKU cannot be resolved.
 */
function resolveSkuUom(sku, cache) {
  if (!sku) return '';
  if (cache && Object.prototype.hasOwnProperty.call(cache, sku)) {
    return cache[sku];
  }
  var variant = getKatanaVariantBySku(sku);
  var uom = resolveVariantUom(variant);
  if (cache) cache[sku] = uom;
  return uom;
}

// ============================================
// SLACK NOTIFICATIONS
// ============================================

/**
 * Send notification to Slack
 */
function sendSlackNotification(message) {
  if (!CONFIG.SLACK_WEBHOOK_URL) return;

  var payload = {
    text: '🔄 *Katana-WASP Sync*\n' + message
  };

  try {
    UrlFetchApp.fetch(CONFIG.SLACK_WEBHOOK_URL, {
      method: 'POST',
      contentType: 'application/json',
      payload: JSON.stringify(payload)
    });
  } catch (error) {
    Logger.log('Slack Error: ' + error.message);
  }
}

// ============================================
// MISTAKE SA CLEANUP
// ============================================

/**
 * Scan the Adjustments Log and list SA IDs that were created by mistake
 * (F5 ShipStation SHOPIFY removals that triggered F2 before the fix).
 *
 * Run this FIRST from the GAS editor — review the output in the execution log.
 * Then run deleteMistakeSAs() to actually remove them from Katana.
 */
function auditMistakeSAs() {
  var saIds = getMistakeSaIds_();
  if (saIds.length === 0) {
    Logger.log('AUDIT: No mistake SAs found (Location=SHOPIFY + Source=WASP) in Adjustments Log.');
    return;
  }
  Logger.log('AUDIT: Found ' + saIds.length + ' unique SA IDs to delete:');
  for (var i = 0; i < saIds.length; i++) {
    Logger.log('  SA ID: ' + saIds[i].id + '  |  SKUs: ' + saIds[i].skus + '  |  Time: ' + saIds[i].timestamp);
  }
  Logger.log('Review the list above, then run deleteMistakeSAs() to delete them from Katana.');
}

/**
 * Delete all mistake SAs from Katana (SHOPIFY-sourced F2 entries).
 * Run auditMistakeSAs() first to review what will be deleted.
 */
function deleteMistakeSAs() {
  var saIds = getMistakeSaIds_();
  if (saIds.length === 0) {
    Logger.log('DELETE: Nothing to delete — no mistake SAs found.');
    return;
  }
  Logger.log('DELETE: Deleting ' + saIds.length + ' SAs from Katana...');
  var deleted = 0;
  var failed = 0;
  for (var i = 0; i < saIds.length; i++) {
    var entry = saIds[i];
    var result = deleteKatanaSA(entry.id);
    if (result.success) {
      Logger.log('  DELETED SA ' + entry.id + ' (' + entry.skus + ')');
      deleted++;
    } else {
      Logger.log('  FAILED  SA ' + entry.id + ' — HTTP ' + result.code + ': ' + result.error);
      failed++;
    }
  }
  Logger.log('DELETE complete: ' + deleted + ' deleted, ' + failed + ' failed.');
}

/**
 * Internal: scan Adjustments Log for SHOPIFY + WASP rows, extract unique SA IDs.
 * @returns {Array} Array of { id, skus, timestamp }
 */
function getMistakeSaIds_() {
  var ss = SpreadsheetApp.openById(CONFIG.DEBUG_SHEET_ID);
  var sheet = ss.getSheetByName('F2 Adjustments');
  if (!sheet || sheet.getLastRow() < 2) return [];

  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 12).getValues();
  var formulas = sheet.getRange(2, 1, sheet.getLastRow() - 1, 12).getFormulas();

  // Columns (0-indexed): 0=Timestamp, 1=Source, 2=Action, 7=Location, 11=SA# (hyperlink)
  var seen = {};
  var results = [];

  for (var i = 0; i < data.length; i++) {
    var source   = String(data[i][1]).trim();
    var location = String(data[i][7]).trim();

    // Only rows where WASP triggered an adjustment from SHOPIFY location
    if (source !== 'WASP' || location.toUpperCase().indexOf('SHOPIFY') < 0) continue;

    // Extract numeric SA ID from hyperlink formula: =HYPERLINK(".../{id}","...")
    var formula = formulas[i][11] || '';
    var saId = null;
    if (formula) {
      var match = formula.match(/stockadjustment\/(\d+)/);
      if (match) saId = match[1];
    }
    if (!saId) continue;
    if (seen[saId]) {
      // Append SKU to existing entry
      seen[saId].skus += ', ' + String(data[i][4]).trim();
      continue;
    }

    seen[saId] = {
      id: saId,
      skus: String(data[i][4]).trim(),
      timestamp: String(data[i][0])
    };
    results.push(seen[saId]);
  }

  return results;
}

// ============================================
// TEST FUNCTIONS
// ============================================

/**
 * Diagnostic: inspect raw Katana PO row fields for a given PO ID.
 * Run manually from Apps Script editor — output appears in execution log.
 * Use this to confirm whether row.quantity is in purchase UoM or stock UoM.
 *
 * Usage: set PO_ID below to any recent PO that has a UoM conversion (e.g. EGG in dozens)
 */
function testPORowFields() {
  var PO_ID = 2596906; // PO-554 HIDDEN ON HALL FARM AND FEED (1 dozen EGG-X)

  Logger.log('=== PO ROW FIELD DIAGNOSTIC ===');
  Logger.log('Fetching PO header for ID: ' + PO_ID);

  var poData = fetchKatanaPO(PO_ID);
  Logger.log('PO header fields: ' + JSON.stringify(Object.keys(poData && poData.data ? poData.data : poData)));
  var po = poData && poData.data ? poData.data : poData;
  Logger.log('PO order_no: ' + (po ? po.order_no : 'n/a'));

  Logger.log('Fetching PO rows...');
  var poRowsData = fetchKatanaPORows(PO_ID);
  var rows = poRowsData && poRowsData.data ? poRowsData.data : (poRowsData || []);
  Logger.log('Row count: ' + rows.length);

  for (var i = 0; i < rows.length; i++) {
    var row = rows[i];
    Logger.log('--- Row ' + (i + 1) + ' ---');
    Logger.log('All fields: ' + JSON.stringify(row));
    Logger.log('quantity: ' + row.quantity);
    Logger.log('received_quantity: ' + row.received_quantity);
    Logger.log('unit_of_measure: ' + row.unit_of_measure);
    Logger.log('purchase_unit_of_measure: ' + row.purchase_unit_of_measure);
    Logger.log('unit_conversion_rate: ' + row.unit_conversion_rate);
    Logger.log('uom: ' + row.uom);
    Logger.log('conversion_rate: ' + row.conversion_rate);
  }
  Logger.log('=== END DIAGNOSTIC ===');
}

/**
 * Test WASP API connection
 */
function testWaspConnection() {
  var result = waspApiCall(
    CONFIG.WASP_BASE_URL + '/public-api/ic/item/infosearch',
    {}
  );
  Logger.log('WASP Test Result: ' + JSON.stringify(result));
  return result.success;
}

/**
 * Test ShipStation API connection
 */
function testShipStationConnection() {
  var result = callShipStationAPI('/orders?pageSize=1', 'GET');
  Logger.log('Status: ' + result.code);
  Logger.log('Connection: ' + (result.code === 200 ? 'SUCCESS' : 'FAILED'));
  return result.code === 200;
}

/**
 * Test sheet logging
 */
function testSheetLogging() {
  logToSheet('TEST_EVENT', {
    test: 'data',
    sku: 'TEST-SKU'
  }, { status: 'success' });
  Logger.log('Test log sent to sheet');
}

/**
 * Test finding a ShipStation order
 */
function testFindShipStationOrder() {
  var orderNumber = '89438';
  var result = findShipStationOrder(orderNumber);

  if (result.success) {
    Logger.log('Order found!');
    Logger.log('Order ID: ' + result.order.orderId);
    Logger.log('Status: ' + result.order.orderStatus);
  } else {
    Logger.log('Order not found');
  }

  return result;
}

/**
 * Test WASP pick order creation (dry run)
 */
function testWaspPickOrderCreate() {
  var testPayload = [{
    PickOrderNumber: 'TEST-' + new Date().getTime(),
    CustomerNumber: 'TEST-CUSTOMER',
    CustomerName: 'Test Customer',
    SiteName: CONFIG.WASP_SITE,
    OrderDate: new Date().toISOString().slice(0, 10),
    IssueOrder: true,
    Notes: 'Test pick order - can be deleted',
    PickOrderLines: [{
      ItemNumber: 'TEST-SKU',
      Quantity: 1,
      LocationCode: FLOWS.PICK_FROM_LOCATION
    }]
  }];

  logToSheet('TEST_PICK_ORDER', {}, testPayload);
  Logger.log('Test payload created (not sent)');
  Logger.log(JSON.stringify(testPayload, null, 2));

  return testPayload;
}

// ============================================
// DEBUG FUNCTIONS
// ============================================

/**
 * Debug a specific Sales Order
 */
function debugSalesOrder(orderNumber) {
  orderNumber = orderNumber || '#89410';

  Logger.log('=== DEBUGGING SALES ORDER: ' + orderNumber + ' ===');

  // Search for the SO
  var searchUrl = 'sales_orders?order_no=' + encodeURIComponent(orderNumber);
  var searchResult = katanaApiCall(searchUrl);
  Logger.log('Search Result: ' + JSON.stringify(searchResult, null, 2));

  if (!searchResult || !searchResult.data || searchResult.data.length === 0) {
    Logger.log('ERROR: No sales order found');
    return null;
  }

  var so = searchResult.data[0];
  var actualId = so.id;
  Logger.log('Found SO - ID: ' + actualId + ', Order#: ' + so.order_no);

  // Fetch SO details
  var soDetails = katanaApiCall('sales_orders/' + actualId);
  Logger.log('=== SO DETAILS ===');
  Logger.log(JSON.stringify(soDetails, null, 2));

  return {
    soId: actualId,
    orderNumber: so.order_no,
    customer: so.customer
  };
}

/**
 * Check script properties
 */
function checkScriptProperties() {
  var props = PropertiesService.getScriptProperties();
  var allProps = props.getProperties();

  Logger.log('=== Script Properties ===');
  for (var key in allProps) {
    // Mask sensitive values
    var value = (key.indexOf('TOKEN') >= 0 || key.indexOf('SECRET') >= 0 || key.indexOf('KEY') >= 0)
      ? '***HIDDEN***'
      : allProps[key];
    Logger.log(key + ' = ' + value);
  }
}

/**
 * Check configuration
 */
function checkConfig() {
  Logger.log('=== Configuration ===');
  Logger.log('WASP_BASE_URL: ' + CONFIG.WASP_BASE_URL);
  Logger.log('WASP_SITE: ' + CONFIG.WASP_SITE);
  Logger.log('WASP_TOKEN: ' + (CONFIG.WASP_TOKEN ? 'Set (' + CONFIG.WASP_TOKEN.length + ' chars)' : 'Missing'));
  Logger.log('KATANA_API_KEY: ' + (CONFIG.KATANA_API_KEY ? 'Set' : 'Missing'));
  Logger.log('SHIPSTATION_API_KEY: ' + (CONFIG.SHIPSTATION_API_KEY ? 'Set' : 'Missing'));
  Logger.log('SLACK_WEBHOOK_URL: ' + (CONFIG.SLACK_WEBHOOK_URL ? 'Set' : 'Not configured'));
  Logger.log('DEBUG_SHEET_ID: ' + CONFIG.DEBUG_SHEET_ID);
  Logger.log('');
  Logger.log('=== Flow Locations ===');
  Logger.log('PO Receiving: ' + FLOWS.PO_RECEIVING_LOCATION);
  Logger.log('MO Ingredients: ' + FLOWS.MO_INGREDIENT_LOCATION);
  Logger.log('MO Output: ' + FLOWS.MO_OUTPUT_LOCATION);
  Logger.log('Pick From: ' + FLOWS.PICK_FROM_LOCATION);
}

/**
 * Create PickMappings sheet if missing
 */
function createPickMappingsSheet() {
  var sheet = getPickMappingsSheet();
  Logger.log('Sheet name: ' + sheet.getName());
  Logger.log('Sheet URL: https://docs.google.com/spreadsheets/d/' +
             CONFIG.DEBUG_SHEET_ID + '/edit#gid=' + sheet.getSheetId());
}

/**
 * Print all locations for reference
 */
function printLocations() {
  Logger.log('=== WASP Locations ===');
  for (var key in LOCATIONS) {
    Logger.log(key + ': ' + LOCATIONS[key]);
  }
  Logger.log('');
  Logger.log('=== Flow Locations ===');
  for (var key in FLOWS) {
    Logger.log(key + ': ' + FLOWS[key]);
  }
}

// ============================================
// FLOW TAB SUMMARY FORMULAS
// ============================================
// Run once to add live stats formulas to Row 1 of each flow tab.
// MCP update_sheet uses setValues() which can't set formulas,
// so this must be run from the Apps Script editor.
// ============================================

/**
 * Set up summary formulas on all flow tabs (F1-F5)
 * Puts live stats in cell E1: items count, success, failed, total qty
 * Run once from Apps Script editor after deploying
 */
function setupFlowSummaryFormulas() {
  var ss = SpreadsheetApp.openById(CONFIG.DEBUG_SHEET_ID);

  // All flow tabs now have 6 columns: Exec ID | Time | Ref# | Details | Status (E) | Error (F)
  var tabs = [
    'F1 Receiving', 'F2 Adjustments', 'F3 Transfers', 'F4 Manufacturing', 'F5 Shipping'
  ];

  for (var i = 0; i < tabs.length; i++) {
    var sheet = ss.getSheetByName(tabs[i]);
    if (!sheet) {
      Logger.log('Tab not found: ' + tabs[i]);
      continue;
    }

    // Single formula in E1: "events | success | failed"
    // COUNTIFS with A<>"" ensures only header rows are counted (sub-items have blank A)
    var formula = '="Events: " & COUNTA(A4:A)'
      + ' & "  |  OK: " & COUNTIFS(A4:A,"<>",E4:E,"<>Failed",E4:E,"<>Error",E4:E,"<>Partial")'
      + ' & "  |  Errors: " & COUNTIFS(A4:A,"<>",E4:E,"Failed")'
      + ' & IF(COUNTIFS(A4:A,"<>",E4:E,"Partial")>0,"  |  Partial: " & COUNTIFS(A4:A,"<>",E4:E,"Partial"),"")';

    sheet.getRange('E1').setFormula(formula);
    sheet.getRange('E1').setFontSize(9).setFontColor('#666666').setFontWeight('normal');

    Logger.log('Set formula on ' + tabs[i] + ' E1');
  }

  Logger.log('All flow summary formulas set up');
}

// ============================================
// ONE-TIME SETUP — Run once then delete
// ============================================

/**
 * Store API keys in ScriptProperties (run once from script editor)
 * After running, verify with checkConfig() then delete this function
 */
function setupApiKeys() {
  // Keys already stored in ScriptProperties — do not re-add here.
  // To update a key: PropertiesService.getScriptProperties().setProperty('KEY_NAME', 'value')
  // To verify: run checkConfig()
  Logger.log('setupApiKeys(): Keys are already stored. Run checkConfig() to verify.');
}

// ============================================
// LOGGING SMOKE TEST — exercises all tabs
// ============================================
// Run from Apps Script editor to verify Activity, Debug,
// and all 5 flow tabs (F1-F5) are working correctly.
// No real API calls — just writes mock data to sheets.
// After running, check ALL tabs in CC tracker sheet.
// ============================================

/**
 * Smoke test ALL logging: Activity, Debug, F1-F5 flow tabs
 * Run from script editor → check every tab in CC tracker
 */
function testAllLogging() {
  Logger.log('=== LOGGING SMOKE TEST START ===');

  // --- F1 Receiving ---
  var f1ExecId = logActivity('F1', 'PO-999  2 items', 'success', 'to RECEIVING-DOCK', [
    { sku: 'TEST-MAT-A', qty: 50, success: true },
    { sku: 'TEST-MAT-B', qty: 25, success: true }
  ], { text: 'PO-999', url: getKatanaWebUrl('po', 99999) });

  logFlowDetail('F1', f1ExecId, {
    ref: 'PO-999', detail: '2 items to RECEIVING-DOCK @ MMH Kelowna', status: 'Complete',
    linkText: 'PO-999', linkUrl: getKatanaWebUrl('po', 99999)
  }, [
    { sku: 'TEST-MAT-A', qty: 50, detail: 'to RECEIVING-DOCK', status: 'Complete' },
    { sku: 'TEST-MAT-B', qty: 25, detail: 'to RECEIVING-DOCK', status: 'Complete' }
  ]);
  Logger.log('F1: OK — execId ' + f1ExecId);

  // --- F2 Adjustment (single) ---
  var f2ExecId = logActivity('F2', 'TEST-SKU-1 x10', 'success', 'WASP to Katana', [
    { sku: 'TEST-SKU-1', qty: 10, success: true }
  ]);

  logFlowDetail('F2', f2ExecId, {
    ref: 'SA', detail: 'TEST-SKU-1 x10  WASP to Katana @ MMH Kelowna', status: 'Synced'
  }, [
    { sku: 'TEST-SKU-1', qty: 10, detail: 'lot: LOT-77  WASP to Katana', status: 'Synced' }
  ]);
  Logger.log('F2 single: OK — execId ' + f2ExecId);

  // --- F2 Adjustment (batch with error) ---
  var f2bExecId = logActivity('F2', '3 items  1 error', 'partial', 'WASP to Katana', [
    { sku: 'TEST-SKU-2', qty: 5, success: true },
    { sku: 'TEST-SKU-3', qty: 8, success: true },
    { sku: 'TEST-SKU-4', qty: 3, success: false, error: 'Variant not found in Katana' }
  ]);

  logFlowDetail('F2', f2bExecId, {
    ref: 'SA', detail: '3 items  WASP to Katana @ MMH Kelowna', status: 'Partial',
    error: 'Variant not found in Katana'
  }, [
    { sku: 'TEST-SKU-2', qty: 5, detail: 'WASP to Katana', status: 'Synced' },
    { sku: 'TEST-SKU-3', qty: 8, detail: 'WASP to Katana', status: 'Synced' },
    { sku: 'TEST-SKU-4', qty: 3, detail: 'WASP to Katana', status: 'Failed', error: 'Variant not found in Katana' }
  ]);
  Logger.log('F2 batch: OK — execId ' + f2bExecId);

  // --- F3 Transfer ---
  var f3ExecId = logActivity('F3', 'ST-999  2 items', 'success', 'MMH Kelowna to Storage Warehouse', [
    { sku: 'TEST-PROD-1', qty: 100, success: true },
    { sku: 'TEST-PROD-2', qty: 200, success: true }
  ], { text: 'ST-999', url: getKatanaWebUrl('st', 88888) });

  logFlowDetail('F3', f3ExecId, {
    ref: 'ST-999', detail: '2 items  MMH Kelowna to Storage Warehouse', status: 'Synced',
    linkText: 'ST-999', linkUrl: getKatanaWebUrl('st', 88888)
  }, [
    { sku: 'TEST-PROD-1', qty: 100, detail: 'MMH Kelowna to Storage Warehouse (lot: BATCH-01)', status: 'Synced' },
    { sku: 'TEST-PROD-2', qty: 200, detail: 'MMH Kelowna to Storage Warehouse', status: 'Synced' }
  ]);
  Logger.log('F3: OK — execId ' + f3ExecId);

  // --- F4 Manufacturing ---
  var f4ExecId = logActivity('F4', 'MO-9999  TEST-FG x500  1 error', 'partial', 'PRODUCTION to PROD-RECEIVING', [
    { sku: 'TEST-FG', qty: 500, success: true, action: 'added to PROD-RECEIVING' },
    { sku: 'TEST-MAT-C', qty: 200, success: true, action: 'consumed from PRODUCTION' },
    { sku: 'TEST-MAT-D', qty: 100, success: false, error: 'Insufficient qty at PRODUCTION' }
  ], { text: 'MO-9999', url: getKatanaWebUrl('mo', 77777) });

  logFlowDetail('F4', f4ExecId, {
    ref: 'MO-9999', detail: 'TEST-FG x500 to PROD-RECEIVING', status: 'Partial',
    linkText: 'MO-9999', linkUrl: getKatanaWebUrl('mo', 77777)
  }, [
    { sku: 'TEST-MAT-C', qty: 200, detail: 'consumed from PRODUCTION', status: 'Complete' },
    { sku: 'TEST-MAT-D', qty: 100, detail: 'consumed from PRODUCTION', status: 'Failed', error: 'Insufficient qty at PRODUCTION' },
    { sku: 'TEST-FG', qty: 500, detail: 'produced to PROD-RECEIVING (batch: BATCH-99)', status: 'Complete' }
  ]);
  Logger.log('F4: OK — execId ' + f4ExecId);

  // --- F5 Shipping ---
  var f5ExecId = logActivity('F5', '#99999  2 items', 'success', 'to SHIPPING-DOCK', [
    { sku: 'TEST-SHIP-1', qty: 3, success: true },
    { sku: 'TEST-SHIP-2', qty: 1, success: true }
  ], { text: '#99999', url: getKatanaWebUrl('so', 66666) });

  logFlowDetail('F5', f5ExecId, {
    ref: '#99999', detail: '2 items to SHIPPING-DOCK', status: 'Shipped',
    linkText: '#99999', linkUrl: getKatanaWebUrl('so', 66666)
  }, [
    { sku: 'TEST-SHIP-1', qty: 3, detail: 'shipped', status: 'Shipped' },
    { sku: 'TEST-SHIP-2', qty: 1, detail: 'shipped', status: 'Shipped' }
  ]);
  Logger.log('F5: OK — execId ' + f5ExecId);

  // --- Debug (Logger.log only now) ---
  logToSheet('F2_SKU_NOT_FOUND', { sku: 'TEST-MISSING', SiteName: 'MMH Kelowna', execId: f2bExecId }, { error: 'No Katana variant matches SKU' });
  logToSheet('BATCH_LOCK_BUSY', { batchId: 'test-batch-123', execId: f2bExecId }, { message: 'Lock held by concurrent webhook' });
  Logger.log('Debug: OK — 2 events logged (Logger.log only)');

  Logger.log('=== LOGGING SMOKE TEST COMPLETE ===');
  Logger.log('Check ALL tabs: Activity, F1 Receiving, F2 Adjustments, F3 Transfers, F4 Manufacturing, F5 Shipping');

  return {
    execIds: { f1: f1ExecId, f2: f2ExecId, f2b: f2bExecId, f3: f3ExecId, f4: f4ExecId, f5: f5ExecId },
    status: 'All flows logged — check tracker sheet'
  };
}

// ============================================
// REAL TEST FOR F2 STOCK ADJUSTMENT
// ============================================

/**
 * Test real WASP quantity removal
 * Run this to test the F2 flow
 */
function testRealRemove() {
  Logger.log('[ignore] REAL_TEST_START');

  var payload = {
    source: 'WASP',
    event: 'quantity_removed',
    AssetTag: 'B-WAX',
    Quantity: 1,
    LocationCode: 'PRODUCTION',
    SiteName: 'MMH Kelowna'
  };

  var result = handleWaspQuantityRemoved(payload);

  Logger.log('[skip] REAL_TEST_RESULT');
  Logger.log(JSON.stringify(result));

  return result;
}

/**
 * One-time cleanup: Remove duplicate G-OIL from RECEIVING-DOCK
 * The ADD_BATCHES bug put 1600 G-OIL (lot 08002/25) at wrong location.
 * Correct stock is at SW-STORAGE. Run once, then delete this function.
 */
function cleanup_RemoveDuplicateGOIL() {
  var result = waspRemoveInventoryWithLot(
    'G-OIL',
    1600,
    'RECEIVING-DOCK',
    '08002/25',
    'Cleanup: duplicate from ADD_BATCHES location bug',
    'MMH Kelowna',
    '2028-02-13'
  );
  Logger.log('G-OIL cleanup result: ' + JSON.stringify(result));
  return result;
}

/**
 * One-time cleanup: Remove 850 duplicate UFC-4OZ from PROD-RECEIVING
 * MO-7159 was processed twice (WK-808 + WK-810), adding 1700 instead of 850.
 * Run once, then delete this function.
 */
function cleanup_RemoveDuplicateMO7159() {
  var result = waspRemoveInventory(
    'UFC-4OZ',
    850,
    'PROD-RECEIVING',
    'Cleanup: MO-7159 duplicate processing (WK-808 + WK-810)',
    'MMH Kelowna'
  );
  Logger.log('MO-7159 duplicate cleanup result: ' + JSON.stringify(result));
  return result;
}

/**
 * Audit duplicate MO processing — scans Activity sheet for F4 entries
 * that appear more than once with Complete status for the same MO ref.
 * For each duplicate, fetches MO output qty from Katana and generates
 * WASP removal commands to correct the excess inventory.
 *
 * Known affected MOs: MO-7183, MO-7184, MO-7171, MO-7175, MO-7185, MO-7189
 *
 * Run once, review Logger output, then run correction functions if correct.
 */
function auditDuplicateMOs() {
  var activitySheet = getActivitySheet();
  var lastRow = activitySheet.getLastRow();
  if (lastRow <= 3) { Logger.log('No activity data'); return; }

  var data = activitySheet.getRange(4, 1, lastRow - 3, 5).getValues();

  // Find all F4 Complete entries and group by MO ref
  var moEntries = {};
  for (var i = 0; i < data.length; i++) {
    if (!data[i][0]) continue;
    if (String(data[i][2]) !== 'F4 Manufacturing') continue;
    var status = String(data[i][4]);
    if (status !== 'Complete' && status !== 'Partial') continue;
    var details = String(data[i][3]);

    // Extract MO ref (e.g., "MO-7183") from details
    var moMatch = details.match(/MO-\d+/);
    if (!moMatch) continue;
    var moRef = moMatch[0];

    if (!moEntries[moRef]) moEntries[moRef] = [];
    moEntries[moRef].push({
      row: i + 4,
      timestamp: data[i][0],
      details: details
    });
  }

  // Find duplicates
  var duplicates = [];
  var moKeys = Object.keys(moEntries);
  for (var k = 0; k < moKeys.length; k++) {
    if (moEntries[moKeys[k]].length > 1) {
      duplicates.push({ moRef: moKeys[k], entries: moEntries[moKeys[k]] });
    }
  }

  if (duplicates.length === 0) {
    Logger.log('No duplicate MOs found');
    return duplicates;
  }

  Logger.log('=== DUPLICATE MO AUDIT ===');
  Logger.log('Found ' + duplicates.length + ' MOs with duplicate processing:');

  var corrections = [];
  for (var d = 0; d < duplicates.length; d++) {
    var dup = duplicates[d];
    Logger.log('\n' + dup.moRef + ' — processed ' + dup.entries.length + ' times:');
    for (var e = 0; e < dup.entries.length; e++) {
      Logger.log('  Row ' + dup.entries[e].row + ': ' + dup.entries[e].timestamp);
    }

    // Extract SKU and quantity from details text (format: "MO-XXXX  SKU xQTY")
    var skuMatch = dup.entries[0].details.match(/MO-\d+\s+(\S+)\s+x(\d+)/);
    if (skuMatch) {
      var sku = skuMatch[1];
      var qty = parseInt(skuMatch[2], 10);
      var excessTimes = dup.entries.length - 1;
      var excessQty = qty * excessTimes;

      Logger.log('  SKU: ' + sku + ', Qty per MO: ' + qty + ', Excess: ' + excessQty + ' (' + excessTimes + ' extra runs)');
      Logger.log('  CORRECTION: waspRemoveInventory("' + sku + '", ' + excessQty + ', "PROD-RECEIVING", "Correction: ' + dup.moRef + ' duplicate x' + excessTimes + '", "MMH Kelowna")');

      corrections.push({
        moRef: dup.moRef,
        sku: sku,
        qtyPerMO: qty,
        timesProcessed: dup.entries.length,
        excessQty: excessQty,
        location: 'PROD-RECEIVING'
      });
    } else {
      Logger.log('  Could not parse SKU/qty from details — review manually');
    }
  }

  Logger.log('\n=== CORRECTION SUMMARY ===');
  for (var c = 0; c < corrections.length; c++) {
    var corr = corrections[c];
    Logger.log(corr.moRef + ': remove ' + corr.excessQty + ' x ' + corr.sku + ' from ' + corr.location);
  }
  Logger.log('\nRun executeDuplicateMOCorrections() to apply these corrections.');

  return corrections;
}

/**
 * Execute corrections generated by auditDuplicateMOs().
 * Removes excess inventory from PROD-RECEIVING for each duplicate-processed MO.
 * Run auditDuplicateMOs() first to review, then run this to apply.
 */
function executeDuplicateMOCorrections() {
  var corrections = auditDuplicateMOs();
  if (!corrections || corrections.length === 0) {
    Logger.log('No corrections to apply');
    return;
  }

  Logger.log('\n=== EXECUTING CORRECTIONS ===');
  var results = [];
  for (var c = 0; c < corrections.length; c++) {
    var corr = corrections[c];
    Logger.log('Removing ' + corr.excessQty + ' x ' + corr.sku + ' from ' + corr.location + ' (' + corr.moRef + ')...');
    var result = waspRemoveInventory(
      corr.sku,
      corr.excessQty,
      corr.location,
      'Correction: ' + corr.moRef + ' duplicate processing (x' + corr.timesProcessed + ')',
      'MMH Kelowna'
    );
    Logger.log('  Result: ' + JSON.stringify(result));
    results.push({ moRef: corr.moRef, sku: corr.sku, qty: corr.excessQty, result: result });
  }

  Logger.log('\n=== CORRECTION RESULTS ===');
  for (var r = 0; r < results.length; r++) {
    var res = results[r];
    var ok = res.result && res.result.success;
    Logger.log((ok ? 'OK' : 'FAILED') + ' ' + res.moRef + ': ' + res.sku + ' x' + res.qty + (ok ? ' removed' : ' FAILED'));
  }

  return results;
}

/**
 * Audit duplicate PO receiving — scans Activity sheet for F1 entries
 * that appear more than once with Received status for the same PO ref.
 * Generates WASP removal commands to correct the excess inventory.
 *
 * Run once, review Logger output, then run correction functions if correct.
 */
function auditDuplicatePOs() {
  var activitySheet = getActivitySheet();
  var lastRow = activitySheet.getLastRow();
  if (lastRow <= 3) { Logger.log('No activity data'); return; }

  var data = activitySheet.getRange(4, 1, lastRow - 3, 5).getValues();

  // Find all F1 Received entries and group by PO ref
  var poEntries = {};
  for (var i = 0; i < data.length; i++) {
    if (!data[i][0]) continue;
    if (String(data[i][2]) !== 'F1 Receiving') continue;
    var status = String(data[i][4]);
    if (status !== 'Received') continue;
    var details = String(data[i][3]);

    var poMatch = details.match(/PO-\d+/);
    if (!poMatch) continue;
    var poRef = poMatch[0];

    if (!poEntries[poRef]) poEntries[poRef] = [];
    poEntries[poRef].push({
      row: i + 4,
      timestamp: data[i][0],
      details: details
    });
  }

  // Find duplicates
  var duplicates = [];
  var poKeys = Object.keys(poEntries);
  for (var k = 0; k < poKeys.length; k++) {
    if (poEntries[poKeys[k]].length > 1) {
      duplicates.push({ poRef: poKeys[k], entries: poEntries[poKeys[k]] });
    }
  }

  if (duplicates.length === 0) {
    Logger.log('No duplicate POs found');
    return duplicates;
  }

  Logger.log('=== DUPLICATE PO AUDIT ===');
  Logger.log('Found ' + duplicates.length + ' POs with duplicate receiving:');

  var corrections = [];
  for (var d = 0; d < duplicates.length; d++) {
    var dup = duplicates[d];
    Logger.log('\n' + dup.poRef + ' — received ' + dup.entries.length + ' times:');
    for (var e = 0; e < dup.entries.length; e++) {
      Logger.log('  Row ' + dup.entries[e].row + ': ' + dup.entries[e].timestamp);
    }

    // Extract SKU and quantity from details text (format: "PO-XXX  SKU xQTY" or "PO-XXX  n items")
    var skuMatch = dup.entries[0].details.match(/PO-\d+\s+(\S+)\s+x(\d+)/);
    if (skuMatch) {
      var sku = skuMatch[1];
      var qty = parseInt(skuMatch[2], 10);
      var excessTimes = dup.entries.length - 1;
      var excessQty = qty * excessTimes;

      Logger.log('  SKU: ' + sku + ', Qty per PO: ' + qty + ', Excess: ' + excessQty);
      Logger.log('  CORRECTION: waspRemoveInventory("' + sku + '", ' + excessQty + ', "RECEIVING-DOCK", "Correction: ' + dup.poRef + ' duplicate x' + excessTimes + '", "MMH Kelowna")');

      corrections.push({
        poRef: dup.poRef,
        sku: sku,
        qtyPerPO: qty,
        timesReceived: dup.entries.length,
        excessQty: excessQty,
        location: 'RECEIVING-DOCK'
      });
    } else {
      Logger.log('  Could not parse SKU/qty — may have multiple items. Review manually.');
      // Try to log the sub-items for manual review
      for (var s = 0; s < dup.entries.length; s++) {
        Logger.log('    Details: ' + dup.entries[s].details);
      }
    }
  }

  Logger.log('\n=== PO CORRECTION SUMMARY ===');
  for (var c = 0; c < corrections.length; c++) {
    var corr = corrections[c];
    Logger.log(corr.poRef + ': remove ' + corr.excessQty + ' x ' + corr.sku + ' from ' + corr.location);
  }

  return corrections;
}

/**
 * Remove test data rows ("TEST DELETE ME") from Activity sheet.
 * Scans from bottom to top and deletes matching rows.
 * Run once after testing is complete.
 */
function cleanupTestRows() {
  var activitySheet = getActivitySheet();
  var lastRow = activitySheet.getLastRow();
  if (lastRow <= 3) { Logger.log('No data to clean'); return; }

  var data = activitySheet.getRange(4, 1, lastRow - 3, 4).getValues();
  var rowsToDelete = [];

  // First pass: find header rows with "TEST DELETE ME"
  var testExecIds = {};
  for (var i = 0; i < data.length; i++) {
    var details = String(data[i][3]);
    if (details.indexOf('TEST DELETE ME') >= 0) {
      rowsToDelete.push(i + 4);
      if (data[i][0]) testExecIds[String(data[i][0])] = true;
    }
  }

  // Second pass: find sub-item rows belonging to test exec IDs
  for (var j = 0; j < data.length; j++) {
    if (!data[j][0] && j > 0) {
      // Sub-item row — check if previous header was a test entry
      var prevRow = j - 1;
      while (prevRow >= 0 && !data[prevRow][0]) prevRow--;
      if (prevRow >= 0 && testExecIds[String(data[prevRow][0])]) {
        if (rowsToDelete.indexOf(j + 4) < 0) rowsToDelete.push(j + 4);
      }
    }
  }

  // Delete from bottom to top to preserve row indices
  rowsToDelete.sort(function(a, b) { return b - a; });

  Logger.log('=== CLEANUP TEST ROWS ===');
  Logger.log('Found ' + rowsToDelete.length + ' test rows to delete');

  for (var d = 0; d < rowsToDelete.length; d++) {
    activitySheet.deleteRow(rowsToDelete[d]);
  }

  Logger.log('Deleted ' + rowsToDelete.length + ' rows');
  return rowsToDelete.length;
}

/**
 * Remove duplicate F4 entries from Activity sheet.
 * Keeps the FIRST entry per MO ref, deletes subsequent ones (header + sub-items).
 * Run from script editor: View → Logs to see results.
 */
function removeDuplicateF4Entries() {
  var activitySheet = getActivitySheet();
  var lastRow = activitySheet.getLastRow();
  if (lastRow <= 3) { Logger.log('No data'); return 0; }

  var data = activitySheet.getRange(4, 1, lastRow - 3, 4).getValues();

  // Find all F4 header rows and group by MO ref
  var seenMOs = {};
  var dupHeaderRows = [];

  for (var i = 0; i < data.length; i++) {
    if (!data[i][0]) continue; // skip sub-items
    if (String(data[i][2]) !== 'F4 Manufacturing') continue;

    var details = String(data[i][3]);
    var moMatch = details.match(/MO-\d+/);
    if (!moMatch) continue;
    var moRef = moMatch[0];

    if (seenMOs[moRef]) {
      // Duplicate — mark for deletion
      dupHeaderRows.push(i + 4); // sheet row
      Logger.log('Duplicate: row ' + (i + 4) + ' ' + String(data[i][0]) + ' ' + moRef);
    } else {
      seenMOs[moRef] = true;
    }
  }

  if (dupHeaderRows.length === 0) {
    Logger.log('No duplicate F4 entries found');
    return 0;
  }

  // Build set of duplicate header rows for quick lookup
  var dupSet = {};
  for (var h = 0; h < dupHeaderRows.length; h++) {
    dupSet[dupHeaderRows[h]] = true;
  }

  // Collect all rows to delete: duplicate headers + their sub-items
  var rowsToDelete = [];
  for (var j = 0; j < data.length; j++) {
    var sheetRow = j + 4;
    if (dupSet[sheetRow]) {
      rowsToDelete.push(sheetRow);
      // Collect sub-item rows (empty col A) immediately below
      var sub = j + 1;
      while (sub < data.length && !data[sub][0]) {
        rowsToDelete.push(sub + 4);
        sub++;
      }
    }
  }

  // Delete from bottom to top
  rowsToDelete.sort(function(a, b) { return b - a; });

  Logger.log('Deleting ' + rowsToDelete.length + ' rows (' + dupHeaderRows.length + ' duplicate entries)');
  for (var d = 0; d < rowsToDelete.length; d++) {
    activitySheet.deleteRow(rowsToDelete[d]);
  }

  Logger.log('Done — removed ' + dupHeaderRows.length + ' duplicate F4 entries');
  return dupHeaderRows.length;
}

/**
 * DIAGNOSTIC: Inspect raw Katana API data for MO ingredients batch_transactions.
 * Run from script editor: testMOBatchData()
 * Change moOrderNo below to match the MO you want to inspect.
 * Check View → Logs for full output.
 */
function findManufacturingOrderByOrderNo_(moOrderNo) {
  var normalizedTarget = String(moOrderNo || '').trim();
  if (!normalizedTarget) return null;

  for (var page = 1; page <= 5; page++) {
    var searchResult = katanaApiCall('manufacturing_orders?limit=250&page=' + page);
    var mos = searchResult && searchResult.data ? searchResult.data : (searchResult || []);
    if (mos.length === 0) break;

    for (var i = 0; i < mos.length; i++) {
      var orderNo = String(mos[i].order_no || '').trim();
      if (orderNo === normalizedTarget || orderNo.indexOf(normalizedTarget) === 0) {
        return mos[i];
      }
    }
  }

  return null;
}

function writeMOBatchDiagnosticSheet_(moOrderNo, mo, rowsWithInclude, rowsWithoutInclude) {
  var ss = SpreadsheetApp.openById(CONFIG.DEBUG_SHEET_ID);
  var sheetName = 'MO Debug';
  var sheet = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);
  var existingOverrides = readMOBatchDebugOverridesFromSheet_(sheet);
  sheet.clearContents();
  sheet.clearFormats();
  sheet.showSheet();

  var outputBatchNode = mo.output_batch || mo.outputBatch || null;
  var headerBatchTransactions = mo.batch_transactions || mo.batchTransactions || [];
  var resolvedOutput = resolveMOOutputBatchData_(mo, mo.variant_id || '', moOrderNo);
  var summaryRows = [
    ['MO Batch Debug', '', '', '', '', ''],
    ['Order #', moOrderNo, 'MO ID', mo.id || '', 'Status', mo.status || ''],
    ['Variant ID', mo.variant_id || '', 'Batch #', resolvedOutput.lot || mo.batch_number || mo.output_batch_number || mo.lot_number || mo.batch || '', 'Done At', mo.done_at || mo.completed_at || mo.updated_at || ''],
    ['Output Batch (raw)', outputBatchNode ? JSON.stringify(outputBatchNode) : '(none)', 'Header Batch Tx (raw)', headerBatchTransactions.length ? JSON.stringify(headerBatchTransactions) : '(none)', '', ''],
    ['Resolved Output Lot', resolvedOutput.lot || '(none)', 'Resolved Output Expiry', resolvedOutput.expiry || '(none)', 'Resolved From', resolvedOutput.source || '(none)'],
    ['Recipe rows WITH include=batch_transactions', rowsWithInclude.length, 'Recipe rows WITHOUT include', rowsWithoutInclude.length, '', '']
  ];
  sheet.getRange(1, 1, summaryRows.length, 6).setValues(summaryRows);
  sheet.getRange(1, 1, 1, 6).merge().setFontWeight('bold').setBackground('#d9ead3');
  sheet.getRange(2, 1, summaryRows.length - 1, 6).setWrap(true);

  var headerRow = 7;
  var headers = [[
    'Row',
    'SKU',
    'Variant ID',
    'Material ID',
    'Batch Tracked',
    'Qty',
    'Consumed Qty',
    'Direct Batch',
    'Direct Expiry',
    'Batch Transactions',
    'Stock Allocations',
    'Picked Batches'
  ]];
  sheet.getRange(headerRow, 1, 1, headers[0].length).setValues(headers).setFontWeight('bold').setBackground('#cfe2f3');

  var dataRows = [];
  for (var i = 0; i < rowsWithInclude.length; i++) {
    var row = rowsWithInclude[i];
    var variantData = row.variant_id ? fetchKatanaVariant(row.variant_id) : null;
    var variant = variantData && variantData.data ? variantData.data : variantData;
    dataRows.push([
      i + 1,
      variant ? (variant.sku || '') : '',
      row.variant_id || '',
      variant ? (variant.material_id || '') : '',
      variant ? String(isKatanaBatchTrackedVariant_(variant)) : '',
      row.quantity || '',
      row.consumed_quantity || row.total_consumed_quantity || row.actual_quantity || '',
      row.batch_number || row.batch_nr || '',
      extractKatanaExpiryDate_(row) || '',
      JSON.stringify(row.batch_transactions || []),
      JSON.stringify(row.stock_allocations || []),
      JSON.stringify(row.picked_batches || [])
    ]);
  }

  if (dataRows.length > 0) {
    sheet.getRange(headerRow + 1, 1, dataRows.length, headers[0].length).setValues(dataRows).setWrap(true);
  } else {
    sheet.getRange(headerRow + 1, 1).setValue('No recipe rows returned.');
  }

  var overrideSectionRow = headerRow + Math.max(dataRows.length, 1) + 3;
  sheet.getRange(overrideSectionRow, 1, 1, 8).merge().setValue('Repair Overrides').setFontWeight('bold').setBackground('#f4cccc');
  var overrideHeaderRow = overrideSectionRow + 1;
  var overrideHeaders = [[
    'SKU',
    'Qty',
    'Batch ID',
    'Resolved Lot',
    'Resolved Expiry',
    'Override Lot',
    'Override Expiry',
    'Notes'
  ]];
  sheet.getRange(overrideHeaderRow, 1, 1, overrideHeaders[0].length).setValues(overrideHeaders).setFontWeight('bold').setBackground('#fce5cd');

  var overrideRows = [];
  for (var oi = 0; oi < rowsWithInclude.length; oi++) {
    var oRow = rowsWithInclude[oi];
    if (!oRow || !oRow.variant_id) continue;
    var oVariantData = fetchKatanaVariant(oRow.variant_id);
    var oVariant = oVariantData && oVariantData.data ? oVariantData.data : oVariantData;
    var oSku = oVariant ? String(oVariant.sku || '').trim() : '';
    if (!oSku) continue;
    var oBt = oRow.batch_transactions || oRow.batchTransactions || [];
    if (!oBt.length) continue;

    for (var ob = 0; ob < oBt.length; ob++) {
      var obResolved = resolveMORepairBatchTransaction_(oRow, oVariant, oBt[ob]);
      var obKey = formatMODebugOverrideKey_(oSku, obResolved.qty || 0, obResolved.batchId || '');
      var existing = existingOverrides[obKey] || {};
      overrideRows.push([
        oSku,
        obResolved.qty || 0,
        obResolved.batchId || '',
        obResolved.lot || '',
        obResolved.expiry || '',
        existing.lot || '',
        existing.expiry || '',
        existing.notes || ''
      ]);
    }
  }

  if (overrideRows.length > 0) {
    sheet.getRange(overrideHeaderRow + 1, 1, overrideRows.length, 8).setValues(overrideRows).setWrap(true);
  } else {
    sheet.getRange(overrideHeaderRow + 1, 1).setValue('No batch-tracked ingredient transactions found.');
  }

  sheet.setFrozenRows(headerRow);
  sheet.autoResizeColumns(1, 12);
  sheet.setColumnWidths(10, 3, 260);
  sheet.setColumnWidths(6, 3, 180);
  ss.setActiveSheet(sheet);
}

function formatMODebugOverrideKey_(sku, qty, batchId) {
  return [
    String(sku || '').trim(),
    String(parseFloat(qty || 0) || 0),
    String(batchId || '').trim()
  ].join('|');
}

function readMOBatchDebugOverridesFromSheet_(sheet) {
  var map = {};
  if (!sheet) return map;
  var lastRow = sheet.getLastRow();
  if (lastRow < 1) return map;

  var values = sheet.getRange(1, 1, lastRow, 8).getValues();
  var startRow = 0;
  for (var i = 0; i < values.length; i++) {
    if (String(values[i][0] || '').trim() === 'Repair Overrides') {
      startRow = i + 3; // title row + header row + first data row
      break;
    }
  }
  if (!startRow) return map;

  for (var r = startRow; r <= values.length; r++) {
    var row = values[r - 1];
    var sku = String(row[0] || '').trim();
    var qty = parseFloat(row[1] || 0) || 0;
    var batchId = String(row[2] || '').trim();
    var overrideLot = String(row[5] || '').trim();
    var overrideExpiry = normalizeBusinessDate_(row[6] || '');
    var notes = String(row[7] || '').trim();
    if (!sku && !batchId && !overrideLot && !overrideExpiry) continue;
    var key = formatMODebugOverrideKey_(sku, qty, batchId);
    map[key] = {
      lot: overrideLot,
      expiry: overrideExpiry,
      notes: notes
    };
  }

  return map;
}

function collectMOBatchDataForOrder_(moOrderNo) {
  moOrderNo = String(moOrderNo || '').trim();
  if (!moOrderNo) throw new Error('No MO order number provided');

  var moSearch = findManufacturingOrderByOrderNo_(moOrderNo);
  if (!moSearch) throw new Error('MO not found: ' + moOrderNo);

  var moId = moSearch.id;
  var moData = fetchKatanaMO(moId);
  var mo = moData && moData.data ? moData.data : (moData || moSearch);
  if (!mo) throw new Error('Failed to fetch MO header: ' + moOrderNo);

  var withInclude = katanaApiCall('manufacturing_order_recipe_rows?manufacturing_order_id=' + moId + '&include=batch_transactions');
  var rowsWithInclude = withInclude && withInclude.data ? withInclude.data : (withInclude || []);
  var withoutInclude = katanaApiCall('manufacturing_order_recipe_rows?manufacturing_order_id=' + moId);
  var rowsWithoutInclude = withoutInclude && withoutInclude.data ? withoutInclude.data : (withoutInclude || []);

  return {
    moOrderNo: String(mo.order_no || moOrderNo).trim(),
    moId: moId,
    mo: mo,
    rowsWithInclude: rowsWithInclude,
    rowsWithoutInclude: rowsWithoutInclude
  };
}

function debugMOBatchDataForOrder_(moOrderNo) {
  var batchData = collectMOBatchDataForOrder_(moOrderNo);
  writeMOBatchDiagnosticSheet_(batchData.moOrderNo, batchData.mo, batchData.rowsWithInclude, batchData.rowsWithoutInclude);
  return {
    moId: batchData.moId,
    rowsWithInclude: batchData.rowsWithInclude.length,
    rowsWithoutInclude: batchData.rowsWithoutInclude.length
  };
}

function debugMOBatchDataPrompt() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt('Debug MO Batch Data', 'Enter the Katana MO order number (for example: MO-7407 TEST)', ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() !== ui.Button.OK) return;
  var moOrderNo = String(response.getResponseText() || '').trim();
  if (!moOrderNo) {
    ui.alert('Enter an MO order number.');
    return;
  }

  var result = debugMOBatchDataForOrder_(moOrderNo);
  ui.alert('MO Debug ready for ' + moOrderNo + ' (' + result.rowsWithInclude + ' recipe rows).');
}

function getMOBatchDebugOrderNumber_() {
  var ss = SpreadsheetApp.openById(CONFIG.DEBUG_SHEET_ID);
  var active = ss.getActiveSheet();
  var candidate = '';

  if (active && active.getName() === 'MO Debug') {
    candidate = String(active.getRange(2, 2).getValue() || '').trim();
    if (candidate) return candidate;
  }

  var sheet = ss.getSheetByName('MO Debug');
  if (sheet) {
    candidate = String(sheet.getRange(2, 2).getValue() || '').trim();
    if (candidate) return candidate;
  }

  return '';
}

function findLatestF4HeaderRowForOrder_(sheet, moOrderNo, outputSku) {
  if (!sheet || !moOrderNo) return 0;
  var lastRow = sheet.getLastRow();
  if (lastRow < 4) return 0;

  var data = sheet.getRange(4, 1, lastRow - 3, 4).getValues();
  var fallbackRow = 0;
  for (var i = data.length - 1; i >= 0; i--) {
    var execId = String(data[i][0] || '').trim();
    var flow = String(data[i][2] || '').trim();
    var details = String(data[i][3] || '').trim();
    if (execId.indexOf('WK-') !== 0) continue;
    if (flow !== 'F4 Manufacturing') continue;
    if (details.indexOf(moOrderNo) < 0) continue;
    if (!fallbackRow) fallbackRow = i + 4;
    if (outputSku && details.indexOf(outputSku) >= 0) return i + 4;
  }

  return fallbackRow;
}

function findAllF4HeaderRowsForOrder_(sheet, moOrderNo, outputSku) {
  if (!sheet || !moOrderNo) return [];
  var lastRow = sheet.getLastRow();
  if (lastRow < 4) return [];

  var data = sheet.getRange(4, 1, lastRow - 3, 4).getValues();
  var rows = [];

  for (var i = 0; i < data.length; i++) {
    var execId = String(data[i][0] || '').trim();
    var flow = String(data[i][2] || '').trim();
    var details = String(data[i][3] || '').trim();
    if (execId.indexOf('WK-') !== 0) continue;
    if (flow !== 'F4 Manufacturing') continue;
    if (details.indexOf(moOrderNo) < 0) continue;
    if (outputSku && details.indexOf(outputSku) < 0) continue;
    rows.push(i + 4);
  }

  return rows;
}

function isQtyOnlyMORepairLine_(detail) {
  var clean = String(detail || '').replace(/[├└─│\s]+/g, ' ').trim();
  return /^x[\d.]+\b/i.test(clean);
}

function parseMORepairSubItem_(detail, parentContext) {
  var parsed = parseActivitySubItem(detail);
  if (!parsed) return null;

  if (isQtyOnlyMORepairLine_(detail) && parentContext && parentContext.sku) {
    parsed.sku = parentContext.sku;
    if (!parsed.location) parsed.location = parentContext.location || '';
  }

  return parsed;
}

function buildMORepairMatchKey_(isOutput, sku, qty, location, lot, expiry) {
  return [
    isOutput ? 'OUT' : 'ING',
    String(sku || '').trim(),
    String(qty || '').trim(),
    String(location || '').trim(),
    String(lot || '').trim(),
    String(expiry || '').trim()
  ].join('|');
}

function resolveMORepairBatchTransaction_(row, variant, btEntry) {
  var batchId = btEntry.batch_id || btEntry.batchId || null;
  var lot = extractKatanaBatchNumber_(btEntry);
  var expiry = extractKatanaExpiryDate_(btEntry);

  var batchStock = btEntry.batch_stock || btEntry.batchStock || null;
  if (!lot && batchStock) {
    lot = extractKatanaBatchNumber_(batchStock) || '';
  }
  if (!expiry && batchStock) {
    expiry = extractKatanaExpiryDate_(batchStock) || '';
  }

  if (batchId && (!lot || !expiry)) {
    var batchInfo = fetchKatanaBatchStock(batchId);
    if (batchInfo) {
      lot = lot || extractKatanaBatchNumber_(batchInfo) || '';
      expiry = expiry || extractKatanaExpiryDate_(batchInfo) || '';
    }
  }

  if ((!lot || !expiry) && batchId && row && row.variant_id) {
    var varResult = katanaApiCall('batch_stocks?variant_id=' + row.variant_id + '&include_deleted=true');
    var varList = varResult && (varResult.data || varResult) || [];
    for (var i = 0; i < varList.length; i++) {
      var candidate = varList[i];
      var candidateId = candidate.batch_id || candidate.id;
      if (String(candidateId) !== String(batchId)) continue;
      lot = lot || candidate.batch_number || candidate.nr || candidate.number || '';
      expiry = expiry || extractKatanaExpiryDate_(candidate) || '';
      break;
    }
  }

  expiry = normalizeBusinessDate_(expiry || '');
  return {
    qty: btEntry.quantity || 0,
    batchId: batchId || '',
    lot: lot || '',
    expiry: expiry || ''
  };
}

function buildMOIngredientRepairQueues_(rowsWithInclude, overrideMap) {
  var queues = {};
  overrideMap = overrideMap || {};

  for (var i = 0; i < rowsWithInclude.length; i++) {
    var row = rowsWithInclude[i];
    if (!row || !row.variant_id) continue;

    var variantData = fetchKatanaVariant(row.variant_id);
    var variant = variantData && variantData.data ? variantData.data : variantData;
    var sku = variant ? String(variant.sku || '').trim() : '';
    if (!sku) continue;

    var batchTracked = !!isKatanaBatchTrackedVariant_(variant);

    if (!queues[sku]) queues[sku] = [];
    var bt = row.batch_transactions || row.batchTransactions || [];
    var queueItems = [];

    if (bt.length > 1) {
      for (var bti = 0; bti < bt.length; bti++) {
        var btResolved = resolveMORepairBatchTransaction_(row, variant, bt[bti]);
        var btOverride = overrideMap[formatMODebugOverrideKey_(sku, btResolved.qty || 0, btResolved.batchId || '')] || null;
        if (btOverride && btOverride.lot) {
          btResolved.lot = btOverride.lot;
          btResolved.expiry = btOverride.expiry || btResolved.expiry || '';
        }
        queueItems.push({
          sku: sku,
          qty: btResolved.qty || 0,
          batchId: btResolved.batchId || '',
          lot: btResolved.lot || '',
          expiry: btResolved.expiry || '',
          location: FLOWS.MO_INGREDIENT_LOCATION,
          batchTracked: true,
          unresolvedReason: btResolved.lot ? '' : ('Katana batch still missing on MO ingredient' + (btResolved.batchId ? ' (batch_id:' + btResolved.batchId + ')' : ''))
        });
      }
    } else {
      var qty = row.total_consumed_quantity || row.consumed_quantity || row.actual_quantity || row.quantity || 0;
      var lot = extractIngredientBatchNumber(row) || '';
      var expiry = normalizeBusinessDate_(extractIngredientExpiryDate(row) || '');
      var singleBatchId = bt.length > 0 ? (bt[0].batch_id || bt[0].batchId || '') : '';
      var singleOverride = overrideMap[formatMODebugOverrideKey_(sku, qty, singleBatchId)] || null;
      if (singleOverride && singleOverride.lot) {
        lot = singleOverride.lot;
        expiry = singleOverride.expiry || expiry || '';
      }
      queueItems.push({
        sku: sku,
        qty: qty,
        batchId: singleBatchId || '',
        lot: lot,
        expiry: expiry,
        location: FLOWS.MO_INGREDIENT_LOCATION,
        batchTracked: batchTracked,
        unresolvedReason: batchTracked && !lot ? ('Katana batch still missing on MO ingredient' + (singleBatchId ? ' (batch_id:' + singleBatchId + ')' : '')) : ''
      });
    }

    for (var qi = 0; qi < queueItems.length; qi++) {
      var q = queueItems[qi];
      if (q.lot) {
        var exactWasp = waspLookupExactLotAndDate(
          sku,
          FLOWS.MO_INGREDIENT_LOCATION,
          q.lot,
          CONFIG.WASP_SITE,
          q.expiry || ''
        );
        if (exactWasp && !exactWasp.ambiguous && ((!q.expiry || exactWasp.dateToleranceApplied) && exactWasp.dateCode)) {
          q.expiry = exactWasp.dateCode;
        }
      }
      queues[sku].push(q);
    }
  }

  inferMissingMORepairLotsFromWasp_(queues);
  return queues;
}

function inferMissingMORepairLotsFromWasp_(queues) {
  for (var sku in queues) {
    if (!queues.hasOwnProperty(sku)) continue;
    var queue = queues[sku] || [];
    if (!queue.length) continue;

    var resolvedKeys = {};
    for (var i = 0; i < queue.length; i++) {
      var item = queue[i];
      if (item.lot) {
        resolvedKeys[(item.lot || '') + '|' + (item.expiry || '')] = true;
      }
    }

    var allRows = waspLookupAllLotRows(sku, FLOWS.MO_INGREDIENT_LOCATION, CONFIG.WASP_SITE);
    if (!allRows.length) continue;

    var claimedKeys = {};
    for (var q = 0; q < queue.length; q++) {
      var unresolved = queue[q];
      if (!unresolved.batchTracked || unresolved.lot || !unresolved.qty) continue;

      var candidates = [];
      for (var r = 0; r < allRows.length; r++) {
        var row = allRows[r];
        var rowKey = (row.lot || '') + '|' + (row.dateCode || '');
        if (resolvedKeys[rowKey]) continue;
        if (claimedKeys[rowKey]) continue;
        if ((parseFloat(row.qtyAvailable || 0) || 0) < (parseFloat(unresolved.qty || 0) || 0)) continue;
        candidates.push(row);
      }

      if (candidates.length === 1) {
        unresolved.lot = candidates[0].lot || '';
        unresolved.expiry = candidates[0].dateCode || '';
        unresolved.unresolvedReason = '';
        unresolved.inferredFromWasp = true;
        claimedKeys[(unresolved.lot || '') + '|' + (unresolved.expiry || '')] = true;
        logToSheet('MO_REPAIR_WASP_LOT_INFERRED', {
          sku: sku,
          qty: unresolved.qty,
          lot: unresolved.lot,
          expiry: unresolved.expiry,
          location: unresolved.location || FLOWS.MO_INGREDIENT_LOCATION
        }, 'Resolved missing Katana batch from unique remaining WASP lot candidate during MO repair');
      }
    }
  }
}

function resolveMOOutputRepairData_(mo) {
  mo = mo || {};
  var variantId = mo.variant_id || '';
  var outputSku = '';
  var productName = '';
  var productCategory = '';
  var variant = null;

  if (variantId) {
    var variantData = fetchKatanaVariant(variantId);
    variant = variantData && variantData.data ? variantData.data : variantData;
    outputSku = variant ? (variant.sku || '') : '';
    productName = variant ? (variant.product ? variant.product.name : variant.name) : '';
    if (!productName) productName = mo.product_name || '';
    productCategory = variant ? (variant.product ? (variant.product.category || variant.product.category_name || '') : '') : '';
    if (!productCategory && variant && variant.product_id) {
      var productData = fetchKatanaProduct(variant.product_id);
      var product = productData && productData.data ? productData.data : productData;
      productCategory = product ? (product.category || product.category_name || '') : '';
    }
  }

  var outputBatchData = resolveMOOutputBatchData_(mo, variantId, mo.order_no || '');
  var lotNumber = outputBatchData.lot || '';
  var expiryDate = outputBatchData.expiry || '';

  if (lotNumber && !expiryDate) {
    var completionSrc = mo.completed_at || mo.done_at || mo.updated_at || null;
    var baseDate = completionSrc ? new Date(completionSrc) : new Date();
    var expDate = new Date(baseDate.getFullYear() + 3, baseDate.getMonth(), baseDate.getDate());
    expiryDate = expDate.getFullYear() + '-' +
      ('0' + (expDate.getMonth() + 1)).slice(-2) + '-' +
      ('0' + expDate.getDate()).slice(-2);
  }

  var stage = detectMOStage(productName, outputSku, productCategory);
  return {
    sku: outputSku,
    qty: mo.actual_quantity || mo.quantity || 0,
    lot: lotNumber || '',
    expiry: expiryDate || '',
    location: stage === 'FINISHED' ? FLOWS.MO_OUTPUT_LOCATION : FLOWS.MO_INGREDIENT_LOCATION,
    unresolvedReason: lotNumber ? '' : 'Katana output batch still missing'
  };
}

function buildMORepairPlanFromSourceRows_(sourceRows, repairContext) {
  var planRows = [];
  var ingredientQueues = {};
  var sku;

  for (sku in repairContext.ingredientsBySku) {
    if (!repairContext.ingredientsBySku.hasOwnProperty(sku)) continue;
    ingredientQueues[sku] = repairContext.ingredientsBySku[sku].slice();
  }

  for (var i = 0; i < sourceRows.length; i++) {
    var sourceRow = sourceRows[i];
    var parsed = sourceRow.parsed;
    var isOutput = parsed.sku === repairContext.output.sku;
    var plan = null;

    if (isOutput) {
      plan = {
        type: 'output',
        sku: repairContext.output.sku,
        qty: repairContext.output.qty || parsed.qty,
        location: repairContext.output.location,
        lot: repairContext.output.lot || '',
        expiry: repairContext.output.expiry || '',
        unresolvedReason: repairContext.output.unresolvedReason || ''
      };
    } else {
      var queue = ingredientQueues[parsed.sku] || [];
      plan = queue.length ? queue.shift() : null;
      ingredientQueues[parsed.sku] = queue;
      if (!plan) {
        plan = {
          type: 'ingredient',
          sku: parsed.sku,
          qty: parsed.qty,
          batchId: '',
          location: parsed.location || FLOWS.MO_INGREDIENT_LOCATION,
          lot: parsed.lot || '',
          expiry: parsed.expiry || '',
          batchTracked: false,
          unresolvedReason: 'No Katana ingredient row matched this failed item'
        };
      } else {
        plan.type = 'ingredient';
      }
    }

    planRows.push({
      rowIndex: sourceRow.rowIndex,
      key: buildMORepairMatchKey_(
        isOutput,
        parsed.sku,
        parsed.qty,
        parsed.location,
        parsed.lot,
        parsed.expiry
      ),
      parsed: parsed,
      plan: plan
    });
  }

  return planRows;
}

function quantitiesEffectivelyEqual_(left, right) {
  var a = parseFloat(left || 0) || 0;
  var b = parseFloat(right || 0) || 0;
  return Math.abs(a - b) <= 0.0001;
}

function trySafeMODateCodeRepairs_(planRows, moOrderNo) {
  var repaired = 0;
  var skipped = 0;
  var details = [];
  var seen = {};

  for (var i = 0; i < planRows.length; i++) {
    var item = planRows[i];
    var plan = item.plan || {};
    if (plan.type === 'output') continue;
    if (!plan.batchTracked || !plan.lot || !plan.expiry) continue;

    var dedupKey = [plan.sku, plan.location, plan.lot, plan.expiry, plan.qty].join('|');
    if (seen[dedupKey]) continue;
    seen[dedupKey] = true;

    var exact = waspLookupExactLotAndDate(
      plan.sku,
      plan.location,
      plan.lot,
      CONFIG.WASP_SITE,
      plan.expiry
    );
    if (exact && !exact.ambiguous) continue;

    var lotRows = waspLookupLotRows(plan.sku, plan.location, CONFIG.WASP_SITE, plan.lot);
    var alternateRows = [];
    for (var lr = 0; lr < lotRows.length; lr++) {
      var lotRow = lotRows[lr];
      if ((lotRow.dateCode || '') === (plan.expiry || '')) continue;
      alternateRows.push(lotRow);
    }

    if (alternateRows.length !== 1) {
      skipped++;
      details.push(plan.sku + ' lot ' + plan.lot + ' skipped: expected 1 alternate row, found ' + alternateRows.length);
      continue;
    }

    var existing = alternateRows[0];
    if (!quantitiesEffectivelyEqual_(existing.qtyAvailable, plan.qty)) {
      skipped++;
      details.push(plan.sku + ' lot ' + plan.lot + ' skipped: qty mismatch existing=' + existing.qtyAvailable + ' expected=' + plan.qty);
      continue;
    }

    var noteBase = '[MO-DEBUG-DATE-REPAIR] ' + moOrderNo + ' ' + plan.sku + ' lot ' + plan.lot;

    markSyncedToWasp(plan.sku, plan.location, 'remove');
    var removeResult = waspRemoveInventoryWithLot(
      plan.sku,
      existing.qtyAvailable,
      plan.location,
      plan.lot,
      noteBase + ' remove wrong date ' + existing.dateCode,
      CONFIG.WASP_SITE,
      existing.dateCode
    );

    if (!removeResult.success) {
      skipped++;
      details.push(plan.sku + ' lot ' + plan.lot + ' skipped: remove old date failed');
      continue;
    }

    markSyncedToWasp(plan.sku, plan.location, 'add');
    var addResult = waspAddInventoryWithLot(
      plan.sku,
      existing.qtyAvailable,
      plan.location,
      plan.lot,
      plan.expiry,
      noteBase + ' add corrected date ' + plan.expiry,
      CONFIG.WASP_SITE
    );

    if (!addResult.success) {
      markSyncedToWasp(plan.sku, plan.location, 'add');
      waspAddInventoryWithLot(
        plan.sku,
        existing.qtyAvailable,
        plan.location,
        plan.lot,
        existing.dateCode,
        noteBase + ' rollback to ' + existing.dateCode,
        CONFIG.WASP_SITE
      );
      skipped++;
      details.push(plan.sku + ' lot ' + plan.lot + ' skipped: add corrected date failed, rolled back');
      continue;
    }

    logAdjustment('Google Sheet', 'Remove', '', plan.sku, '', CONFIG.WASP_SITE, plan.location, plan.lot, existing.dateCode, -Math.abs(existing.qtyAvailable), '', null, 'OK');
    logAdjustment('Google Sheet', 'Add', '', plan.sku, '', CONFIG.WASP_SITE, plan.location, plan.lot, plan.expiry, Math.abs(existing.qtyAvailable), '', null, 'OK');

    repaired++;
    details.push(plan.sku + ' lot ' + plan.lot + ' corrected from ' + existing.dateCode + ' to ' + plan.expiry);
  }

  return {
    repaired: repaired,
    skipped: skipped,
    details: details
  };
}

function executeMORepairPlan_(planRows) {
  var executed = [];

  for (var i = 0; i < planRows.length; i++) {
    var item = planRows[i];
    var plan = item.plan || {};
    var action = plan.type === 'output' ? 'ADD' : 'REMOVE';
    var location = plan.location || (plan.type === 'output' ? FLOWS.MO_OUTPUT_LOCATION : FLOWS.MO_INGREDIENT_LOCATION);
    var result = null;
    var failureReason = plan.unresolvedReason || '';

    if (!failureReason && plan.type === 'output' && !plan.lot) {
      failureReason = 'Katana output batch still missing';
    }
    if (!failureReason && plan.type !== 'output' && plan.batchTracked && !plan.lot) {
      failureReason = 'Katana batch still missing on MO ingredient';
    }

    if (!failureReason) {
      markSyncedToWasp(plan.sku, location, action === 'ADD' ? 'add' : 'remove');
      var notes = 'MO debug repair ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
      if (action === 'ADD') {
        if (plan.lot || plan.expiry) {
          result = waspAddInventoryWithLot(plan.sku, plan.qty, location, plan.lot, plan.expiry, notes, CONFIG.WASP_SITE);
        } else {
          result = waspAddInventory(plan.sku, plan.qty, location, notes, CONFIG.WASP_SITE);
        }
      } else {
        if (plan.lot) {
          result = waspRemoveInventoryWithLot(plan.sku, plan.qty, location, plan.lot, notes, CONFIG.WASP_SITE, plan.expiry || '');
        } else {
          result = waspRemoveInventory(plan.sku, plan.qty, location, notes, CONFIG.WASP_SITE);
        }
      }

      if (!result.success) {
        failureReason = cleanErrorMessage(parseWaspError(result.response, action, plan.sku, location) || result.error || 'Unknown error');
      }
    }

    executed.push({
      key: item.key,
      parsed: item.parsed,
      plan: plan,
      success: !failureReason && !!(result && result.success),
      successStatus: plan.type === 'output' ? 'Produced' : 'Consumed',
      error: failureReason,
      result: result
    });
  }

  return executed;
}

function cloneMORepairResultQueues_(results) {
  var queues = {};
  for (var i = 0; i < results.length; i++) {
    var res = results[i];
    if (!queues[res.key]) queues[res.key] = [];
    queues[res.key].push(res);
  }
  return queues;
}

function rewriteMORepairDetailLine_(originalLine, plan) {
  var line = String(originalLine || '');
  line = line.replace(/\s+(PRODUCTION|PROD-RECEIVING|RECEIVING-DOCK|SHOPIFY|SW-STORAGE)\b.*$/, '');
  if (plan.location) line += '  ' + plan.location;
  if (!plan.lot && plan.batchId) line += '  batch_id:' + plan.batchId;
  if (plan.lot) line += '  lot:' + plan.lot;
  if (plan.expiry) line += '  exp:' + normalizeBusinessDate_(plan.expiry);
  return line;
}

function getRichTextLinkSpans_(cell) {
  var spans = [];
  var richText = cell.getRichTextValue();
  if (!richText || !richText.getRuns) return spans;
  var runs = richText.getRuns() || [];
  var cursor = 0;

  for (var i = 0; i < runs.length; i++) {
    var run = runs[i];
    var text = run.getText ? String(run.getText() || '') : '';
    var url = run.getLinkUrl ? run.getLinkUrl() : '';
    if (url) {
      spans.push({
        start: cursor,
        end: cursor + text.length,
        url: url
      });
    }
    cursor += text.length;
  }

  return spans;
}

function setTextWithLinkSpans_(cell, text, spans) {
  text = String(text || '');
  spans = spans || [];

  if (!spans.length) {
    cell.setValue(text);
    return;
  }

  var builder = SpreadsheetApp.newRichTextValue().setText(text);
  for (var i = 0; i < spans.length; i++) {
    var span = spans[i];
    if (!span.url) continue;
    if (span.start >= text.length) continue;
    builder.setLinkUrl(span.start, Math.min(span.end, text.length), span.url);
  }
  cell.setRichTextValue(builder.build());
}

function setMORepairDetailCellText_(cell, text, qtyColor, linkSpans) {
  text = String(text || '');
  linkSpans = linkSpans || [];
  var builder = SpreadsheetApp.newRichTextValue().setText(text);

  for (var i = 0; i < linkSpans.length; i++) {
    var span = linkSpans[i];
    if (!span.url) continue;
    if (span.start >= text.length) continue;
    builder.setLinkUrl(span.start, Math.min(span.end, text.length), span.url);
  }

  if (qtyColor) {
    var qtyMatch = text.match(/x[\d.]+/);
    if (qtyMatch && qtyMatch.index >= 0) {
      var qtyHex = qtyColor === 'green' ? '#008000' : qtyColor === 'grey' ? '#999999' : '#cc0000';
      var qtyStyle = SpreadsheetApp.newTextStyle().setForegroundColor(qtyHex).build();
      builder.setTextStyle(qtyMatch.index, qtyMatch.index + qtyMatch[0].length, qtyStyle);
    }
  }

  cell.setRichTextValue(builder.build());
}

function styleMORepairSubRow_(sheet, row, status, error) {
  var detailCell = sheet.getRange(row, 4);
  var statusCell = sheet.getRange(row, 5);
  var errorCell = sheet.getRange(row, 6);
  var normStatus = String(status || '').trim().toLowerCase();
  var normError = String(error || '').trim();

  if (normStatus === 'skipped' || normStatus === 'skip-ok') {
    detailCell.setBackground('#fff8e1');
  } else if (normError) {
    detailCell.setBackground('#ffebee');
  } else {
    detailCell.setBackground('#e8f5e9');
  }

  if (normStatus === 'failed') {
    statusCell.setBackground('#f8d7da');
  } else if (normStatus === 'skipped') {
    statusCell.setBackground('#fff8e1');
  } else if (normStatus) {
    statusCell.setBackground('#d4edda');
  }

  if (normError) {
    errorCell.setBackground('#fff0f0');
  } else {
    errorCell.setBackground('#ffffff');
  }
}

function styleMORepairHeaderRow_(sheet, headerRow) {
  var statusCell = sheet.getRange(headerRow, 5);
  var errorCell = sheet.getRange(headerRow, 6);
  var status = String(statusCell.getValue() || '').trim().toLowerCase();
  var error = String(errorCell.getValue() || '').trim();

  if (status === 'partial') {
    statusCell.setBackground('#fff3cd');
  } else if (status === 'failed') {
    statusCell.setBackground('#f8d7da');
  } else if (status) {
    statusCell.setBackground('#d4edda');
  }

  if (error) {
    errorCell.setBackground('#fff0f0');
  } else {
    errorCell.setBackground('#ffffff');
  }
}

function getF4BlockEndRow_(sheet, headerRow) {
  var lastRow = sheet.getLastRow();
  for (var row = headerRow + 1; row <= lastRow; row++) {
    var execId = String(sheet.getRange(row, 1).getValue() || '').trim();
    if (execId.indexOf('WK-') === 0) return row - 1;
  }
  return lastRow;
}

function deleteDuplicateMORepairBlocks_(sheet, headerRows, keepHeaderRow) {
  if (!sheet || !headerRows || headerRows.length <= 1) return 0;
  var deleted = 0;
  var rowsToDelete = [];

  for (var i = 0; i < headerRows.length; i++) {
    if (headerRows[i] === keepHeaderRow) continue;
    rowsToDelete.push({
      start: headerRows[i],
      end: getF4BlockEndRow_(sheet, headerRows[i])
    });
  }

  rowsToDelete.sort(function(a, b) { return b.start - a.start; });
  for (var j = 0; j < rowsToDelete.length; j++) {
    var block = rowsToDelete[j];
    var count = block.end - block.start + 1;
    if (count <= 0) continue;
    sheet.deleteRows(block.start, count);
    deleted += count;
  }

  return deleted;
}

function updateMORepairHeaderSummary_(sheet, headerRow) {
  var lastRow = sheet.getLastRow();
  var row = headerRow + 1;
  var failCount = 0;
  var skipCount = 0;

  while (row <= lastRow) {
    var execId = String(sheet.getRange(row, 1).getValue() || '').trim();
    if (execId.indexOf('WK-') === 0) break;
    var detail = String(sheet.getRange(row, 4).getValue() || '').trim();
    if (!detail) { row++; continue; }
    var status = String(sheet.getRange(row, 5).getValue() || '').trim();
    if (status === 'Failed') failCount++;
    if (status === 'Skipped') skipCount++;
    row++;
  }

  var detailCell = sheet.getRange(headerRow, 4);
  var linkSpans = getRichTextLinkSpans_(detailCell);
  var detailText = String(detailCell.getValue() || '');
  detailText = detailText.replace(/\s+\d+\s+errors?\b/g, '');
  if (failCount > 0) {
    detailText += '  ' + failCount + ' error' + (failCount > 1 ? 's' : '');
  }
  setTextWithLinkSpans_(detailCell, detailText.trim(), linkSpans);

  if (failCount === 0 && skipCount === 0) {
    sheet.getRange(headerRow, 6).setValue('');
  } else {
    var parts = [];
    if (skipCount > 0) parts.push(skipCount + ' skipped');
    if (failCount > 0) parts.push(failCount + ' failed');
    sheet.getRange(headerRow, 6).setValue(parts.join(', '));
  }

  styleMORepairHeaderRow_(sheet, headerRow);
}

function applyMORepairResultsToSheet_(sheet, headerRow, outputSku, executedResults) {
  if (!sheet || !headerRow) return { updatedRows: 0 };
  var lastRow = sheet.getLastRow();
  var queues = cloneMORepairResultQueues_(executedResults);
  var updatedRows = 0;
  var parentContext = null;

  for (var row = headerRow + 1; row <= lastRow; row++) {
    var execId = String(sheet.getRange(row, 1).getValue() || '').trim();
    if (execId.indexOf('WK-') === 0) break;

    var detail = String(sheet.getRange(row, 4).getValue() || '').trim();
    if (!detail) continue;

    var parsedAny = parseMORepairSubItem_(detail, parentContext);
    if (parsedAny && parsedAny.sku && !isQtyOnlyMORepairLine_(detail)) {
      parentContext = {
        sku: parsedAny.sku,
        location: parsedAny.location || ''
      };
    }

    var status = String(sheet.getRange(row, 5).getValue() || '').trim();
    if (status !== 'Failed' && status !== 'Skipped') continue;

    var parsed = parseMORepairSubItem_(detail, parentContext);
    if (!parsed || !parsed.sku) continue;

    var isOutput = parsed.sku === outputSku;
    var key = buildMORepairMatchKey_(
      isOutput,
      parsed.sku,
      parsed.qty,
      parsed.location,
      parsed.lot,
      parsed.expiry
    );
    var queue = queues[key] || [];
    if (!queue.length) continue;

    var repair = queue.shift();
    queues[key] = queue;
    var detailCell = sheet.getRange(row, 4);

    if (repair.success) {
      var rebuilt = rewriteMORepairDetailLine_(detail, repair.plan);
      setMORepairDetailCellText_(detailCell, rebuilt, isOutput ? 'green' : 'red');
      sheet.getRange(row, 5).setValue(repair.successStatus);
      sheet.getRange(row, 6).setValue('');
      styleMORepairSubRow_(sheet, row, repair.successStatus, '');
    } else {
      if (repair.plan && (repair.plan.location || repair.plan.lot || repair.plan.expiry)) {
        setMORepairDetailCellText_(detailCell, rewriteMORepairDetailLine_(detail, repair.plan), isOutput ? 'green' : 'red');
      }
      sheet.getRange(row, 6).setValue(repair.error || '');
      styleMORepairSubRow_(sheet, row, status, repair.error || '');
    }
    updatedRows++;
  }

  var lastCol = Math.max(6, sheet.getLastColumn());
  var data = sheet.getRange(1, 1, sheet.getLastRow(), Math.min(7, lastCol)).getValues();
  var execIdText = String(sheet.getRange(headerRow, 1).getValue() || '').trim();
  var touchedHeaders = {};
  touchedHeaders[execIdText] = { execId: execIdText, flow: 'F4 Manufacturing' };
  updateRetryHeaderStatuses(sheet, data, touchedHeaders);
  updateMORepairHeaderSummary_(sheet, headerRow);
  styleMORepairHeaderRow_(sheet, headerRow);

  return { updatedRows: updatedRows };
}

function collectFailedMORepairSourceRows_(sheet, headerRow, outputSku) {
  var rows = [];
  if (!sheet || !headerRow) return rows;
  var lastRow = sheet.getLastRow();
  var parentContext = null;

  for (var row = headerRow + 1; row <= lastRow; row++) {
    var execId = String(sheet.getRange(row, 1).getValue() || '').trim();
    if (execId.indexOf('WK-') === 0) break;

    var detail = String(sheet.getRange(row, 4).getValue() || '').trim();
    var parsedAny = parseMORepairSubItem_(detail, parentContext);
    if (parsedAny && parsedAny.sku && !isQtyOnlyMORepairLine_(detail)) {
      parentContext = {
        sku: parsedAny.sku,
        location: parsedAny.location || ''
      };
    }

    var status = String(sheet.getRange(row, 5).getValue() || '').trim();
    if (status !== 'Failed' && status !== 'Skipped') continue;

    var parsed = parseMORepairSubItem_(detail, parentContext);
    if (!parsed || !parsed.sku) continue;

    rows.push({
      rowIndex: row,
      detail: detail,
      parsed: parsed,
      isOutput: parsed.sku === outputSku
    });
  }

  return rows;
}

function repairMOFromDebugForOrder_(moOrderNo) {
  moOrderNo = String(moOrderNo || '').trim();
  if (!moOrderNo) throw new Error('No MO order number provided');

  var batchData = collectMOBatchDataForOrder_(moOrderNo);
  var outputRepair = resolveMOOutputRepairData_(batchData.mo);

  if (!outputRepair.lot) {
    Utilities.sleep(5000);
    batchData = collectMOBatchDataForOrder_(moOrderNo);
    outputRepair = resolveMOOutputRepairData_(batchData.mo);
  }

  writeMOBatchDiagnosticSheet_(batchData.moOrderNo, batchData.mo, batchData.rowsWithInclude, batchData.rowsWithoutInclude);
  var overrideMap = readMOBatchDebugOverridesFromSheet_(SpreadsheetApp.openById(CONFIG.DEBUG_SHEET_ID).getSheetByName('MO Debug'));

  var ss = SpreadsheetApp.openById(CONFIG.DEBUG_SHEET_ID);
  var activitySheet = ss.getSheetByName('Activity');
  var flowSheet = ss.getSheetByName(FLOW_TAB_NAMES.F4);
  var activityHeaderRows = findAllF4HeaderRowsForOrder_(activitySheet, batchData.moOrderNo, outputRepair.sku);
  var flowHeaderRows = findAllF4HeaderRowsForOrder_(flowSheet, batchData.moOrderNo, outputRepair.sku);
  var activityHeaderRow = activityHeaderRows.length ? activityHeaderRows[0] : 0;
  var flowHeaderRow = flowHeaderRows.length ? flowHeaderRows[0] : 0;
  if (!activityHeaderRows.length && !flowHeaderRows.length) {
    throw new Error('No F4 Activity row found for ' + batchData.moOrderNo);
  }

  var activitySourceRows = [];
  for (var ah = 0; ah < activityHeaderRows.length; ah++) {
    var activityCandidateRows = collectFailedMORepairSourceRows_(activitySheet, activityHeaderRows[ah], outputRepair.sku);
    if (activityCandidateRows.length > activitySourceRows.length) activitySourceRows = activityCandidateRows;
  }
  var flowSourceRows = [];
  for (var fh = 0; fh < flowHeaderRows.length; fh++) {
    var flowCandidateRows = collectFailedMORepairSourceRows_(flowSheet, flowHeaderRows[fh], outputRepair.sku);
    if (flowCandidateRows.length > flowSourceRows.length) flowSourceRows = flowCandidateRows;
  }
  var sourceRows = flowSourceRows.length > activitySourceRows.length ? flowSourceRows : activitySourceRows;
  if (!sourceRows.length) {
    var activityDeletedNoop = deleteDuplicateMORepairBlocks_(activitySheet, activityHeaderRows, activityHeaderRow);
    var flowDeletedNoop = deleteDuplicateMORepairBlocks_(flowSheet, flowHeaderRows, flowHeaderRow);
    return {
      moOrderNo: batchData.moOrderNo,
      repaired: 0,
      failed: 0,
      activityUpdated: 0,
      flowUpdated: 0,
      activityDeleted: activityDeletedNoop,
      flowDeleted: flowDeletedNoop,
      message: 'No failed or skipped F4 sub-items found to repair'
    };
  }

  var repairContext = {
    output: outputRepair,
    ingredientsBySku: buildMOIngredientRepairQueues_(batchData.rowsWithInclude, overrideMap)
  };
  var planRows = buildMORepairPlanFromSourceRows_(sourceRows, repairContext);
  var lotDateRepair = trySafeMODateCodeRepairs_(planRows, batchData.moOrderNo);
  var executed = executeMORepairPlan_(planRows);

  var repaired = 0;
  var failed = 0;
  var failedDetails = [];
  for (var i = 0; i < executed.length; i++) {
    if (executed[i].success) repaired++;
    else {
      failed++;
      failedDetails.push(
        String(executed[i].plan && executed[i].plan.sku || executed[i].parsed && executed[i].parsed.sku || 'UNKNOWN') +
        ': ' +
        String(executed[i].error || 'Unknown error')
      );
    }
  }

  var activityApplied = activityHeaderRow
    ? applyMORepairResultsToSheet_(activitySheet, activityHeaderRow, outputRepair.sku, executed)
    : { updatedRows: 0 };
  var flowApplied = flowHeaderRow
    ? applyMORepairResultsToSheet_(flowSheet, flowHeaderRow, outputRepair.sku, executed)
    : { updatedRows: 0 };
  var activityDeleted = deleteDuplicateMORepairBlocks_(activitySheet, activityHeaderRows, activityHeaderRow);
  var flowDeleted = deleteDuplicateMORepairBlocks_(flowSheet, flowHeaderRows, flowHeaderRow);

  return {
    moOrderNo: batchData.moOrderNo,
    lotDatesRepaired: lotDateRepair.repaired,
    lotDateRepairsSkipped: lotDateRepair.skipped,
    repaired: repaired,
    failed: failed,
    failedDetails: failedDetails,
    activityUpdated: activityApplied.updatedRows,
    flowUpdated: flowApplied.updatedRows,
    activityDeleted: activityDeleted,
    flowDeleted: flowDeleted,
    rowsWithInclude: batchData.rowsWithInclude.length,
    lotDateRepairDetails: lotDateRepair.details
  };
}

function repairMOFromDebugPrompt() {
  var ui = SpreadsheetApp.getUi();
  var moOrderNo = getMOBatchDebugOrderNumber_();

  if (!moOrderNo) {
    var response = ui.prompt('Repair MO From Debug', 'Enter the Katana MO order number (for example: MO-7407 TEST)', ui.ButtonSet.OK_CANCEL);
    if (response.getSelectedButton() !== ui.Button.OK) return;
    moOrderNo = String(response.getResponseText() || '').trim();
  }

  if (!moOrderNo) {
    ui.alert('Enter an MO order number.');
    return;
  }

  var result = repairMOFromDebugForOrder_(moOrderNo);
  ui.alert(
    'MO repair complete for ' + result.moOrderNo +
    '\nLot dates repaired: ' + result.lotDatesRepaired +
    '\nRepaired: ' + result.repaired +
    '\nStill failed: ' + result.failed +
    (result.failedDetails && result.failedDetails.length ? '\nFailed items: ' + result.failedDetails.join(' | ') : '') +
    '\nActivity rows updated: ' + result.activityUpdated +
    '\nF4 rows updated: ' + result.flowUpdated +
    '\nDuplicate rows removed: ' + ((result.activityDeleted || 0) + (result.flowDeleted || 0))
  );
}

function testMOBatchData() {
  // ---- CHANGE THIS to the MO you want to inspect ----
  var moOrderNo = 'MO-7255';
  // ---------------------------------------------------

  debugMOBatchDataForOrder_(moOrderNo);
  return;

  Logger.log('=== MO BATCH DATA DIAGNOSTIC ===');
  Logger.log('Target: ' + moOrderNo);

  // Step 1: Find the MO by scanning recent MOs
  var moId = null;
  var mo = null;
  var searchTarget = moOrderNo.replace(/^MO-/i, ''); // strip MO- prefix for flexible matching

  var searchResult = katanaApiCall('manufacturing_orders?limit=250&page=1');
  var mos = searchResult && searchResult.data ? searchResult.data : (searchResult || []);

  // Log first 5 order_no values so we can see the format
  Logger.log('Sample order_no values from API:');
  for (var s = 0; s < Math.min(5, mos.length); s++) {
    Logger.log('  [' + s + '] id=' + mos[s].id + '  order_no="' + mos[s].order_no + '"  status=' + mos[s].status);
  }
  Logger.log('Page 1 returned ' + mos.length + ' MOs');

  // Search with flexible matching (with or without MO- prefix)
  for (var page = 1; page <= 5 && !moId; page++) {
    if (page > 1) {
      searchResult = katanaApiCall('manufacturing_orders?limit=250&page=' + page);
      mos = searchResult && searchResult.data ? searchResult.data : (searchResult || []);
      Logger.log('Page ' + page + ' returned ' + mos.length + ' MOs');
    }
    if (mos.length === 0) break;

    for (var m = 0; m < mos.length; m++) {
      var orderNo = String(mos[m].order_no || '').trim();
      // Match if order_no starts with target (handles "MO-7255 TEST DELETE ME" etc.)
      if (orderNo === moOrderNo || orderNo.indexOf(moOrderNo) === 0) {
        mo = mos[m];
        moId = mo.id;
        break;
      }
    }
  }

  if (!moId) {
    Logger.log('ERROR: MO ' + moOrderNo + ' not found. Check the sample order_no format above.');
    return;
  }
  Logger.log('Found MO — ID: ' + moId + ', status: ' + mo.status + ', order_no: ' + mo.order_no);
  Logger.log('MO batch_number: ' + (mo.batch_number || '(none)'));
  Logger.log('MO variant_id: ' + mo.variant_id);

  // Step 2: Fetch recipe rows WITH include=batch_transactions
  Logger.log('\n--- Recipe rows WITH &include=batch_transactions ---');
  var withInclude = katanaApiCall('manufacturing_order_recipe_rows?manufacturing_order_id=' + moId + '&include=batch_transactions');
  var rowsA = withInclude && withInclude.data ? withInclude.data : (withInclude || []);
  Logger.log('Rows returned: ' + rowsA.length);

  for (var i = 0; i < rowsA.length; i++) {
    var row = rowsA[i];
    Logger.log('\n  Row ' + i + ' keys: ' + Object.keys(row).join(', '));
    Logger.log('  Row ' + i + ' variant_id: ' + row.variant_id);
    Logger.log('  Row ' + i + ' quantity: ' + row.quantity);
    Logger.log('  Row ' + i + ' consumed_quantity: ' + (row.consumed_quantity || row.total_consumed_quantity || '(none)'));
    Logger.log('  Row ' + i + ' batch_number (direct): ' + (row.batch_number || row.batch_nr || '(none)'));
    Logger.log('  Row ' + i + ' batch_transactions: ' + JSON.stringify(row.batch_transactions || '(not present)'));
    Logger.log('  Row ' + i + ' stock_allocations: ' + JSON.stringify(row.stock_allocations || '(not present)'));
    Logger.log('  Row ' + i + ' picked_batches: ' + JSON.stringify(row.picked_batches || '(not present)'));

    // Resolve SKU for readability
    if (row.variant_id) {
      var v = fetchKatanaVariant(row.variant_id);
      var vd = v && v.data ? v.data : v;
      var sku = vd ? (vd.sku || '') : '';
      Logger.log('  Row ' + i + ' SKU: ' + sku);
    }

    // If batch_transactions has entries, try resolving batch_id
    var bt = row.batch_transactions || [];
    if (bt.length > 0) {
      for (var b = 0; b < bt.length; b++) {
        Logger.log('  batch_transactions[' + b + '] keys: ' + Object.keys(bt[b]).join(', '));
        Logger.log('  batch_transactions[' + b + '] full: ' + JSON.stringify(bt[b]));
        var bId = bt[b].batch_id || bt[b].batchId || null;
        if (bId) {
          Logger.log('  Resolving batch_id ' + bId + ' via /batch_stocks/' + bId + '...');
          var batchInfo = fetchKatanaBatchStock(bId);
          Logger.log('  batch_stocks result: ' + JSON.stringify(batchInfo));
        }
      }
    }
  }

  // Step 3: Fetch recipe rows WITHOUT include (baseline comparison)
  Logger.log('\n--- Recipe rows WITHOUT &include (baseline) ---');
  var withoutInclude = katanaApiCall('manufacturing_order_recipe_rows?manufacturing_order_id=' + moId);
  var rowsB = withoutInclude && withoutInclude.data ? withoutInclude.data : (withoutInclude || []);
  Logger.log('Rows returned: ' + rowsB.length);
  for (var j = 0; j < rowsB.length; j++) {
    Logger.log('  Row ' + j + ' keys: ' + Object.keys(rowsB[j]).join(', '));
    Logger.log('  Row ' + j + ' batch_transactions: ' + JSON.stringify(rowsB[j].batch_transactions || '(not present)'));
  }

  // Step 4: For each ingredient variant, check batch_stocks directly
  Logger.log('\n--- Direct batch_stocks lookup per variant ---');
  var checkedVariants = {};
  for (var k = 0; k < rowsA.length; k++) {
    var vid = rowsA[k].variant_id;
    if (!vid || checkedVariants[vid]) continue;
    checkedVariants[vid] = true;

    var batchStocks = katanaApiCall('batch_stocks?variant_id=' + vid);
    var batches = batchStocks && batchStocks.data ? batchStocks.data : (batchStocks || []);
    Logger.log('  variant_id ' + vid + ': ' + batches.length + ' batch_stocks found');
    for (var bs = 0; bs < batches.length; bs++) {
      Logger.log('    batch ' + bs + ': ' + JSON.stringify(batches[bs]));
    }
  }

  Logger.log('\n=== DIAGNOSTIC COMPLETE ===');
  Logger.log('Check the output above to see if batch_transactions are populated.');
}

/**
 * DIAGNOSTIC: Test WASP lot lookup for a specific item
 * Run from script editor: testWaspLotLookup()
 * Check View → Logs for output.
 */
function testWaspLotLookup() {
  var itemNumber = 'LUP-1';
  var location = 'PRODUCTION';

  Logger.log('=== WASP LOT LOOKUP DIAGNOSTIC ===');
  Logger.log('Item: ' + itemNumber + '  Location: ' + location);

  // Test 1: Paginate advancedinventorysearch to find item across ALL pages
  Logger.log('\n--- Test 1: advancedinventorysearch (paginated scan) ---');
  var url1 = CONFIG.WASP_BASE_URL + '/public-api/ic/item/advancedinventorysearch';
  var found = false;
  for (var pg = 1; pg <= 20; pg++) {
    var result1 = waspApiCall(url1, { SearchPattern: itemNumber, PageSize: 100, PageNumber: pg });
    if (!result1.success) { Logger.log('Page ' + pg + ' FAILED'); break; }
    var resp = JSON.parse(result1.response);
    var rows = resp.Data || resp.data || [];
    if (rows.length === 0) { Logger.log('Page ' + pg + ' empty — end of data'); break; }
    Logger.log('Page ' + pg + ': ' + rows.length + ' rows');
    for (var i = 0; i < rows.length; i++) {
      var rowItem = rows[i].ItemNumber || rows[i].itemNumber || '';
      if (rowItem === itemNumber) {
        Logger.log('  FOUND on page ' + pg + ' row ' + i + ': ' + JSON.stringify(rows[i]));
        found = true;
      }
    }
    if (found) break;
    if (rows.length < 100) { Logger.log('Last page (< 100 rows)'); break; }
  }
  if (!found) Logger.log('NOT FOUND in any page — item has no WASP inventory');

  // Test 2: waspLookupItemLots() (now with pagination)
  Logger.log('\n--- Test 2: waspLookupItemLots() at ' + location + ' ---');
  var lotResult = waspLookupItemLots(itemNumber, location);
  Logger.log('Lot found: ' + (lotResult || 'null'));

  // Test 3: waspLookupItemLotAndDate() (now with pagination)
  Logger.log('\n--- Test 3: waspLookupItemLotAndDate() at ' + location + ' ---');
  var lotAndDate = waspLookupItemLotAndDate(itemNumber, location);
  Logger.log('Result: ' + JSON.stringify(lotAndDate || 'null'));

  Logger.log('\n=== DIAGNOSTIC COMPLETE ===');
}

/**
 * Test real WASP quantity addition
 */
function testRealAdd() {
  Logger.log('[ignore] REAL_TEST_START');

  var payload = {
    source: 'WASP',
    event: 'quantity_added',
    AssetTag: 'B-WAX',
    Quantity: 1,
    LocationCode: 'PRODUCTION',
    SiteName: 'MMH Kelowna'
  };

  var result = handleWaspQuantityAdded(payload);

  Logger.log('[skip] REAL_TEST_RESULT');
  Logger.log(JSON.stringify(result));

  return result;
}

// ============================================
// WEEKLY ARCHIVE — ACTIVITY + FLOW TABS
// ============================================
// Runs every Sunday 11 PM via time trigger.
// 1. Copies entire spreadsheet to Drive archive folder
// 2. Clears data rows from Activity + F1-F5 tabs (keeps headers rows 1-3)
// 3. Resets exec ID counter → WK-001 on next event
// StockTransfers tab is NOT cleared.
// ============================================

/**
 * Short month names for archive file naming
 */
var MONTH_NAMES = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun',
                   'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];

/**
 * Weekly archive — copy spreadsheet to Drive, clear logs, reset counter.
 * Triggered every Sunday 11 PM, or run manually from GAS editor.
 */
function weeklyArchive() {
  var tz = Session.getScriptTimeZone();
  var now = new Date();

  // Calculate week range: Monday (6 days ago) through Sunday (today)
  var sunday = new Date(now);
  var monday = new Date(now);
  monday.setDate(sunday.getDate() - 6);

  // Week number (ISO-like: week of the Monday date)
  var jan1 = new Date(monday.getFullYear(), 0, 1);
  var daysSinceJan1 = Math.floor((monday - jan1) / 86400000);
  var weekNum = Math.ceil((daysSinceJan1 + jan1.getDay() + 1) / 7);
  var weekStr = weekNum < 10 ? 'W0' + weekNum : 'W' + weekNum;

  // Build file name: "WK-Archive 2026-W09 (Feb 24 - Mar 02)"
  var monDay = monday.getDate();
  var monMonth = MONTH_NAMES[monday.getMonth()];
  var sunDay = sunday.getDate();
  var sunMonth = MONTH_NAMES[sunday.getMonth()];
  var monDayStr = monDay < 10 ? '0' + monDay : '' + monDay;
  var sunDayStr = sunDay < 10 ? '0' + sunDay : '' + sunDay;

  var archiveName = 'WK-Archive ' + monday.getFullYear() + '-' + weekStr
    + ' (' + monMonth + ' ' + monDayStr + ' - ' + sunMonth + ' ' + sunDayStr + ')';

  // 1. Copy spreadsheet to archive folder
  var archiveFolderId = PropertiesService.getScriptProperties().getProperty('ARCHIVE_FOLDER_ID');
  if (!archiveFolderId) {
    Logger.log('weeklyArchive ERROR: ARCHIVE_FOLDER_ID not set. Run setArchiveFolderId() first.');
    return { success: false, error: 'ARCHIVE_FOLDER_ID not configured' };
  }

  var folder = DriveApp.getFolderById(archiveFolderId);
  var file = DriveApp.getFileById(CONFIG.DEBUG_SHEET_ID);
  file.makeCopy(archiveName, folder);

  Logger.log('weeklyArchive: copied to "' + archiveName + '"');

  // 2. Clear data rows from Activity + F1-F5 (keep header rows 1-3)
  var ss = SpreadsheetApp.openById(CONFIG.DEBUG_SHEET_ID);
  var tabsToClear = [
    'Activity',
    'F1 Receiving',
    'F2 Adjustments',
    'F3 Transfers',
    'F4 Manufacturing',
    'F5 Shipping',
    'WebhookQueue'
  ];

  var totalRowsCleared = 0;
  for (var t = 0; t < tabsToClear.length; t++) {
    var sheet = ss.getSheetByName(tabsToClear[t]);
    if (!sheet) continue;
    var lastRow = sheet.getLastRow();
    if (lastRow > 3) {
      var rowCount = lastRow - 3;
      sheet.getRange(4, 1, rowCount, sheet.getMaxColumns()).clear();
      totalRowsCleared += rowCount;
    }
  }

  // 3. Reset exec ID counter — next getNextExecId() starts at WK-001
  PropertiesService.getScriptProperties().setProperty('EXEC_ID_COUNTER', '0');

  Logger.log('weeklyArchive complete: ' + totalRowsCleared + ' rows cleared, counter reset');

  return { success: true, archive: archiveName, rowsCleared: totalRowsCleared };
}

/**
 * Set up weekly archive trigger — Sunday 11 PM.
 * Run once from GAS editor after deploy.
 */
function setupWeeklyArchiveTrigger() {
  // Delete existing triggers for this function
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'weeklyArchive') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }

  // Create Sunday 11 PM trigger
  ScriptApp.newTrigger('weeklyArchive')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.SUNDAY)
    .atHour(23)
    .create();

  Logger.log('weeklyArchive trigger created: Sunday 11 PM');
}

/**
 * Store archive folder ID in ScriptProperties.
 * Run once from GAS editor: setArchiveFolderId('your-folder-id-here')
 * Get folder ID from Drive URL: https://drive.google.com/drive/folders/{FOLDER_ID}
 */
function setArchiveFolderId(folderId) {
  if (!folderId) {
    Logger.log('Usage: setArchiveFolderId("your-folder-id-from-drive-url")');
    return;
  }
  PropertiesService.getScriptProperties().setProperty('ARCHIVE_FOLDER_ID', folderId);
  Logger.log('ARCHIVE_FOLDER_ID set to: ' + folderId);
}

// ============================================
// ADJUSTMENTS LOG TAB SETUP
// ============================================

/**
 * Create (or fully reset) the "Adjustments Log" tab in the debug sheet.
 * Safe to re-run: deletes and rebuilds the tab if it already exists.
 * Run once from the GAS editor after deploying.
 */
function setupAdjustmentsLogTab() {
  var ss = SpreadsheetApp.openById(CONFIG.DEBUG_SHEET_ID);

  // Remove existing tab if present
  var existing = ss.getSheetByName('F2 Adjustments');
  if (existing) {
    ss.deleteSheet(existing);
    Logger.log('Deleted existing "Adjustments Log" tab');
  }

  // Also remove old "Adjustments" tab if it still exists from before the rename
  var oldTab = ss.getSheetByName('Adjustments');
  if (oldTab) {
    ss.deleteSheet(oldTab);
    Logger.log('Deleted old "Adjustments" tab');
  }

  // Create fresh tab
  var sheet = ss.insertSheet('F2 Adjustments');
  SpreadsheetApp.flush(); // Let the sheet creation settle before formatting

  var headers   = ['Timestamp', 'Source', 'Action', 'User', 'SKU', 'Item Name', 'Site', 'Location', 'Lot / Batch', 'Expiry', 'Diff', 'Katana SA#', 'Status'];
  var colWidths = [155, 100, 80, 120, 120, 180, 120, 130, 120, 90, 70, 130, 80];

  // Header row
  sheet.getRange(1, 1, 1, 13).setValues([headers]);
  sheet.getRange(1, 1, 1, 13).setBackground('#1F2937').setFontColor('#ffffff').setFontWeight('bold');
  sheet.setFrozenRows(1);
  SpreadsheetApp.flush(); // Flush header writes before column widths + formatting

  // Column widths
  for (var c = 0; c < colWidths.length; c++) {
    sheet.setColumnWidth(c + 1, colWidths[c]);
  }

  // Conditional formatting
  var fullRange   = sheet.getRange(2, 1, 998, 13); // whole rows A:M
  var diffRange   = sheet.getRange(2, 11, 998, 1); // col K  Diff
  var statusRange = sheet.getRange(2, 13, 998, 1); // col M  Status
  var rules = [];

  // Row background by Source (col B): WASP=blue, Sync/Google Sheet=yellow
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$B2="WASP"')         .setBackground('#dbeafe').setRanges([fullRange]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$B2="Sync Sheet"')   .setBackground('#fef9c3').setRanges([fullRange]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$B2="Google Sheet"') .setBackground('#fef9c3').setRanges([fullRange]).build());

  // Status cell color
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('OK')   .setBackground('#c8e6c9').setRanges([statusRange]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('ERROR').setBackground('#ffcdd2').setRanges([statusRange]).build());

  // Diff cell color
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThan(0).setBackground('#c8e6c9').setFontColor('#2e7d32').setRanges([diffRange]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenNumberLessThan(0)   .setBackground('#ffcdd2').setFontColor('#c62828').setRanges([diffRange]).build());

  sheet.setConditionalFormatRules(rules);

  Logger.log('"Adjustments Log" tab created successfully with 13 columns and conditional formatting.');
}

/**
 * Re-apply updated CF rules to the existing Adjustments Log tab without deleting data.
 * Run once from Apps Script editor: updateAdjustmentsLogCF()
 */
function updateAdjustmentsLogCF() {
  var ss = SpreadsheetApp.openById(CONFIG.DEBUG_SHEET_ID);
  var sheet = ss.getSheetByName('F2 Adjustments');
  if (!sheet) {
    Logger.log('Adjustments Log tab not found');
    return;
  }

  var fullRange   = sheet.getRange(2, 1, 998, 13);
  var diffRange   = sheet.getRange(2, 11, 998, 1);
  var statusRange = sheet.getRange(2, 13, 998, 1);
  var rules = [];

  // Row background by Source (col B): WASP=blue, Sync/Google Sheet=yellow
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$B2="WASP"')         .setBackground('#dbeafe').setRanges([fullRange]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$B2="Sync Sheet"')   .setBackground('#fef9c3').setRanges([fullRange]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$B2="Google Sheet"') .setBackground('#fef9c3').setRanges([fullRange]).build());

  // Status cell color
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('OK')   .setBackground('#c8e6c9').setRanges([statusRange]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('ERROR').setBackground('#ffcdd2').setRanges([statusRange]).build());

  // Diff cell color
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThan(0).setBackground('#c8e6c9').setFontColor('#2e7d32').setRanges([diffRange]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenNumberLessThan(0)   .setBackground('#ffcdd2').setFontColor('#c62828').setRanges([diffRange]).build());

  sheet.setConditionalFormatRules(rules);
  Logger.log('Adjustments Log CF updated.');
}

/**
 * Create (or fully reset) the "Webhook Queue" tab in the debug sheet.
 * Safe to re-run: deletes and rebuilds the tab if it already exists.
 * Run once from the GAS editor after deploying.
 */
function setupWebhookQueueTab() {
  var ss = SpreadsheetApp.openById(CONFIG.DEBUG_SHEET_ID);

  var existing = ss.getSheetByName('Webhook Queue');
  if (existing) {
    ss.deleteSheet(existing);
    Logger.log('Deleted existing "Webhook Queue" tab');
  }

  var sheet = ss.insertSheet('Webhook Queue');
  SpreadsheetApp.flush();

  var headers   = ['Timestamp', 'Action', 'Status', 'Result', 'Payload'];
  var colWidths = [160, 220, 80, 160, 460];

  sheet.getRange(1, 1, 1, 5).setValues([headers]);
  sheet.getRange(1, 1, 1, 5).setBackground('#1F2937').setFontColor('#ffffff').setFontWeight('bold');
  sheet.setFrozenRows(1);
  SpreadsheetApp.flush();

  for (var c = 0; c < colWidths.length; c++) {
    sheet.setColumnWidth(c + 1, colWidths[c]);
  }

  // Color Status column: green=ok/processed/reverted, yellow=skipped/ignored, red=error
  var statusRange = sheet.getRange(2, 3, 998, 1);
  var rules = [];
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('ok').setBackground('#c8e6c9').setRanges([statusRange]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('processed').setBackground('#c8e6c9').setRanges([statusRange]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('reverted').setBackground('#c8e6c9').setRanges([statusRange]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('skipped').setBackground('#fff3cd').setRanges([statusRange]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('ignored').setBackground('#fff3cd').setRanges([statusRange]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('error').setBackground('#ffcdd2').setRanges([statusRange]).build());
  sheet.setConditionalFormatRules(rules);

  Logger.log('"Webhook Queue" tab created successfully.');
}

/**
 * Test logWebhookQueue by writing a fake entry directly to the Webhook Queue tab.
 * Run this from the GAS editor to confirm the code works.
 * If a row appears in the tab, the function is working — only deployment is needed.
 */
function testLogWebhookQueue() {
  var fakePayload = { action: 'purchase_order.received', id: 'TEST-000', test: true };
  var fakeResult  = { status: 'ok', message: 'test entry' };
  logWebhookQueue(fakePayload, fakeResult);
  Logger.log('testLogWebhookQueue: done — check "Webhook Queue" tab for a TEST-000 row.');
}

// ============================================
// BULK COST SYNC: Katana → WASP
// ============================================

/**
 * Fetch the last non-zero average_cost_after from inventory_movements for a variant.
 * Used as fallback when current average_cost = 0 (item is out of stock).
 * inventory_movements returns newest-first by default.
 */
function fetchLastAverageCostFromMovements(variantId) {
  var result = katanaApiCall('inventory_movements?variant_id=' + variantId + '&limit=50');
  if (!result) return 0;
  var rows = Array.isArray(result) ? result : (result.data || []);
  // Pass 1: last non-zero average_cost_after (item currently out of stock but had stock recently)
  for (var i = 0; i < rows.length; i++) {
    var cost = parseFloat(rows[i].average_cost_after || 0);
    if (cost > 0) return cost;
  }
  // Pass 2: cost_after is 0 on all movements — use average_cost_before of the most recent
  // movement that had a value (captures cost just before it dropped to zero permanently)
  for (var j = 0; j < rows.length; j++) {
    var costBefore = parseFloat(rows[j].average_cost_before || 0);
    if (costBefore > 0) return costBefore;
  }
  return 0;
}

/**
 * Sync all item costs from Katana average cost → WASP.
 * Source: GET /inventory → average_cost field (dynamic, recalculates on every PO/MO).
 * Fallback: GET /inventory_movements → last non-zero average_cost_after (for out-of-stock items).
 *
 * Run manually or via daily trigger (setupCostSyncTrigger).
 */
function syncAllCostsKatanaToWasp() {
  var BATCH_SIZE  = 50;
  var DELAY_MS    = 300;

  Logger.log('=== syncAllCostsKatanaToWasp START ===');

  // ── 1. Build variantId → average_cost map from GET /inventory ─────────────
  var costMap = {};
  var invPage = 1;
  while (invPage <= 50) {
    var invResult = katanaApiCall('inventory?limit=250&page=' + invPage);
    if (!invResult) break;
    var invRows = Array.isArray(invResult) ? invResult : (invResult.data || []);
    if (!invRows || invRows.length === 0) break;
    for (var ii = 0; ii < invRows.length; ii++) {
      var ir = invRows[ii];
      if (!ir.variant_id) continue;
      var c = parseFloat(ir.average_cost || 0);
      // average_cost is company-wide (same across locations) — take highest seen
      if (!costMap[ir.variant_id] || c > costMap[ir.variant_id]) {
        costMap[ir.variant_id] = c;
      }
    }
    if (invRows.length < 250) break;
    invPage++;
  }
  Logger.log('Inventory cost map built: ' + Object.keys(costMap).length + ' variants');

  // ── 2. Fetch all variants → resolve SKU + cost ────────────────────────────
  var updates = [];
  var fallbackCount = 0;
  var varPage = 1;
  while (varPage <= 50) {
    var vResult = katanaApiCall('variants?limit=250&page=' + varPage);
    if (!vResult) break;
    var variants = Array.isArray(vResult) ? vResult : (vResult.data || []);
    if (!variants || variants.length === 0) break;

    for (var vi = 0; vi < variants.length; vi++) {
      var v = variants[vi];
      if (!v.sku || isSkippedSku(v.sku)) continue;

      var varCost = costMap[v.id] || 0;

      // Fallback: out-of-stock items have average_cost=0 — use last movement value
      if (varCost <= 0) {
        Utilities.sleep(500); // throttle fallback calls — avoid Katana 429
        varCost = fetchLastAverageCostFromMovements(v.id);
        if (varCost > 0) fallbackCount++;
      }

      if (varCost > 0) {
        // Apply UOM conversion factor if Katana and WASP use different units.
        // factor = multiply Katana qty by this to get WASP qty.
        // So cost per WASP unit = Katana cost / factor.
        // (e.g. Katana tracks in dozen at $3.60 → WASP cost per each = $3.60 / 12 = $0.30)
        var factor = getUomConversionFactor(v.sku);
        var waspCost = (factor !== 1) ? (varCost / factor) : varCost;

        updates.push({
          ItemNumber:    v.sku,
          Cost:          waspCost,
          // Do NOT include StockingUnit/PurchaseUnit/SalesUnit here —
          // cost sync should not overwrite UOM. Use bulkFixItemUomsAndCategories() for that.
          DimensionInfo: { DimensionUnit: 'Inch', WeightUnit: 'Pound', VolumeUnit: 'Cubic Inch' }
        });
      }
    }

    if (variants.length < 250) break;
    varPage++;
    Utilities.sleep(DELAY_MS);
  }

  Logger.log('Items to update: ' + updates.length + ' (' + fallbackCount + ' used movement fallback)');

  // ── 3. Batch update WASP ──────────────────────────────────────────────────
  var url     = CONFIG.WASP_BASE_URL + '/public-api/ic/item/updateInventoryItems';
  var success = 0;
  var failed  = 0;
  var skipped = 0; // items not in WASP (-64001) — excluded from failure count
  var errors  = [];

  for (var b = 0; b < updates.length; b += BATCH_SIZE) {
    var batch = updates.slice(b, b + BATCH_SIZE);
    var r = waspApiCall(url, batch);
    if (r.success) {
      success += batch.length;
    } else {
      // Batch failed — one bad item kills the whole batch.
      // Retry each item individually to isolate the bad ones.
      Logger.log('Batch ' + (Math.floor(b / BATCH_SIZE) + 1) + ' failed — retrying ' + batch.length + ' items individually');
      for (var ri = 0; ri < batch.length; ri++) {
        var singleR = waspApiCall(url, [batch[ri]]);
        if (singleR.success) {
          success++;
        } else {
          var errStr = String(singleR.response || '');
          // -64001 = "item does not exist in the system" — skip silently
          if (errStr.indexOf('-64001') > -1 || errStr.toLowerCase().indexOf('does not exist') > -1) {
            skipped++;
            Logger.log('SKIP (not in WASP): ' + batch[ri].ItemNumber);
          } else {
            failed++;
            errors.push(batch[ri].ItemNumber + ': ' + errStr.substring(0, 150));
          }
        }
        Utilities.sleep(200); // throttle individual retries
      }
    }
    if (b + BATCH_SIZE < updates.length) Utilities.sleep(DELAY_MS);
  }

  Logger.log('=== DONE — Updated: ' + success + ' | Skipped (not in WASP): ' + skipped + ' | Failed: ' + failed + ' ===');
  if (errors.length > 0) Logger.log('Errors:\n' + errors.join('\n'));

  return { success: success, failed: failed, skipped: skipped, total: updates.length, fallback: fallbackCount, errors: errors };
}

// ============================================
// COMBINED METADATA SYNC: Cost + UOM + Category (Katana → WASP)
// ============================================

/**
 * Sync cost, UOM, and category from Katana to WASP for all items in one pass.
 *
 * - Cost: Katana average_cost (live). Falls back to last non-zero inventory_movement
 *         when average_cost = 0 (item is out of stock).
 * - UOM: Katana product/material UOM written as-is (e.g. "PCS", "g", "kg").
 *         Only updated when Katana and WASP differ (case-insensitive comparison).
 *         Skipped for items in UOM_CONVERSIONS (intentional unit mismatch).
 * - Category: Katana product/material category. Skipped when Katana has no category.
 *
 * Run from GAS editor:
 *   syncAllItemMetadata()        ← dry run (preview, no changes)
 *   syncAllItemMetadata(false)   ← live (makes changes in WASP)
 */
function syncAllItemMetadata(dryRun, maxItems) {
  if (dryRun === undefined || dryRun === null) dryRun = true;
  if (!maxItems) maxItems = 999;
  Logger.log('=== syncAllItemMetadata ' + (dryRun ? '(DRY RUN)' : '(LIVE)') + (maxItems < 999 ? ' [first ' + maxItems + ']' : '') + ' START ===');

  // 1. Katana: UOM + category per SKU (raw UOM, no normalization)
  var katanaMap = buildKatanaUomMap_(); // { sku: { variantId, uom, category } }
  Logger.log('Katana map: ' + Object.keys(katanaMap).length + ' SKUs');

  // 2. Katana: average_cost per variantId from /inventory
  var costMap = {};
  var invPage = 1;
  while (invPage <= 50) {
    var invResult = katanaApiCall('inventory?limit=250&page=' + invPage);
    if (!invResult) break;
    var invRows = Array.isArray(invResult) ? invResult : (invResult.data || []);
    if (!invRows || invRows.length === 0) break;
    for (var ii = 0; ii < invRows.length; ii++) {
      var ir = invRows[ii];
      if (!ir.variant_id) continue;
      var c = parseFloat(ir.average_cost || 0);
      if (!costMap[ir.variant_id] || c > costMap[ir.variant_id]) costMap[ir.variant_id] = c;
    }
    if (invRows.length < 250) break;
    invPage++;
    Utilities.sleep(200);
  }
  Logger.log('Cost map: ' + Object.keys(costMap).length + ' variants');

  // 3. WASP: current StockingUnit + CategoryDescription per SKU
  var waspMap = buildWaspUomMap_(); // { sku: { uom, category } }
  Logger.log('WASP map: ' + Object.keys(waspMap).length + ' items');

  // 4. Build update list
  var updates = [];
  var fallbackCount = 0;
  var skippedNotInWasp = 0;
  var skippedNoChange = 0;

  var skus = Object.keys(katanaMap);
  for (var i = 0; i < skus.length; i++) {
    var sku = skus[i];
    var kInfo = katanaMap[sku]; // { variantId, uom, category }
    var wInfo = waspMap[sku];   // { uom, category } or undefined

    if (!wInfo) { skippedNotInWasp++; continue; }

    // Cost: live average_cost, fall back to last non-zero movement if 0
    var cost = costMap[kInfo.variantId] || 0;
    if (cost <= 0) {
      Utilities.sleep(300);
      cost = fetchLastAverageCostFromMovements(kInfo.variantId);
      if (cost > 0) fallbackCount++;
    }
    var factor = getUomConversionFactor(sku);
    var waspCost = cost > 0 ? (factor !== 1 ? cost / factor : cost) : 0;

    // UOM mismatch — compare mapped Katana name against WASP name (case-insensitive)
    var katanaUom = kInfo.uom;
    var uomMismatch = false;
    if (katanaUom && !UOM_CONVERSIONS[sku]) {
      var mappedKatanaUom_ = mapKatanaUomToWasp_(katanaUom);
      uomMismatch = mappedKatanaUom_.toLowerCase() !== (wInfo.uom || '').toLowerCase();
    }

    // Category mismatch (skip items with no category in Katana)
    var katanaCat = (kInfo.category || '').trim();
    var catMismatch = katanaCat &&
      katanaCat.toLowerCase() !== (wInfo.category || '').trim().toLowerCase();

    // Skip if nothing to update
    if (waspCost <= 0 && !uomMismatch && !catMismatch) { skippedNoChange++; continue; }

    // StockingUnit must always be present — WASP returns -57072 if omitted.
    // Map Katana abbreviation (EA, PC, g) to WASP unit name (Each, grams, etc.)
    var uomToSend = mapKatanaUomToWasp_(uomMismatch ? katanaUom : (wInfo.uom || 'Each'));

    var payload = {
      ItemNumber:    sku,
      StockingUnit:  uomToSend,
      PurchaseUnit:  uomToSend,
      SalesUnit:     uomToSend,
      DimensionInfo: { DimensionUnit: 'Inch', WeightUnit: 'Pound', VolumeUnit: 'Cubic Inch' }
    };
    var logParts = [];

    if (waspCost > 0) {
      payload.Cost = waspCost;
      logParts.push('Cost=' + waspCost.toFixed(4));
    }
    if (uomMismatch) {
      logParts.push('UOM: "' + wInfo.uom + '" -> "' + mapKatanaUomToWasp_(katanaUom) + '"');
    }
    if (catMismatch) {
      payload.CategoryDescription = katanaCat;
      logParts.push('Cat: "' + wInfo.category + '" → "' + katanaCat + '"');
    }

    updates.push({ payload: payload, log: sku + ': ' + logParts.join(' | ') });
  }

  // Apply maxItems cap
  if (updates.length > maxItems) updates = updates.slice(0, maxItems);

  Logger.log('Updates to apply: ' + updates.length +
    ' (' + fallbackCount + ' cost fallbacks, ' +
    skippedNotInWasp + ' not in WASP, ' +
    skippedNoChange + ' no change)');

  if (dryRun) {
    for (var d = 0; d < updates.length; d++) Logger.log('[DRY] ' + updates[d].log);
    Logger.log('Dry run complete — call syncAllItemMetadata(false) to apply.');
    return { total: updates.length, fallback: fallbackCount };
  }

  // 5. Batch push to WASP
  var url = CONFIG.WASP_BASE_URL + '/public-api/ic/item/updateInventoryItems';
  var BATCH_SIZE = 50;
  var DELAY_MS   = 300;
  var success = 0;
  var failed  = 0;
  var notInWasp = 0;
  var errors  = [];

  for (var b = 0; b < updates.length; b += BATCH_SIZE) {
    var batch = [];
    for (var bi = b; bi < Math.min(b + BATCH_SIZE, updates.length); bi++) {
      batch.push(updates[bi].payload);
    }
    var r = waspApiCall(url, batch);
    if (r.success) {
      success += batch.length;
    } else {
      // Retry individually to isolate bad items
      for (var ri = 0; ri < batch.length; ri++) {
        var sr = waspApiCall(url, [batch[ri]]);
        if (sr.success) {
          success++;
        } else {
          var errStr = String(sr.response || '');
          if (errStr.indexOf('-64001') > -1 || errStr.toLowerCase().indexOf('does not exist') > -1) {
            // Category mismatch (-64001) — retry without CategoryDescription so cost+UOM still apply
            if (batch[ri].CategoryDescription) {
              var fallbackPayload = {};
              for (var k in batch[ri]) {
                if (batch[ri].hasOwnProperty(k) && k !== 'CategoryDescription') {
                  fallbackPayload[k] = batch[ri][k];
                }
              }
              var fr = waspApiCall(url, [fallbackPayload]);
              if (fr.success) {
                success++;
                notInWasp++; // count category as unset but item was updated
              } else {
                notInWasp++;
              }
            } else {
              notInWasp++;
            }
          } else {
            failed++;
            errors.push(batch[ri].ItemNumber + ': ' + errStr.substring(0, 150));
          }
        }
        Utilities.sleep(200);
      }
    }
    if (b + BATCH_SIZE < updates.length) Utilities.sleep(DELAY_MS);
  }

  Logger.log('=== DONE === Updated: ' + success +
    ' | Not in WASP: ' + notInWasp +
    ' | Failed: ' + failed +
    ' | Skipped (not in WASP): ' + skippedNotInWasp +
    ' | Cost fallbacks: ' + fallbackCount);
  if (errors.length > 0) Logger.log('Errors:\n' + errors.join('\n'));

  return { success: success, failed: failed, notInWasp: notInWasp,
           skipped: skippedNotInWasp, fallback: fallbackCount, errors: errors };
}

/**
 * Daily trigger wrapper — called by time-based trigger.
 */
function syncCostsDailyTrigger() {
  var result = syncAllItemMetadata(false);
  Logger.log('Daily metadata sync: updated=' + result.success + ' skipped=' + result.skipped + ' failed=' + result.failed + ' fallback=' + result.fallback);
}

/**
 * Set up (or replace) a daily 3am trigger for syncCostsDailyTrigger.
 * Run once from the GAS editor to activate automatic daily sync.
 */
function setupCostSyncTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var t = 0; t < triggers.length; t++) {
    if (triggers[t].getHandlerFunction() === 'syncCostsDailyTrigger') {
      ScriptApp.deleteTrigger(triggers[t]);
    }
  }
  ScriptApp.newTrigger('syncCostsDailyTrigger')
    .timeBased()
    .everyDays(1)
    .atHour(13)
    .create();
  Logger.log('Daily cost sync trigger set for 1pm');
}

// ============================================
// 5-ITEM TEST SYNC — verify before bulk run
// ============================================

/**
 * Test sync on exactly 5 items before running on all ~300.
 * Verifies Cost, UOM, and Category updates work end-to-end.
 *
 * Usage (GAS editor):
 *   syncTestItems()        <- dry run: shows what would change, no writes
 *   syncTestItems(false)   <- LIVE: applies changes + read-back verification
 *
 * Edit TEST_SKUS below to target specific items.
 */
/** Live version of syncTestItems — applies changes to 5 test SKUs in WASP. */
function syncTestItemsLive() { syncTestItems(false); }

/** Run Cost + UOM + Category sync on first 10 items only (LIVE — for testing). */
function syncFirst10Live() { syncAllItemMetadata(false, 10); }

/** Run full Cost + UOM + Category sync on all ~300 items (LIVE). */
function syncAllItemMetadataLive() { syncAllItemMetadata(false); }

function syncTestItems(dryRun) {
  if (dryRun === undefined || dryRun === null) dryRun = true;

  var TEST_SKUS = [
    'CAR-4OZ',
    'UFC-4OZ',
    'TTFC-1OZ',
    'WH-8OZ-W',
    'B-PURPLE-GREEN-DUO'
  ];

  Logger.log('=== syncTestItems ' + (dryRun ? '(DRY RUN)' : '(LIVE)') + ' ===');
  Logger.log('Test SKUs: ' + TEST_SKUS.join(', '));
  Logger.log('');

  // 1. Katana: full UOM+category map (needed for variant IDs)
  Logger.log('Step 1: Building Katana map...');
  var katanaMap = buildKatanaUomMap_();

  // 2. Cost: scan Katana inventory for these variant IDs only
  Logger.log('Step 2: Fetching costs from Katana inventory...');
  var testVariantIds = {};
  for (var ts = 0; ts < TEST_SKUS.length; ts++) {
    var kEntry = katanaMap[TEST_SKUS[ts]];
    if (kEntry) testVariantIds[kEntry.variantId] = true;
  }
  var costMap = {};
  var invPage = 1;
  while (invPage <= 50) {
    var invResult = katanaApiCall('inventory?limit=250&page=' + invPage);
    if (!invResult) break;
    var invRows = Array.isArray(invResult) ? invResult : (invResult.data || []);
    if (!invRows || invRows.length === 0) break;
    for (var ii = 0; ii < invRows.length; ii++) {
      var ir = invRows[ii];
      if (!ir.variant_id || !testVariantIds[ir.variant_id]) continue;
      var c = parseFloat(ir.average_cost || 0);
      if (!costMap[ir.variant_id] || c > costMap[ir.variant_id]) costMap[ir.variant_id] = c;
    }
    if (invRows.length < 250) break;
    invPage++;
    Utilities.sleep(200);
  }

  // 3. WASP: full item map (StockingUnit + CategoryDescription)
  Logger.log('Step 3: Building WASP map...');
  var waspMap = buildWaspUomMap_();

  // 4. Compare and build update list
  Logger.log('Step 4: Comparing...');
  Logger.log('');
  var updates = [];

  for (var i = 0; i < TEST_SKUS.length; i++) {
    var sku = TEST_SKUS[i];
    var kData = katanaMap[sku];
    var wData = waspMap[sku];

    if (!kData) { Logger.log('[SKIP] ' + sku + ' — not found in Katana'); continue; }
    if (!wData) { Logger.log('[SKIP] ' + sku + ' — not found in WASP');   continue; }

    // Cost
    var cost = costMap[kData.variantId] || 0;
    if (cost <= 0) {
      Utilities.sleep(300);
      cost = fetchLastAverageCostFromMovements(kData.variantId);
    }
    var factor = getUomConversionFactor(sku);
    var waspCost = cost > 0 ? (factor !== 1 ? cost / factor : cost) : 0;

    // UOM mismatch — compare mapped Katana name against WASP name (case-insensitive)
    var katanaUom = kData.uom;
    var uomMismatch = false;
    if (katanaUom && !UOM_CONVERSIONS[sku]) {
      var mappedKUom_ = mapKatanaUomToWasp_(katanaUom);
      uomMismatch = mappedKUom_.toLowerCase() !== (wData.uom || '').toLowerCase();
    }

    // Category mismatch (skip items with no category in Katana)
    var katanaCat = (kData.category || '').trim();
    var catMismatch = katanaCat &&
      katanaCat.toLowerCase() !== (wData.category || '').trim().toLowerCase();

    Logger.log(sku + ':');
    Logger.log('  Katana  cost=' + (waspCost > 0 ? waspCost.toFixed(4) : 'none') +
               '  uom=' + (katanaUom || 'n/a') + '  cat=' + (katanaCat || 'n/a'));
    Logger.log('  WASP    uom=' + (wData.uom || 'n/a') + '  cat=' + (wData.category || 'n/a'));

    var logParts = [];
    var uomToSend = mapKatanaUomToWasp_(uomMismatch ? katanaUom : (wData.uom || 'Each'));
    var payload = {
      ItemNumber:    sku,
      StockingUnit:  uomToSend,
      PurchaseUnit:  uomToSend,
      SalesUnit:     uomToSend,
      DimensionInfo: { DimensionUnit: 'Inch', WeightUnit: 'Pound', VolumeUnit: 'Cubic Inch' }
    };

    if (waspCost > 0) { payload.Cost = waspCost; logParts.push('Cost=' + waspCost.toFixed(4)); }
    if (uomMismatch)  { logParts.push('UOM: "' + wData.uom + '" -> "' + mapKatanaUomToWasp_(katanaUom) + '"'); }
    if (catMismatch)  { payload.CategoryDescription = katanaCat; logParts.push('Cat: "' + wData.category + '" -> "' + katanaCat + '"'); }

    if (logParts.length === 0) {
      Logger.log('  -> No changes needed');
      Logger.log('');
      continue;
    }

    Logger.log('  -> ' + (dryRun ? '[DRY RUN] ' : '') + logParts.join(' | '));
    Logger.log('');
    updates.push({ sku: sku, payload: payload, log: logParts.join(' | ') });
  }

  if (updates.length === 0) {
    Logger.log('Nothing to update for test SKUs.');
    return;
  }

  if (dryRun) {
    Logger.log(updates.length + ' item(s) would be updated. Call syncTestItems(false) to apply.');
    return;
  }

  // 5. Apply one-by-one (not batched — clearer per-item logging for test)
  Logger.log('--- Applying changes ---');
  var url = CONFIG.WASP_BASE_URL + '/public-api/ic/item/updateInventoryItems';
  var okCount = 0;
  var failCount = 0;

  for (var u = 0; u < updates.length; u++) {
    var upd = updates[u];
    var r = waspApiCall(url, [upd.payload]);

    if (r.success) {
      Logger.log('OK: ' + upd.sku + ' — ' + upd.log);
      okCount++;
    } else {
      var errStr = String(r.response || '');
      // -64001: CategoryDescription not in WASP — retry without it so cost+UOM still apply
      if (upd.payload.CategoryDescription &&
          (errStr.indexOf('-64001') > -1 || errStr.toLowerCase().indexOf('does not exist') > -1)) {
        var fb = {};
        for (var fk in upd.payload) {
          if (upd.payload.hasOwnProperty(fk) && fk !== 'CategoryDescription') fb[fk] = upd.payload[fk];
        }
        var fr = waspApiCall(url, [fb]);
        if (fr.success) {
          Logger.log('OK (cat fallback -64001): ' + upd.sku + ' — cost+UOM updated, category skipped');
          okCount++;
        } else {
          Logger.log('FAIL: ' + upd.sku + ' — ' + errStr.substring(0, 200));
          failCount++;
        }
      } else {
        Logger.log('FAIL: ' + upd.sku + ' — ' + errStr.substring(0, 200));
        failCount++;
      }
    }
    Utilities.sleep(300);
  }

  // 6. Read-back: verify each item was actually updated in WASP
  Logger.log('');
  Logger.log('--- Read-back verification ---');
  var infosearchUrl = CONFIG.WASP_BASE_URL + '/public-api/ic/item/infosearch';
  for (var v = 0; v < updates.length; v++) {
    var vSku = updates[v].sku;
    Utilities.sleep(200);
    var rb = waspApiCall(infosearchUrl, { SearchPattern: vSku, PageNumber: 1, PageSize: 5 });
    if (!rb.success) { Logger.log(vSku + ': read-back failed'); continue; }
    var rbData = JSON.parse(rb.response);
    var rbRows = rbData.Data || rbData.data || [];
    var rbItem = null;
    for (var ri = 0; ri < rbRows.length; ri++) {
      if ((rbRows[ri].ItemNumber || rbRows[ri].itemNumber) === vSku) { rbItem = rbRows[ri]; break; }
    }
    if (!rbItem) { Logger.log(vSku + ': not found in read-back'); continue; }
    Logger.log(vSku + ' (after): cost=' + rbItem.Cost +
               '  uom=' + (rbItem.StockingUnit || rbItem.UnitOfMeasure || '?') +
               '  cat=' + (rbItem.CategoryDescription || '?'));
  }

  Logger.log('');
  Logger.log('=== syncTestItems DONE === OK: ' + okCount + ' | Failed: ' + failCount);
}

/**
 * Debug: dump ALL raw fields for a SKU from both Katana and WASP.
 * Shows uom, purchase uom, supplier pricing, cost, category — everything.
 * Useful for spotting UOM conversions before running bulk sync.
 *
 * Usage (GAS editor): edit SKU_TO_CHECK below, then run debugSkuRaw()
 */
function debugSkuRaw() {
  var SKU_TO_CHECK = 'EGG-X'; // <-- change this to any SKU

  Logger.log('=== debugSkuRaw: ' + SKU_TO_CHECK + ' ===');
  Logger.log('');

  // ---- KATANA: find variant by SKU (same approach as buildKatanaUomMap_) ----
  Logger.log('--- Katana: variant lookup ---');
  var katanaVariantId = null;
  var katanaProductId = null;
  var katanaMaterialId = null;
  var vResult = katanaApiCall('variants?sku=' + encodeURIComponent(SKU_TO_CHECK) + '&limit=10');
  var vRows = Array.isArray(vResult) ? vResult : (vResult && vResult.data ? vResult.data : []);
  // If sku filter not supported, fall back to first page scan
  if (vRows.length === 0) {
    var vPage = 1;
    while (vPage <= 10 && !katanaVariantId) {
      var vPageResult = katanaApiCall('variants?limit=250&page=' + vPage);
      var vPageRows = Array.isArray(vPageResult) ? vPageResult : (vPageResult && vPageResult.data ? vPageResult.data : []);
      for (var vi2 = 0; vi2 < vPageRows.length; vi2++) {
        if ((vPageRows[vi2].sku || '') === SKU_TO_CHECK) { vRows = [vPageRows[vi2]]; break; }
      }
      if (vPageRows.length < 250 || katanaVariantId) break;
      vPage++;
      Utilities.sleep(200);
    }
  }
  for (var vi = 0; vi < vRows.length; vi++) {
    var v = vRows[vi];
    if ((v.sku || '') !== SKU_TO_CHECK) continue;
    katanaVariantId = v.id;
    katanaProductId = v.product_id || null;
    katanaMaterialId = v.material_id || null;
    Logger.log('Variant found: id=' + v.id + '  sku=' + v.sku +
               '  product_id=' + v.product_id + '  material_id=' + v.material_id);
    Logger.log('  variant raw keys = ' + Object.keys(v).join(', '));
    break;
  }
  if (!katanaVariantId) Logger.log('Variant NOT found in Katana for SKU: ' + SKU_TO_CHECK);

  // ---- KATANA: fetch parent product or material ----
  Logger.log('');
  if (katanaProductId) {
    Logger.log('--- Katana: product (id=' + katanaProductId + ') ---');
    var prodResult = katanaApiCall('products/' + katanaProductId);
    var prod = Array.isArray(prodResult) ? prodResult[0] : prodResult;
    if (prod) {
      Logger.log('  name            = ' + prod.name);
      Logger.log('  uom             = ' + prod.uom);
      Logger.log('  category_name   = ' + prod.category_name);
      Logger.log('  raw keys        = ' + Object.keys(prod).join(', '));
    } else { Logger.log('  Not found.'); }
  }
  if (katanaMaterialId) {
    Logger.log('--- Katana: material (id=' + katanaMaterialId + ') ---');
    var matResult = katanaApiCall('materials/' + katanaMaterialId);
    var mat = Array.isArray(matResult) ? matResult[0] : matResult;
    if (mat) {
      Logger.log('  name            = ' + mat.name);
      Logger.log('  uom             = ' + mat.uom);
      Logger.log('  purchase_uom    = ' + mat.purchase_uom);
      Logger.log('  category_name   = ' + mat.category_name);
      Logger.log('  raw keys        = ' + Object.keys(mat).join(', '));
    } else { Logger.log('  Not found.'); }
  }

  // ---- KATANA: inventory (cost) ----
  Logger.log('');
  Logger.log('--- Katana: inventory (cost) ---');
  if (katanaVariantId) {
    var invResult = katanaApiCall('inventory?variant_id=' + katanaVariantId + '&limit=50');
    var invRows = Array.isArray(invResult) ? invResult : (invResult && invResult.data ? invResult.data : []);
    if (invRows.length === 0) Logger.log('No inventory rows found for variant_id=' + katanaVariantId);
    for (var ii = 0; ii < invRows.length; ii++) {
      var inv = invRows[ii];
      Logger.log('  location_id=' + inv.location_id +
                 '  qty=' + inv.in_stock +
                 '  avg_cost=' + inv.average_cost +
                 '  raw keys=' + Object.keys(inv).join(', '));
    }
  } else {
    Logger.log('No variant_id found — skipping inventory lookup.');
  }

  // ---- KATANA: supplier pricing (purchase orders / supplier info) ----
  Logger.log('');
  Logger.log('--- Katana: purchase orders (latest 3 for this SKU) ---');
  var poResult = katanaApiCall('purchase_orders?search=' + encodeURIComponent(SKU_TO_CHECK) + '&limit=3');
  var poRows = Array.isArray(poResult) ? poResult : (poResult && poResult.data ? poResult.data : []);
  if (poRows.length === 0) {
    Logger.log('No POs found.');
  } else {
    for (var poi = 0; poi < poRows.length; poi++) {
      var po = poRows[poi];
      var lines = po.line_items || po.purchase_order_rows || [];
      for (var li = 0; li < lines.length; li++) {
        var line = lines[li];
        var lineSku = line.sku || line.variant_code || (line.variant && line.variant.sku) || '';
        if (lineSku !== SKU_TO_CHECK) continue;
        Logger.log('  PO #' + (po.order_no || po.id) +
                   '  purchase_uom=' + (line.purchase_uom || line.uom || '?') +
                   '  price=' + (line.price || line.unit_price || '?') +
                   '  qty=' + (line.quantity || '?') +
                   '  line raw keys=' + Object.keys(line).join(', '));
      }
    }
  }

  // ---- WASP: infosearch ----
  Logger.log('');
  Logger.log('--- WASP: infosearch ---');
  var waspUrl = CONFIG.WASP_BASE_URL + '/public-api/ic/item/infosearch';
  var wResult = waspApiCall(waspUrl, { SearchPattern: SKU_TO_CHECK, PageNumber: 1, PageSize: 5 });
  if (!wResult.success) {
    Logger.log('WASP call failed: ' + wResult.response);
  } else {
    var wData = JSON.parse(wResult.response);
    var wRows = wData.Data || wData.data || [];
    var wItem = null;
    for (var wi = 0; wi < wRows.length; wi++) {
      if ((wRows[wi].ItemNumber || wRows[wi].itemNumber || '') === SKU_TO_CHECK) {
        wItem = wRows[wi]; break;
      }
    }
    if (!wItem) {
      Logger.log('Not found in WASP.');
    } else {
      Logger.log('  ItemNumber        = ' + wItem.ItemNumber);
      Logger.log('  Cost              = ' + wItem.Cost);
      Logger.log('  StockingUnit      = ' + wItem.StockingUnit);
      Logger.log('  PurchaseUnit      = ' + wItem.PurchaseUnit);
      Logger.log('  SalesUnit         = ' + wItem.SalesUnit);
      Logger.log('  CategoryDescription = ' + wItem.CategoryDescription);
      Logger.log('  TrackbyInfo       = ' + JSON.stringify(wItem.TrackbyInfo));
      Logger.log('  All WASP keys     = ' + Object.keys(wItem).join(', '));
    }
  }

  Logger.log('');
  Logger.log('=== debugSkuRaw DONE ===');
}

/**
 * Debug: send a single item update to WASP and log the full raw response.
 * Run this to see exactly what WASP returns (error code + message).
 */
function debugWaspSingleUpdate() {
  var baseUrl = CONFIG.WASP_BASE_URL;
  var options = {
    method: 'POST',
    headers: { 'Authorization': 'Bearer ' + CONFIG.WASP_TOKEN, 'Content-Type': 'application/json' },
    muteHttpExceptions: true
  };

  // Step 1: read current cost via infosearch (exact ItemNumber lookup)
  Logger.log('=== STEP 1: read current cost via infosearch ===');
  options.payload = JSON.stringify({ SearchPattern: 'CAR-4OZ', PageNumber: 1, PageSize: 5 });
  var r1 = UrlFetchApp.fetch(baseUrl + '/public-api/ic/item/infosearch', options);
  var s1 = JSON.parse(r1.getContentText());
  var car4oz = (s1.Data || [])[0] || null;
  Logger.log('Current cost: ' + (car4oz ? car4oz.Cost : 'NOT FOUND'));

  // Step 2: update with StockingUnit='Each' + partial nested DimensionInfo (units only, no dimension values)
  Logger.log('=== STEP 2: update StockingUnit=Each + partial DimensionInfo ===');
  options.payload = JSON.stringify([{
    ItemNumber:   'CAR-4OZ',
    Cost:         1.1111,
    StockingUnit: 'Each',
    PurchaseUnit: 'Each',
    SalesUnit:    'Each',
    DimensionInfo: { DimensionUnit: 'Inch', WeightUnit: 'Pound', VolumeUnit: 'Cubic Inch' }
  }]);
  var r2 = UrlFetchApp.fetch(baseUrl + '/public-api/ic/item/updateInventoryItems', options);
  Logger.log('Response: ' + r2.getContentText());

  // Step 3: read back to see if cost changed
  Logger.log('=== STEP 3: read back cost ===');
  options.payload = JSON.stringify({ SearchPattern: 'CAR-4OZ', PageNumber: 1, PageSize: 5 });
  var r3 = UrlFetchApp.fetch(baseUrl + '/public-api/ic/item/infosearch', options);
  var s3 = JSON.parse(r3.getContentText());
  var car4oz_after = (s3.Data || [])[0] || null;
  Logger.log('Cost after:  ' + (car4oz_after ? car4oz_after.Cost : 'NOT FOUND'));
  Logger.log('Expected 1.1111: ' + (car4oz_after && Math.abs(car4oz_after.Cost - 1.1111) < 0.0001 ? 'YES - SUCCESS!' : 'NO'));
}

/**
 * DEBUG: Probe WASP for category-related endpoints and field names.
 * Tries common endpoint patterns and update payload variations to find
 * how to set an item's category via the API.
 * Run from GAS editor — check execution log for results.
 */
function debugWaspCategory() {
  var baseUrl = CONFIG.WASP_BASE_URL;
  var options = {
    method: 'POST',
    headers: { 'Authorization': 'Bearer ' + CONFIG.WASP_TOKEN, 'Content-Type': 'application/json' },
    muteHttpExceptions: true
  };

  // Step 1: read CAR-4OZ via infosearch — note current CategoryDescription and all fields
  Logger.log('=== STEP 1: current item state (infosearch) ===');
  options.payload = JSON.stringify({ SearchPattern: 'CAR-4OZ', PageNumber: 1, PageSize: 5 });
  var r1 = UrlFetchApp.fetch(baseUrl + '/public-api/ic/item/infosearch', options);
  var s1 = JSON.parse(r1.getContentText());
  var item = (s1.Data || [])[0] || null;
  Logger.log('CategoryDescription: ' + (item ? item.CategoryDescription : 'NOT FOUND'));
  Logger.log('All fields: ' + Object.keys(item || {}).join(', '));

  // Step 2: probe GET-style endpoints for category list
  var categoryEndpoints = [
    '/public-api/ic/category/list',
    '/public-api/ic/category/search',
    '/public-api/ic/item/category',
    '/public-api/ic/itemcategory/list',
    '/public-api/ic/itemcategory/search'
  ];
  Logger.log('=== STEP 2: probing category list endpoints ===');
  for (var i = 0; i < categoryEndpoints.length; i++) {
    options.payload = JSON.stringify({ PageNumber: 1, PageSize: 10 });
    var r = UrlFetchApp.fetch(baseUrl + categoryEndpoints[i], options);
    Logger.log(categoryEndpoints[i] + ' → HTTP ' + r.getResponseCode() + ' | ' + r.getContentText().substring(0, 200));
  }

  // Step 3: try updateInventoryItems with CategoryID (numeric) instead of CategoryDescription
  Logger.log('=== STEP 3: updateInventoryItems with CategoryID field ===');
  options.payload = JSON.stringify([{
    ItemNumber:    'CAR-4OZ',
    CategoryID:    1,
    StockingUnit:  'Each',
    PurchaseUnit:  'Each',
    SalesUnit:     'Each',
    DimensionInfo: { DimensionUnit: 'Inch', WeightUnit: 'Pound', VolumeUnit: 'Cubic Inch' }
  }]);
  var r3 = UrlFetchApp.fetch(baseUrl + '/public-api/ic/item/updateInventoryItems', options);
  Logger.log('CategoryID=1 response: ' + r3.getContentText().substring(0, 300));

  // Step 4: try updateInventoryItems with Category (plain string, no "Description" suffix)
  Logger.log('=== STEP 4: updateInventoryItems with Category field ===');
  options.payload = JSON.stringify([{
    ItemNumber:    'CAR-4OZ',
    Category:      'FINISHED GOODS',
    StockingUnit:  'Each',
    PurchaseUnit:  'Each',
    SalesUnit:     'Each',
    DimensionInfo: { DimensionUnit: 'Inch', WeightUnit: 'Pound', VolumeUnit: 'Cubic Inch' }
  }]);
  var r4 = UrlFetchApp.fetch(baseUrl + '/public-api/ic/item/updateInventoryItems', options);
  Logger.log('Category field response: ' + r4.getContentText().substring(0, 300));

  // Step 5: read back via infosearch — did CategoryDescription change?
  Logger.log('=== STEP 5: read back CategoryDescription ===');
  options.payload = JSON.stringify({ SearchPattern: 'CAR-4OZ', PageNumber: 1, PageSize: 5 });
  var r5 = UrlFetchApp.fetch(baseUrl + '/public-api/ic/item/infosearch', options);
  var s5 = JSON.parse(r5.getContentText());
  var item5 = (s5.Data || [])[0] || null;
  Logger.log('CategoryDescription after steps 3+4: "' + (item5 ? item5.CategoryDescription : 'NOT FOUND') + '"');

  // Step 6: try CategoryDescription directly (the documented field name)
  Logger.log('=== STEP 6: updateInventoryItems with CategoryDescription field ===');
  options.payload = JSON.stringify([{
    ItemNumber:          'CAR-4OZ',
    CategoryDescription: 'FINISHED GOODS',
    StockingUnit:        'Each',
    PurchaseUnit:        'Each',
    SalesUnit:           'Each',
    DimensionInfo: { DimensionUnit: 'Inch', WeightUnit: 'Pound', VolumeUnit: 'Cubic Inch' }
  }]);
  var r6 = UrlFetchApp.fetch(baseUrl + '/public-api/ic/item/updateInventoryItems', options);
  Logger.log('CategoryDescription response: ' + r6.getContentText().substring(0, 400));

  // Step 7: probe category create endpoints
  Logger.log('=== STEP 7: probe category create/manage endpoints ===');
  var createEndpoints = [
    '/public-api/ic/category/create',
    '/public-api/ic/category/save',
    '/public-api/ic/itemcategory/create',
    '/public-api/ic/itemcategory/save',
    '/public-api/ic/item/createcategory'
  ];
  for (var j = 0; j < createEndpoints.length; j++) {
    options.payload = JSON.stringify({ Name: 'TEST' });
    var rc = UrlFetchApp.fetch(baseUrl + createEndpoints[j], options);
    Logger.log(createEndpoints[j] + ' → HTTP ' + rc.getResponseCode() + ' | ' + rc.getContentText().substring(0, 150));
  }

  // Step 8: read back after step 6 to see if CategoryDescription was set
  Logger.log('=== STEP 8: read back after CategoryDescription update ===');
  options.payload = JSON.stringify({ SearchPattern: 'CAR-4OZ', PageNumber: 1, PageSize: 5 });
  var r8 = UrlFetchApp.fetch(baseUrl + '/public-api/ic/item/infosearch', options);
  var s8 = JSON.parse(r8.getContentText());
  var item8 = (s8.Data || [])[0] || null;
  Logger.log('CategoryDescription final: "' + (item8 ? item8.CategoryDescription : 'NOT FOUND') + '"');
}

/**
 * Test helper: sync cost for ONE specific SKU using the new average_cost source.
 * Change SKU_TO_TEST, run from the editor, then check the item in WASP UI.
 */
function testSyncOneCost() {
  var SKU_TO_TEST = 'WH-8OZ-W'; // ← change to any real SKU

  // Get variant ID
  var variant = getKatanaVariantBySku(SKU_TO_TEST);
  if (!variant) { Logger.log('Variant not found: ' + SKU_TO_TEST); return; }

  // Try current average cost from inventory
  var invResult = katanaApiCall('inventory?variant_id=' + variant.id + '&limit=10');
  var cost = 0;
  if (invResult) {
    var rows = Array.isArray(invResult) ? invResult : (invResult.data || []);
    for (var i = 0; i < rows.length; i++) {
      var c = parseFloat(rows[i].average_cost || 0);
      if (c > cost) cost = c;
    }
  }
  Logger.log('SKU: ' + SKU_TO_TEST + '  average_cost from inventory: ' + cost);

  // Fallback if zero
  if (cost <= 0) {
    cost = fetchLastAverageCostFromMovements(variant.id);
    Logger.log('Fallback from movements: ' + cost);
  }

  if (cost <= 0) { Logger.log('No cost found — nothing to update'); return; }

  var r = waspUpdateItemCost(SKU_TO_TEST, cost);
  Logger.log('WASP response: code=' + r.code + '  success=' + r.success + '  body=' + String(r.response).substring(0, 300));
}

/**
 * DEBUG: Dump all variant fields + recipe rows for a SKU.
 * Run this to find which field holds the BOM total cost.
 * Check the Execution Log after running.
 */
function debugVariantCostFields() {
  var SKU_TO_TEST = 'UFC-4OZ'; // ← change to any SKU with a BOM

  // 1. All variant fields
  var variant = getKatanaVariantBySku(SKU_TO_TEST);
  if (!variant) { Logger.log('Variant not found: ' + SKU_TO_TEST); return; }
  Logger.log('=== VARIANT FIELDS ===');
  Logger.log(JSON.stringify(variant, null, 2));

  // 2. Recipe rows for this variant (BOM ingredients)
  var variantId = variant.id;
  if (!variantId) { Logger.log('No variant id'); return; }
  var recipes = katanaApiCall('recipes?product_variant_id=' + variantId + '&limit=50');
  Logger.log('=== RECIPE ROWS ===');
  Logger.log(JSON.stringify(recipes, null, 2));
}

/**
 * DEBUG: Fetch one item from WASP advancedinfosearch and log ALL fields.
 * Run this to find the dimension UoM field name and value used by the API.
 */
function debugWaspItemFields() {
  var SKU = 'CAR-4OZ'; // ← any valid SKU
  var url = CONFIG.WASP_BASE_URL + '/public-api/ic/item/advancedinfosearch';
  var result = waspApiCall(url, { SearchPattern: SKU, PageNumber: 1, PageSize: 5 });
  Logger.log(JSON.stringify(result, null, 2));
}

/**
 * Sync item costs to WASP from the "Cost Import" tab in the Command Center sheet.
 *
 * HOW TO USE:
 *   1. Open the Command Center spreadsheet.
 *   2. Create a tab named exactly: Cost Import
 *   3. Paste the InventoryItems-katana-cost-per-unit.xlsx data in starting at A1
 *      (Row 1 = headers, Col B = SKU, Col F = Average cost)
 *   4. Run syncCostsFromSheet() from the GAS editor (09_Utils.gs)
 *
 * Skips rows where SKU is blank or cost is 0.
 * Batches WASP calls with a short delay to avoid rate limits.
 */
function syncCostsFromSheet() {
  var ss = SpreadsheetApp.openById(CONFIG.DEBUG_SHEET_ID);
  var sheet = ss.getSheetByName('Cost Import');
  if (!sheet) {
    Logger.log('ERROR: "Cost Import" tab not found in Command Center. Create it and paste the xlsx data first.');
    return;
  }

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) { Logger.log('No data rows found in Cost Import tab.'); return; }

  var data = sheet.getRange(2, 1, lastRow - 1, 6).getValues(); // rows 2 onward, cols A-F

  var updated = 0;
  var skipped = 0;
  var errors = 0;

  Logger.log('=== syncCostsFromSheet START — ' + (lastRow - 1) + ' rows ===');

  for (var i = 0; i < data.length; i++) {
    var sku  = String(data[i][1] || '').trim(); // col B
    var uom  = String(data[i][4] || '').trim(); // col E
    var cost = parseFloat(data[i][5] || 0);     // col F

    if (!sku || cost <= 0) { skipped++; continue; }

    var result = waspUpdateItemCost(sku, cost, uom);

    if (result.success) {
      updated++;
      Logger.log('OK  ' + sku + '  uom=' + uom + '  → $' + cost);
    } else {
      errors++;
      Logger.log('FAIL  ' + sku + '  uom=' + uom + '  cost=' + cost + '  response=' + String(result.response).substring(0, 120));
    }

    // Small delay every 50 rows to avoid WASP rate limits
    if ((i + 1) % 50 === 0) Utilities.sleep(300);
  }

  Logger.log('=== DONE: updated=' + updated + '  skipped=' + skipped + '  errors=' + errors + ' ===');
}

// ============================================
// UOM FIX — Try updating UOM via API
// ============================================

/**
 * Test whether WASP API allows changing UOM on an item with transaction history.
 * The UI blocks this, but the API sometimes has fewer restrictions.
 *
 * Run from GAS editor (09_Utils.gs): testUpdateItemUom()
 * Check execution log for result.
 */
function testUpdateItemUom() {
  var SKU = 'B-WAX';
  var NEW_UOM = 'g'; // grams

  Logger.log('=== testUpdateItemUom ===');
  Logger.log('SKU: ' + SKU + '  New UOM: ' + NEW_UOM);

  var url = CONFIG.WASP_BASE_URL + '/public-api/ic/item/updateInventoryItems';
  var payload = [{
    ItemNumber:   SKU,
    StockingUnit: NEW_UOM,
    PurchaseUnit: NEW_UOM,
    SalesUnit:    NEW_UOM,
    DimensionInfo: {
      DimensionUnit: 'Inch',
      WeightUnit:    'Pound',
      VolumeUnit:    'Cubic Inch'
    }
  }];

  Logger.log('Sending: ' + JSON.stringify(payload));
  var r = waspApiCall(url, payload);
  Logger.log('HTTP: ' + r.code + '  success: ' + r.success);
  Logger.log('Response: ' + String(r.response).substring(0, 500));

  if (r.success) {
    Logger.log('SUCCESS — API accepted the UOM change. Check B-WAX in WASP UI to confirm.');
  } else {
    Logger.log('FAILED — response above shows why. Try different DimensionInfo values.');
  }
}

// ============================================
// UOM CONVERSION AUDIT
// ============================================

/**
 * Scan all Katana materials and purchasable products for items where
 * purchase_uom differs from stocking uom.
 *
 * These are the items that need an entry in UOM_CONVERSIONS (00_Config.js)
 * before running the bulk sync — otherwise the cost written to WASP will
 * be based on the wrong unit.
 *
 * Output: one line per mismatch → "SKU | stocking=X | purchase=Y"
 * Also prints a ready-to-paste UOM_CONVERSIONS block at the end.
 *
 * Run auditKatanaConversions() in 09_Utils.js from the GAS editor.
 */
/**
 * Log direct Katana supply-details URLs for a list of SKUs.
 * Run logMaterialLinks() in 09_Utils.js — copy the links from the log.
 */
/**
 * Check WASP stocking units for the 13 items with purchase_uom conversions.
 * Run checkWaspUnitsForConversions() in 09_Utils.js
 */
function checkWaspUnitsForConversions() {
  var SKUS = [
    'NI-ATPSWABS','NI-NAPKINS','NI-GLOVEMD','NI-GLOVESM','NI-GLOVELR','NI-GLOVECR',
    'NI-STEELWOOL','NI-PAPERTOWEL8','NI-SCRUB-10','NI-FIRE-BNKT',
    'INFBAG-160','FBA-BARCODE','FBA-FRAGILE',
    'H-ARNF','H-COML','EGG-X',
    'EO-THY','EO-LAV','EO-PEP','EO-CS','EO-TT','EO-CB','B-WAX','B-PROP'
  ];
  var url = CONFIG.WASP_BASE_URL + '/public-api/ic/item/infosearch';
  Logger.log('SKU                | WASP StockingUnit | Katana stocking | Katana purchase');
  Logger.log('-------------------+-------------------+-----------------+----------------');

  var KATANA_INFO = {
    'NI-ATPSWABS': { stock: 'EA',    purchase: 'BOX'    },
    'NI-NAPKINS':  { stock: 'pack',  purchase: 'box'    },
    'NI-GLOVEMD':  { stock: 'EA',    purchase: 'pack'   },
    'NI-GLOVESM':  { stock: 'EA',    purchase: 'pack'   },
    'NI-GLOVELR':  { stock: 'EA',    purchase: 'pack'   },
    'NI-GLOVECR':  { stock: 'EA',    purchase: 'pack'   },
    'NI-STEELWOOL':{ stock: 'PC',    purchase: 'pack'   },
    'NI-PAPERTOWEL8':{ stock:'ROLLS',purchase: 'pcs'    },
    'NI-SCRUB-10': { stock: 'pc',    purchase: 'pack'   },
    'NI-FIRE-BNKT':{ stock: 'pc',    purchase: 'pack'   },
    'INFBAG-160':  { stock: 'PC',    purchase: 'pack'   },
    'FBA-BARCODE': { stock: 'PC',    purchase: 'pack'   },
    'FBA-FRAGILE': { stock: 'PC',    purchase: 'pack'   },
    'H-ARNF':      { stock: 'g',     purchase: 'lbs'    },
    'H-COML':      { stock: 'g',     purchase: 'lbs'    },
    'EGG-X':       { stock: 'pcs',   purchase: 'dozen'  },
    'EO-THY':      { stock: 'g',     purchase: 'kg'     },
    'EO-LAV':      { stock: 'g',     purchase: 'kg'     },
    'EO-PEP':      { stock: 'g',     purchase: 'kg'     },
    'EO-CS':       { stock: 'g',     purchase: 'kg'     },
    'EO-TT':       { stock: 'g',     purchase: 'kg'     },
    'EO-CB':       { stock: 'g',     purchase: 'kg'     },
    'B-WAX':       { stock: 'g',     purchase: 'kg'     },
    'B-PROP':      { stock: 'g',     purchase: 'kg'     }
  };

  for (var i = 0; i < SKUS.length; i++) {
    var sku = SKUS[i];
    var r = waspApiCall(url, { SearchPattern: sku, PageNumber: 1, PageSize: 5 });
    var waspUnit = 'NOT IN WASP';
    if (r.success) {
      var data = JSON.parse(r.response);
      var rows = data.Data || data.data || [];
      for (var j = 0; j < rows.length; j++) {
        if ((rows[j].ItemNumber || '') === sku) { waspUnit = rows[j].StockingUnit || '?'; break; }
      }
    }
    var ki = KATANA_INFO[sku] || { stock: '?', purchase: '?' };
    var flag = (waspUnit !== 'NOT IN WASP' && waspUnit !== ki.stock) ? '  *** MISMATCH' : '';
    Logger.log(
      sku + Array(20 - sku.length).join(' ') + '| ' +
      waspUnit + Array(18 - waspUnit.length).join(' ') + '| ' +
      ki.stock + Array(16 - ki.stock.length).join(' ') + '| ' +
      ki.purchase + flag
    );
    Utilities.sleep(300);
  }
  Logger.log('');
  Logger.log('*** = WASP stocking unit differs from Katana stocking unit → needs UOM_CONVERSIONS entry');
}

function logMaterialLinks() {
  var TARGET_SKUS = [
    'NI-ATPSWABS','NI-NAPKINS','NI-GLOVEMD','NI-GLOVESM','NI-GLOVELR','NI-GLOVECR',
    'NI-STEELWOOL','NI-PAPERTOWEL8','NI-SCRUB-10','NI-FIRE-BNKT',
    'INFBAG-160','FBA-BARCODE','FBA-FRAGILE'
  ];

  Logger.log('=== Katana Supply Details Links ===');

  // Build SKU → material_id map via /variants
  var skuToMaterialId = {};
  var page = 1;
  while (page <= 20) {
    var result = katanaApiCall('variants?limit=250&page=' + page);
    var rows = result && result.data ? result.data : (result || []);
    if (!rows || rows.length === 0) break;
    for (var i = 0; i < rows.length; i++) {
      var v = rows[i];
      if (v.material_id && TARGET_SKUS.indexOf(v.sku) > -1) {
        skuToMaterialId[v.sku] = v.material_id;
      }
    }
    if (rows.length < 250) break;
    page++;
    Utilities.sleep(200);
  }

  for (var s = 0; s < TARGET_SKUS.length; s++) {
    var sku = TARGET_SKUS[s];
    var mid = skuToMaterialId[sku];
    if (mid) {
      Logger.log(sku + '  →  https://factory.katanamrp.com/material/' + mid + '/supplydetails');
    } else {
      Logger.log(sku + '  →  NOT FOUND (check if it is a product, not material)');
    }
  }
  Logger.log('=== DONE ===');
}

function auditKatanaConversions() {
  Logger.log('=== auditKatanaConversions ===');
  Logger.log('Scanning materials and products for purchase_uom ≠ stocking uom...');
  Logger.log('');

  var mismatches = []; // { sku, name, stockUom, purchaseUom, type }

  // ---- 1. Materials ----
  var mPage = 1;
  var matTotal = 0;
  while (mPage <= 20) {
    var mResult = katanaApiCall('materials?limit=250&page=' + mPage);
    var mats = mResult && mResult.data ? mResult.data : (mResult || []);
    if (!mats || mats.length === 0) break;
    matTotal += mats.length;
    for (var mi = 0; mi < mats.length; mi++) {
      var mat = mats[mi];
      var stockUom   = (mat.uom          || '').trim();
      var purchaseUom = (mat.purchase_uom || '').trim();
      if (!purchaseUom || purchaseUom === '' || purchaseUom === stockUom) continue;
      // Find SKUs for this material via its variants
      var mVariants = mat.variants || [];
      for (var mvi = 0; mvi < mVariants.length; mvi++) {
        var mv = mVariants[mvi];
        var sku = mv.sku || mv.variant_code || '';
        if (!sku || isSkippedSku(sku)) continue;
        mismatches.push({ sku: sku, name: mat.name || '', stockUom: stockUom, purchaseUom: purchaseUom, type: 'material' });
      }
    }
    if (mats.length < 250) break;
    mPage++;
    Utilities.sleep(200);
  }
  Logger.log('Materials scanned: ' + matTotal);

  // ---- 2. Products ----
  // Katana products can have a purchase_uom if the product is purchasable (Buy checkbox)
  var pPage = 1;
  var prodTotal = 0;
  while (pPage <= 20) {
    var pResult = katanaApiCall('products?limit=250&page=' + pPage);
    var prods = pResult && pResult.data ? pResult.data : (pResult || []);
    if (!prods || prods.length === 0) break;
    prodTotal += prods.length;
    for (var pi = 0; pi < prods.length; pi++) {
      var prod = prods[pi];
      var pStockUom    = (prod.uom          || '').trim();
      var pPurchaseUom  = (prod.purchase_uom || '').trim();
      if (!pPurchaseUom || pPurchaseUom === '' || pPurchaseUom === pStockUom) continue;
      var pVariants = prod.variants || [];
      for (var pvi = 0; pvi < pVariants.length; pvi++) {
        var pv = pVariants[pvi];
        var pSku = pv.sku || pv.variant_code || '';
        if (!pSku || isSkippedSku(pSku)) continue;
        mismatches.push({ sku: pSku, name: prod.name || '', stockUom: pStockUom, purchaseUom: pPurchaseUom, type: 'product' });
      }
    }
    if (prods.length < 250) break;
    pPage++;
    Utilities.sleep(200);
  }
  Logger.log('Products scanned:  ' + prodTotal);
  Logger.log('');

  // ---- Results ----
  if (mismatches.length === 0) {
    Logger.log('No UOM conversions found — purchase_uom matches stocking uom for all items.');
    Logger.log('UOM_CONVERSIONS in 00_Config.js can stay empty.');
    return;
  }

  Logger.log(mismatches.length + ' conversion(s) found:');
  Logger.log('');
  for (var i = 0; i < mismatches.length; i++) {
    var m = mismatches[i];
    Logger.log('[' + m.type + ']  ' + m.sku + '  |  name="' + m.name +
               '"  |  stocking=' + m.stockUom + '  |  purchase=' + m.purchaseUom);
  }

  // ---- Ready-to-paste UOM_CONVERSIONS block ----
  Logger.log('');
  Logger.log('--- Paste into UOM_CONVERSIONS in 00_Config.js (add correct factor for each) ---');
  Logger.log('var UOM_CONVERSIONS = {');
  for (var j = 0; j < mismatches.length; j++) {
    var mm = mismatches[j];
    Logger.log('  \'' + mm.sku + '\': { katanaUom: \'' + mm.stockUom +
               '\', purchaseUom: \'' + mm.purchaseUom + '\', factor: ??? },  // ' + mm.name);
  }
  Logger.log('};');
  Logger.log('');
  Logger.log('NOTE: Replace ??? with the actual conversion factor (e.g. if 1 BOX = 12 pcs, factor=12).');
  Logger.log('=== auditKatanaConversions DONE ===');
}

// ============================================
// UOM AUDIT & BULK FIX: Katana → WASP
// ============================================

/**
 * Debug: log ALL fields returned by the Katana products and materials API.
 * Run this once to confirm the correct field names for UOM and category.
 * Check the execution log for output.
 */
function debugKatanaItemFields() {
  Logger.log('=== PRODUCT FIELDS ===');
  var pResult = katanaApiCall('products?limit=1');
  var products = pResult && pResult.data ? pResult.data : (pResult || []);
  if (products.length > 0) {
    Logger.log('Fields: ' + Object.keys(products[0]).join(', '));
    Logger.log('Sample: ' + JSON.stringify(products[0]));
  } else {
    Logger.log('No products returned');
  }

  Logger.log('=== MATERIAL FIELDS ===');
  var mResult = katanaApiCall('materials?limit=1');
  var materials = mResult && mResult.data ? mResult.data : (mResult || []);
  if (materials.length > 0) {
    Logger.log('Fields: ' + Object.keys(materials[0]).join(', '));
    Logger.log('Sample: ' + JSON.stringify(materials[0]));
  } else {
    Logger.log('No materials returned');
  }
}

/**
 * Internal: build a map of SKU → Katana UOM (normalized).
 * Fetches products + materials in bulk first, then joins to variants.
 * Avoids N individual API calls.
 * @returns {Object} { sku: { variantId, uom } }
 */
function buildKatanaUomMap_() {
  // 1. Products: productId → { uom, category }
  var productInfo = {};
  var pPage = 1;
  while (pPage <= 20) {
    var pResult = katanaApiCall('products?limit=250&page=' + pPage);
    var products = pResult && pResult.data ? pResult.data : (pResult || []);
    if (!products || products.length === 0) break;
    for (var pi = 0; pi < products.length; pi++) {
      var p = products[pi];
      // Katana products use unit_of_measure; materials use uom
      if (p.id) productInfo[p.id] = { uom: p.uom || '', category: p.category_name || '' };
    }
    if (products.length < 250) break;
    pPage++;
    Utilities.sleep(200);
  }
  Logger.log('Products map: ' + Object.keys(productInfo).length);

  // 2. Materials: materialId → { uom, category }
  var materialInfo = {};
  var mPage = 1;
  while (mPage <= 20) {
    var mResult = katanaApiCall('materials?limit=250&page=' + mPage);
    var materials = mResult && mResult.data ? mResult.data : (mResult || []);
    if (!materials || materials.length === 0) break;
    for (var mi = 0; mi < materials.length; mi++) {
      var m = materials[mi];
      if (m.id) materialInfo[m.id] = { uom: m.uom || '', category: m.category_name || '' };
    }
    if (materials.length < 250) break;
    mPage++;
    Utilities.sleep(200);
  }
  Logger.log('Materials map: ' + Object.keys(materialInfo).length);

  // 3. Variants: join to product/material info
  var variantMap = {};
  var vPage = 1;
  while (vPage <= 20) {
    var vResult = katanaApiCall('variants?limit=250&page=' + vPage);
    var variants = vResult && vResult.data ? vResult.data : (vResult || []);
    if (!variants || variants.length === 0) break;
    for (var vi = 0; vi < variants.length; vi++) {
      var v = variants[vi];
      if (!v.sku || isSkippedSku(v.sku)) continue;
      var info = productInfo[v.product_id] || materialInfo[v.material_id] || { uom: '', category: '' };
      // Store raw Katana UOM (no normalization) — written as-is to WASP; normalizeUomForCompare_ used only for comparison.
      variantMap[v.sku] = { variantId: v.id, uom: (info.uom || '').trim(), category: info.category };
    }
    if (variants.length < 250) break;
    vPage++;
    Utilities.sleep(200);
  }
  return variantMap;
}

/**
 * Internal: build a map of ItemNumber → StockingUnit from WASP catalog.
 * Also logs raw field names from the first item to help debugging.
 * @returns {Object} { itemNumber: stockingUnit }
 */
/**
 * Builds a map of WASP ItemNumber → { uom, category } using the infosearch endpoint.
 * infosearch returns StockingUnit, PurchaseUnit, SalesUnit, CategoryDescription.
 * (advancedinfosearch only returns SalesUnit — not suitable for UOM audit.)
 * @returns {Object} { itemNumber: { uom: string, category: string } }
 */
function buildWaspUomMap_() {
  var waspMap = {};
  var url = CONFIG.WASP_BASE_URL + '/public-api/ic/item/infosearch';
  var loggedFields = false;

  for (var pg = 1; pg <= 50; pg++) {
    var result = waspApiCall(url, { SearchPattern: '', PageSize: 100, PageNumber: pg });
    if (!result.success) { Logger.log('WASP infosearch page ' + pg + ' failed'); break; }
    var resp = JSON.parse(result.response);
    var rows = resp.Data || resp.data || [];
    if (rows.length === 0) break;

    if (!loggedFields && rows.length > 0) {
      Logger.log('WASP infosearch fields: ' + Object.keys(rows[0]).join(', '));
      loggedFields = true;
    }

    for (var i = 0; i < rows.length; i++) {
      var itemNo = rows[i].ItemNumber || rows[i].itemNumber || '';
      var uom = rows[i].StockingUnit || '';
      var category = rows[i].CategoryDescription || '';
      if (itemNo) waspMap[itemNo] = { uom: uom, category: category };
    }
    if (rows.length < 100) break;
    Utilities.sleep(200);
  }
  return waspMap;
}

/**
 * Normalize a UOM string to a canonical form for comparison.
 * Extends normalizeUom() with more cases.
 */
function normalizeUomForCompare_(uom) {
  if (!uom) return '';
  var u = (uom + '').trim().toLowerCase();
  if (u === 'ea' || u === 'each') return 'each';
  if (u === 'pc' || u === 'pcs' || u === 'piece' || u === 'pieces') return 'pc';
  if (u === 'g' || u === 'gram' || u === 'grams') return 'g';
  if (u === 'kg' || u === 'kilogram' || u === 'kilograms') return 'kg';
  if (u === 'ml' || u === 'milliliter' || u === 'millilitre') return 'ml';
  if (u === 'l' || u === 'liter' || u === 'litre') return 'l';
  if (u === 'oz' || u === 'ounce') return 'oz';
  if (u === 'lb' || u === 'lbs' || u === 'pound') return 'lb';
  if (u === 'roll' || u === 'rolls') return 'roll';
  return u;
}

/**
 * Audit: compare Katana UOM + category vs WASP StockingUnit + CategoryDescription for all items.
 * Logs a summary and full mismatch lists for UOM and category.
 * Run from GAS editor (09_Utils.gs): auditItemUoms()
 * @returns {Object} { uom: { matched, mismatched, missingInWasp, noUomInKatana },
 *                     category: { matched, mismatched } }
 */
function auditItemUoms() {
  Logger.log('=== auditItemUoms START ===');

  Logger.log('Building Katana map (products + materials + variants)...');
  var katanaMap = buildKatanaUomMap_();
  Logger.log('Katana: ' + Object.keys(katanaMap).length + ' SKUs');

  Logger.log('Building WASP map (infosearch — StockingUnit + CategoryDescription)...');
  var waspMap = buildWaspUomMap_();
  Logger.log('WASP: ' + Object.keys(waspMap).length + ' items');

  // UOM tracking
  var uomMatched = [];
  var uomMismatched = [];
  var missingInWasp = [];
  var noUomInKatana = [];

  // Category tracking
  var catMatched = [];
  var catMismatched = [];

  var skus = Object.keys(katanaMap);
  for (var i = 0; i < skus.length; i++) {
    var sku = skus[i];
    var kInfo = katanaMap[sku];
    var wInfo = waspMap[sku];

    if (!wInfo) {
      if (!kInfo.uom) {
        noUomInKatana.push(sku);
      } else {
        missingInWasp.push({ sku: sku, katanaUom: kInfo.uom, katanaCat: kInfo.category });
      }
      continue;
    }

    var kUomRaw = kInfo.uom;
    var wUomRaw = wInfo.uom || '';

    // ── UOM comparison ──
    if (!kUomRaw) {
      noUomInKatana.push(sku);
    } else {
      var conv = UOM_CONVERSIONS[sku];
      if (conv) {
        Logger.log('CONVERSION ' + sku + ': Katana=' + kUomRaw + ' WASP=' + wUomRaw + ' factor=' + conv.factor);
        uomMatched.push(sku);
      } else {
        var kUomNorm = normalizeUomForCompare_(kUomRaw);
        var wUomNorm = normalizeUomForCompare_(wUomRaw);
        if (kUomNorm === wUomNorm) {
          uomMatched.push(sku);
        } else {
          uomMismatched.push({ sku: sku, katana: kUomRaw, wasp: wUomRaw });
        }
      }
    }

    // ── Category comparison ──
    var kCat = (kInfo.category || '').trim();
    var wCat = (wInfo.category || '').trim();
    if (kCat && wCat) {
      if (kCat.toLowerCase() === wCat.toLowerCase()) {
        catMatched.push(sku);
      } else {
        catMismatched.push({ sku: sku, katana: kCat, wasp: wCat });
      }
    }
    // If either side has no category, skip silently (not an error we can fix)
  }

  Logger.log('\n=== UOM RESULTS ===');
  Logger.log('Matched:         ' + uomMatched.length);
  Logger.log('Mismatched:      ' + uomMismatched.length);
  Logger.log('Missing in WASP: ' + missingInWasp.length);
  Logger.log('No UOM (Katana): ' + noUomInKatana.length);

  if (uomMismatched.length > 0) {
    Logger.log('\n--- UOM MISMATCHES (fixable by bulkFixItemUomsAndCategories) ---');
    for (var m = 0; m < uomMismatched.length; m++) {
      var mm = uomMismatched[m];
      Logger.log('  ' + mm.sku + ':  Katana="' + mm.katana + '"  WASP="' + mm.wasp + '"');
    }
  }

  if (missingInWasp.length > 0) {
    Logger.log('\n--- MISSING IN WASP (not fixable — item not in WASP catalog) ---');
    for (var n = 0; n < missingInWasp.length; n++) {
      Logger.log('  ' + missingInWasp[n].sku + '  (Katana UOM: ' + missingInWasp[n].katanaUom + ')');
    }
  }

  if (noUomInKatana.length > 0) {
    Logger.log('\n--- NO UOM IN KATANA (review manually) ---');
    Logger.log('  ' + noUomInKatana.join(', '));
  }

  Logger.log('\n=== CATEGORY RESULTS ===');
  Logger.log('Matched:    ' + catMatched.length);
  Logger.log('Mismatched: ' + catMismatched.length);

  if (catMismatched.length > 0) {
    Logger.log('\n--- CATEGORY MISMATCHES (fixable by bulkFixItemUomsAndCategories) ---');
    for (var c = 0; c < catMismatched.length; c++) {
      var cm = catMismatched[c];
      Logger.log('  ' + cm.sku + ':  Katana="' + cm.katana + '"  WASP="' + cm.wasp + '"');
    }
  }

  Logger.log('\n=== END ===');
  return {
    uom: { matched: uomMatched, mismatched: uomMismatched, missingInWasp: missingInWasp, noUomInKatana: noUomInKatana },
    category: { matched: catMatched, mismatched: catMismatched }
  };
}

/**
 * Bulk fix: update WASP StockingUnit/PurchaseUnit/SalesUnit and CategoryDescription
 * to match Katana for all mismatched items.
 * Always run with dryRun=true first to preview changes.
 *
 * Run from GAS editor (09_Utils.gs):
 *   bulkFixItemUomsAndCategories()         ← dry run (safe, no changes)
 *   bulkFixItemUomsAndCategories(false)    ← live (makes changes in WASP)
 *
 * Alias: bulkFixItemUoms() still works for backwards compatibility.
 */
function bulkFixItemUomsAndCategories(dryRun) {
  if (dryRun === undefined || dryRun === null) dryRun = true;
  Logger.log('=== bulkFixItemUomsAndCategories ' + (dryRun ? '(DRY RUN)' : '(LIVE)') + ' ===');

  var audit = auditItemUoms();
  var uomMismatched = audit.uom.mismatched;
  var catMismatched = audit.category.mismatched;

  // Build a combined index: sku → { fixUom, fixCat, newUom, newCat }
  var toFix = {};

  for (var u = 0; u < uomMismatched.length; u++) {
    var um = uomMismatched[u];
    if (!toFix[um.sku]) toFix[um.sku] = {};
    toFix[um.sku].fixUom = true;
    toFix[um.sku].newUom = um.katana;
    toFix[um.sku].oldUom = um.wasp;
  }

  for (var c = 0; c < catMismatched.length; c++) {
    var cm = catMismatched[c];
    if (!toFix[cm.sku]) toFix[cm.sku] = {};
    toFix[cm.sku].fixCat = true;
    toFix[cm.sku].newCat = cm.katana;
    toFix[cm.sku].oldCat = cm.wasp;
  }

  var skus = Object.keys(toFix);
  if (skus.length === 0) {
    Logger.log('All UOMs and categories match — nothing to fix.');
    return;
  }

  Logger.log('\n' + skus.length + ' items to update:');

  var url = CONFIG.WASP_BASE_URL + '/public-api/ic/item/updateInventoryItems';
  var success = 0;
  var failed = 0;

  for (var i = 0; i < skus.length; i++) {
    var sku = skus[i];
    var fix = toFix[sku];
    var parts = [];
    if (fix.fixUom) parts.push('UOM: "' + fix.oldUom + '" → "' + fix.newUom + '"');
    if (fix.fixCat) parts.push('Cat: "' + fix.oldCat + '" → "' + fix.newCat + '"');
    Logger.log((dryRun ? '[DRY RUN] ' : '') + sku + '  ' + parts.join('  |  '));
    if (dryRun) continue;

    var payload = {
      ItemNumber:    sku,
      DimensionInfo: { DimensionUnit: 'Inch', WeightUnit: 'Pound', VolumeUnit: 'Cubic Inch' }
    };
    if (fix.fixUom) {
      payload.StockingUnit = fix.newUom;
      payload.PurchaseUnit = fix.newUom;
      payload.SalesUnit    = fix.newUom;
    }
    if (fix.fixCat) {
      payload.CategoryDescription = fix.newCat;
    }

    var r = waspApiCall(url, [payload]);
    if (r.success) {
      success++;
      Logger.log('  OK');
    } else {
      failed++;
      Logger.log('  FAILED: ' + String(r.response).substring(0, 200));
    }
    Utilities.sleep(200);
  }

  if (!dryRun) {
    Logger.log('\n=== DONE — Updated: ' + success + ' | Failed: ' + failed + ' ===');
  } else {
    Logger.log('\nDry run complete — run bulkFixItemUomsAndCategories(false) to apply changes.');
  }
}

// Backwards-compatible alias
function bulkFixItemUoms(dryRun) {
  return bulkFixItemUomsAndCategories(dryRun);
}

/**
 * Test: fetch current UOM from WASP via infosearch, then write it back unchanged.
 * This confirms: (a) infosearch returns StockingUnit, (b) UOM can be updated on items with stock.
 * Change SKU_WITH_STOCK to any item you know has inventory.
 *
 * Run from GAS editor (09_Utils.gs): testUpdateItemUomWithStock()
 */
function testUpdateItemUomWithStock() {
  var SKU_WITH_STOCK = 'UFC-4OZ'; // ← change to any item with stock > 0

  Logger.log('=== testUpdateItemUomWithStock ===');
  Logger.log('SKU: ' + SKU_WITH_STOCK);

  // Step 1: Read current UOM from WASP via infosearch
  var searchUrl = CONFIG.WASP_BASE_URL + '/public-api/ic/item/infosearch';
  var sr = waspApiCall(searchUrl, { SearchPattern: SKU_WITH_STOCK, PageNumber: 1, PageSize: 5 });
  Logger.log('infosearch HTTP: ' + sr.code + '  success: ' + sr.success);
  if (!sr.success) {
    Logger.log('FAILED to fetch item: ' + String(sr.response).substring(0, 300));
    return;
  }

  var resp = JSON.parse(sr.response);
  var rows = resp.Data || resp.data || [];
  Logger.log('Rows returned: ' + rows.length);
  if (rows.length === 0) {
    Logger.log('Item not found in WASP catalog.');
    return;
  }

  var item = rows[0];
  var currentUom      = item.StockingUnit || '';
  var currentCategory = item.CategoryDescription || '';
  Logger.log('Current StockingUnit:      "' + currentUom + '"');
  Logger.log('Current CategoryDescription: "' + currentCategory + '"');

  // Step 2: Write the same UOM back (no-op test — stock is unaffected)
  Logger.log('\nWriting same UOM back (no-op) to confirm API accepts update on item with stock...');
  var updateUrl = CONFIG.WASP_BASE_URL + '/public-api/ic/item/updateInventoryItems';
  var ur = waspApiCall(updateUrl, [{
    ItemNumber:    SKU_WITH_STOCK,
    StockingUnit:  currentUom,
    PurchaseUnit:  currentUom,
    SalesUnit:     currentUom,
    DimensionInfo: { DimensionUnit: 'Inch', WeightUnit: 'Pound', VolumeUnit: 'Cubic Inch' }
  }]);

  Logger.log('Update HTTP: ' + ur.code + '  success: ' + ur.success);
  Logger.log('Response: ' + String(ur.response).substring(0, 400));

  if (ur.success) {
    Logger.log('\nCONFIRMED: WASP allows UOM update on items with stock.');
  } else {
    Logger.log('\nFAILED — check error above.');
  }
}

// ===================== HEALTH MONITOR HEARTBEAT =====================

/**
 * Reports run status to the Systems Health Monitor web app.
 * Requires Script Property: HEALTH_MONITOR_URL
 * Silent — never throws, never blocks execution.
 */
function sendHeartbeatToMonitor_(systemName, status, details) {
  try {
    var url = PropertiesService.getScriptProperties().getProperty('HEALTH_MONITOR_URL');
    if (!url) return;
    var payload = (status === 'success')
      ? { system: systemName, status: 'success', details: details || '' }
      : { system: systemName, error: details || 'Unknown error', severity: 'error' };
    UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
  } catch (hbErr) {
    Logger.log('Heartbeat send failed (non-blocking): ' + hbErr.message);
  }
}
