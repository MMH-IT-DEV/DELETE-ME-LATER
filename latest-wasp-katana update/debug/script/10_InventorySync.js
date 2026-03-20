// ============================================
// 10_InventorySync.gs - SHEET-BASED INVENTORY SYNC
// ============================================
// Full reset sync: Zero WASP → Re-Add from Katana
// Uses the Sync Google Sheet for plans and results.
//
// Run order:
//   1. syncBuildReAddPlan()  — Fetch Katana, write Re-Add Plan tab
//   2. syncZeroFromSheet()   — Execute Zero Plan (zero all WASP items)
//   3. syncReAddFromSheet()  — Execute Re-Add Plan (add from Katana)
//
// Dependencies (defined in other .gs files):
//   syncBuildKatanaMap()        — 11_SyncHelpers.gs
//   syncWaspCallWithRetry_()    — 11_SyncHelpers.gs
//   waspApiCall()               — 05_WaspAPI.gs
//   waspAddInventoryWithLot()   — 05_WaspAPI.gs
//   waspRemoveInventoryWithLot()— 05_WaspAPI.gs
//   CONFIG, SYNC_CONFIG         — 00_Config.gs
// ============================================

// Data starts at row 7 (index 6) due to title rows 1-5
var SYNC_DATA_ROW = 7;

// ============================================
// PHASE 1: ZERO WASP INVENTORY FROM SHEET
// ============================================

/**
 * Read the Zero Plan tab and zero out each PENDING item in WASP.
 *
 * For items WITH a Lot value: uses item/remove endpoint.
 * For items WITHOUT a Lot: uses item/adjust with negative qty.
 *
 * Updates the Status column (H) in Zero Plan after each call.
 * Writes every result to the Results tab.
 */
function syncZeroFromSheet() {
  Logger.log('===== SYNC ZERO PHASE =====');
  var ss = SpreadsheetApp.openById(SYNC_CONFIG.SYNC_SHEET_ID);
  var zeroPlan = ss.getSheetByName('Zero Plan');
  var resultsTab = ss.getSheetByName('Results');

  if (!zeroPlan) {
    Logger.log('ERROR: Zero Plan tab not found');
    return;
  }

  var data = zeroPlan.getDataRange().getValues();
  // Row 0 = headers: SKU(0), Description(1), Site(2), Location(3),
  //                   Current Qty(4), Adjustment(5), Lot(6), Status(7)

  var adjustUrl = CONFIG.WASP_BASE_URL + '/public-api/transactions/item/adjust';
  var today = new Date().toISOString().split('T')[0];
  var processed = 0;
  var okCount = 0;
  var errCount = 0;
  var skipCount = 0;

  for (var i = SYNC_DATA_ROW - 1; i < data.length; i++) {
    var row = data[i];
    var sku = String(row[0]).trim();
    var site = String(row[2]).trim();
    var location = String(row[3]).trim();
    var currentQty = Number(row[4]) || 0;
    var adjustment = Number(row[5]) || 0;
    var lot = String(row[6] || '').trim();
    var status = String(row[7]).trim();

    // Only process PENDING rows
    if (status !== 'PENDING') {
      skipCount++;
      continue;
    }
    if (adjustment === 0 || currentQty === 0) {
      zeroPlan.getRange(i + 1, 8).setValue('SKIP');
      skipCount++;
      continue;
    }

    var result;
    var newStatus;
    var error;

    if (lot && lot !== '' && lot !== 'undefined') {
      // Lot-tracked: use remove endpoint (qty is positive for remove)
      var removeQty = Math.abs(adjustment);
      result = waspRemoveInventoryWithLot(sku, removeQty, location, lot, 'Sync zero ' + today, site);
    } else {
      // Non-lot: use adjust endpoint with negative qty
      var payload = [{
        ItemNumber: sku,
        Quantity: adjustment,
        SiteName: site,
        LocationCode: location,
        Notes: 'Sync zero ' + today
      }];
      result = waspApiCall(adjustUrl, payload);
    }

    if (result.success) {
      newStatus = 'OK';
      error = '';
      okCount++;
    } else {
      newStatus = 'ERROR';
      error = (result.response || '').substring(0, 200);
      errCount++;
    }

    // Update Status in Zero Plan (column H = col 8)
    zeroPlan.getRange(i + 1, 8).setValue(newStatus);

    // Write to Results tab
    syncAppendResult_(resultsTab, 'ZERO', sku, site, location, adjustment, lot, newStatus, error, result.response);

    processed++;
    Utilities.sleep(SYNC_CONFIG.RATE_LIMIT_MS);

    if (processed % 25 === 0) {
      Logger.log('  Zero progress: ' + processed + ' processed, OK=' + okCount + ', ERR=' + errCount);
      SpreadsheetApp.flush();
    }

    if (processed >= SYNC_CONFIG.MAX_EXECUTE_ITEMS) {
      Logger.log('  Reached MAX_EXECUTE_ITEMS limit (' + SYNC_CONFIG.MAX_EXECUTE_ITEMS + ')');
      break;
    }
  }

  Logger.log('===== ZERO PHASE COMPLETE =====');
  Logger.log('Processed=' + processed + ', OK=' + okCount + ', ERROR=' + errCount + ', Skipped=' + skipCount);
}

// ============================================
// PHASE 2a: BUILD RE-ADD PLAN FROM KATANA
// ============================================

/**
 * Fetch current Katana inventory and write the Re-Add Plan tab.
 * Uses syncBuildKatanaMap() to get qty per SKU|site|location,
 * then writes each entry as a PENDING row.
 *
 * Note: Lot and DateCode columns are left blank — Katana doesn't
 * provide lot info in the inventory API. These can be manually
 * filled or auto-populated from WASP item config.
 */
function syncBuildReAddPlan() {
  Logger.log('===== BUILD RE-ADD PLAN =====');
  var ss = SpreadsheetApp.openById(SYNC_CONFIG.SYNC_SHEET_ID);
  var reAddPlan = ss.getSheetByName('Re-Add Plan');

  if (!reAddPlan) {
    Logger.log('ERROR: Re-Add Plan tab not found');
    return;
  }

  // Build Katana map: { 'SKU|site|location': qty }
  var katanaMap = syncBuildKatanaMap();
  var keys = Object.keys(katanaMap);
  Logger.log('Katana map: ' + keys.length + ' entries');

  // Read Item Map for descriptions and categories (optional enrichment)
  var itemMap = {};
  var itemMapTab = ss.getSheetByName('Item Map');
  if (itemMapTab) {
    var mapData = itemMapTab.getDataRange().getValues();
    // Headers: SKU(0), Description(1), Category(2), Target Site(3), Target Location(4), Action(5), Notes(6)
    for (var m = SYNC_DATA_ROW - 1; m < mapData.length; m++) {
      var mapSku = String(mapData[m][0]).trim();
      if (mapSku) {
        itemMap[mapSku] = {
          description: mapData[m][1] || '',
          category: mapData[m][2] || ''
        };
      }
    }
    Logger.log('Item Map loaded: ' + Object.keys(itemMap).length + ' entries');
  }

  // Build rows for Re-Add Plan
  // Headers: SKU(A), Description(B), Category(C), Target Site(D), Target Location(E),
  //          Katana Qty(F), Lot(G), DateCode(H), Adjustment(I), Status(J)
  var rows = [];
  for (var k = 0; k < keys.length; k++) {
    var key = keys[k];
    var parts = key.split('|');
    var sku = parts[0];
    var site = parts[1];
    var location = parts[2];
    var qty = katanaMap[key];

    if (qty <= 0) continue;

    var info = itemMap[sku] || { description: '', category: '' };

    rows.push([
      sku,
      info.description,
      info.category,
      site,
      location,
      qty,
      '',       // Lot — blank, Katana doesn't provide
      '',       // DateCode — blank
      qty,      // Adjustment = same as Katana Qty for full reset
      'PENDING'
    ]);
  }

  // Sort by site, then location, then SKU
  rows.sort(function(a, b) {
    var cmp = String(a[3]).localeCompare(String(b[3]));
    if (cmp !== 0) return cmp;
    cmp = String(a[4]).localeCompare(String(b[4]));
    if (cmp !== 0) return cmp;
    return String(a[0]).localeCompare(String(b[0]));
  });

  // Clear existing data (keep headers)
  if (reAddPlan.getLastRow() >= SYNC_DATA_ROW) {
    reAddPlan.getRange(SYNC_DATA_ROW, 1, reAddPlan.getLastRow() - SYNC_DATA_ROW + 1, 10).clearContent();
  }

  // Write data
  if (rows.length > 0) {
    reAddPlan.getRange(SYNC_DATA_ROW, 1, rows.length, 10).setValues(rows);
  }

  SpreadsheetApp.flush();
  Logger.log('===== RE-ADD PLAN COMPLETE =====');
  Logger.log('Rows written: ' + rows.length);
}

// ============================================
// PHASE 2b: EXECUTE RE-ADD FROM SHEET
// ============================================

/**
 * Read the Re-Add Plan tab and add each PENDING item to WASP.
 *
 * For items WITH a Lot value: uses item/add endpoint (waspAddInventoryWithLot).
 * For items WITHOUT a Lot: uses item/adjust with positive qty.
 *
 * Updates the Status column (J) in Re-Add Plan after each call.
 * Writes every result to the Results tab.
 */
function syncReAddFromSheet() {
  Logger.log('===== SYNC RE-ADD PHASE =====');
  var ss = SpreadsheetApp.openById(SYNC_CONFIG.SYNC_SHEET_ID);
  var reAddPlan = ss.getSheetByName('Re-Add Plan');
  var resultsTab = ss.getSheetByName('Results');

  if (!reAddPlan) {
    Logger.log('ERROR: Re-Add Plan tab not found');
    return;
  }

  var data = reAddPlan.getDataRange().getValues();
  // Row 0 = headers: SKU(0), Description(1), Category(2), Target Site(3),
  //   Target Location(4), Katana Qty(5), Lot(6), DateCode(7), Adjustment(8), Status(9)

  var adjustUrl = CONFIG.WASP_BASE_URL + '/public-api/transactions/item/adjust';
  var today = new Date().toISOString().split('T')[0];
  var processed = 0;
  var okCount = 0;
  var errCount = 0;
  var skipCount = 0;

  for (var i = SYNC_DATA_ROW - 1; i < data.length; i++) {
    var row = data[i];
    var sku = String(row[0]).trim();
    var site = String(row[3]).trim();
    var location = String(row[4]).trim();
    var adjustment = Number(row[8]) || 0;
    var lot = String(row[6] || '').trim();
    var dateCode = String(row[7] || '').trim();
    var status = String(row[9]).trim();

    if (status !== 'PENDING') {
      skipCount++;
      continue;
    }
    if (adjustment === 0) {
      reAddPlan.getRange(i + 1, 10).setValue('SKIP');
      skipCount++;
      continue;
    }

    var result;
    var newStatus;
    var error;

    if (lot && lot !== '' && lot !== 'undefined') {
      // Lot-tracked: use add endpoint
      result = waspAddInventoryWithLot(sku, adjustment, location, lot, dateCode, 'Sync add ' + today, site);
    } else {
      // Non-lot: use adjust endpoint with positive qty
      var payload = [{
        ItemNumber: sku,
        Quantity: adjustment,
        SiteName: site,
        LocationCode: location,
        Notes: 'Sync add ' + today
      }];
      result = waspApiCall(adjustUrl, payload);
    }

    if (result.success) {
      newStatus = 'OK';
      error = '';
      okCount++;
    } else {
      newStatus = 'ERROR';
      error = (result.response || '').substring(0, 200);
      errCount++;
    }

    // Update Status in Re-Add Plan (column J = col 10)
    reAddPlan.getRange(i + 1, 10).setValue(newStatus);

    // Write to Results tab
    syncAppendResult_(resultsTab, 'RE-ADD', sku, site, location, adjustment, lot, newStatus, error, result.response);

    processed++;
    Utilities.sleep(SYNC_CONFIG.RATE_LIMIT_MS);

    if (processed % 25 === 0) {
      Logger.log('  Re-Add progress: ' + processed + ' processed, OK=' + okCount + ', ERR=' + errCount);
      SpreadsheetApp.flush();
    }

    if (processed >= SYNC_CONFIG.MAX_EXECUTE_ITEMS) {
      Logger.log('  Reached MAX_EXECUTE_ITEMS limit (' + SYNC_CONFIG.MAX_EXECUTE_ITEMS + ')');
      break;
    }
  }

  Logger.log('===== RE-ADD PHASE COMPLETE =====');
  Logger.log('Processed=' + processed + ', OK=' + okCount + ', ERROR=' + errCount + ', Skipped=' + skipCount);
}

// ============================================
// PRIVATE HELPERS
// ============================================

/**
 * Append a single result row to the Results tab.
 */
function syncAppendResult_(resultsTab, phase, sku, site, location, qtyChange, lot, status, error, rawResponse) {
  if (!resultsTab) return;
  var nextRow = resultsTab.getLastRow() + 1;
  var responseSnippet = '';
  if (rawResponse) {
    responseSnippet = String(rawResponse).substring(0, 300);
  }
  resultsTab.getRange(nextRow, 1, 1, 10).setValues([[
    new Date().toISOString(),
    phase,
    sku,
    site,
    location,
    qtyChange,
    lot || '',
    status,
    error || '',
    responseSnippet
  ]]);
}

/**
 * Utility: count rows with a given status in a specific column of a sheet.
 * Useful for progress checking.
 */
function syncCountStatus_(sheet, statusCol, statusValue) {
  var data = sheet.getDataRange().getValues();
  var count = 0;
  for (var i = SYNC_DATA_ROW - 1; i < data.length; i++) {
    if (String(data[i][statusCol]).trim() === statusValue) {
      count++;
    }
  }
  return count;
}

/**
 * Quick status check — run from GAS editor to see progress.
 */
function syncCheckProgress() {
  var ss = SpreadsheetApp.openById(SYNC_CONFIG.SYNC_SHEET_ID);

  var zeroPlan = ss.getSheetByName('Zero Plan');
  if (zeroPlan) {
    Logger.log('=== Zero Plan ===');
    Logger.log('  PENDING: ' + syncCountStatus_(zeroPlan, 7, 'PENDING'));
    Logger.log('  OK:      ' + syncCountStatus_(zeroPlan, 7, 'OK'));
    Logger.log('  ERROR:   ' + syncCountStatus_(zeroPlan, 7, 'ERROR'));
    Logger.log('  SKIP:    ' + syncCountStatus_(zeroPlan, 7, 'SKIP'));
  }

  var reAddPlan = ss.getSheetByName('Re-Add Plan');
  if (reAddPlan) {
    Logger.log('=== Re-Add Plan ===');
    Logger.log('  PENDING: ' + syncCountStatus_(reAddPlan, 9, 'PENDING'));
    Logger.log('  OK:      ' + syncCountStatus_(reAddPlan, 9, 'OK'));
    Logger.log('  ERROR:   ' + syncCountStatus_(reAddPlan, 9, 'ERROR'));
    Logger.log('  SKIP:    ' + syncCountStatus_(reAddPlan, 9, 'SKIP'));
  }
}
