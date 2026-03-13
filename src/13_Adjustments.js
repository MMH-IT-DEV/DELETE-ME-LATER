// ============================================
// 13_Adjustments.gs - MANUAL STOCK ADJUSTMENTS
// ============================================
// Processes manual corrections from the "Adjustments" sheet
// in the debug spreadsheet. Allows fixing failed F4/F1 items,
// adding missing lots, correcting quantities, etc.
//
// Sheet layout (Adjustments tab):
//   A(1)=Action  B(2)=SKU  C(3)=Qty  D(4)=Site  E(5)=Location
//   F(6)=Lot  G(7)=Expiry  H(8)=Notes  I(9)=Status  J(10)=Error
//
// Actions: ADD, REMOVE, ADJUST (positive=add, negative=remove)
// Status: PENDING → Complete/Failed (auto-set by processAdjustments)
//
// Run: Execute processAdjustments() from script editor or menu.
// ============================================

/**
 * Create/get the Adjustments sheet with proper headers.
 * Returns the sheet object.
 */
function getAdjustmentsSheet() {
  var ss = SpreadsheetApp.openById(CONFIG.DEBUG_SHEET_ID);
  var sheet = ss.getSheetByName('Adjustments');

  if (!sheet) {
    sheet = ss.insertSheet('Adjustments');

    // Row 1: Title
    sheet.getRange(1, 1, 1, 10).setValues([
      ['STOCK ADJUSTMENTS', '', '', '', '', '', '', '', '', '']
    ]);
    sheet.getRange(1, 1).setFontSize(12).setFontWeight('bold');

    // Row 2: Instructions
    sheet.getRange(2, 1, 1, 10).setValues([
      ['Enter rows below, set Status to PENDING, then run processAdjustments()', '', '', '', '', '', '', '', '', '']
    ]);
    sheet.getRange(2, 1).setFontColor('#666666');

    // Row 3: Headers
    sheet.getRange(3, 1, 1, 10).setValues([
      ['Action', 'SKU', 'Qty', 'Site', 'Location', 'Lot', 'Expiry', 'Notes', 'Status', 'Error']
    ]);
    sheet.getRange(3, 1, 1, 10).setFontWeight('bold').setBackground('#e8eaf6');

    sheet.setFrozenRows(3);

    // Column widths
    sheet.setColumnWidth(1, 80);   // Action
    sheet.setColumnWidth(2, 140);  // SKU
    sheet.setColumnWidth(3, 60);   // Qty
    sheet.setColumnWidth(4, 120);  // Site
    sheet.setColumnWidth(5, 120);  // Location
    sheet.setColumnWidth(6, 140);  // Lot
    sheet.setColumnWidth(7, 100);  // Expiry
    sheet.setColumnWidth(8, 200);  // Notes
    sheet.setColumnWidth(9, 80);   // Status
    sheet.setColumnWidth(10, 250); // Error

    // Add data validation for Action column (rows 4-100)
    var actionRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['ADD', 'REMOVE', 'ADJUST'], true)
      .setAllowInvalid(false)
      .build();
    sheet.getRange(4, 1, 97, 1).setDataValidation(actionRule);

    // Add data validation for Status column (rows 4-100)
    var statusRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['PENDING', 'Complete', 'Failed', 'Skip'], true)
      .setAllowInvalid(false)
      .build();
    sheet.getRange(4, 9, 97, 1).setDataValidation(statusRule);

    // Default site
    sheet.getRange(4, 4).setValue(CONFIG.WASP_SITE);
    sheet.getRange(4, 5).setValue('PRODUCTION');
    sheet.getRange(4, 9).setValue('PENDING');

    Logger.log('Created Adjustments sheet');
  }

  return sheet;
}

/**
 * Process all PENDING rows in the Adjustments sheet.
 * Calls WASP API for each row and updates Status/Error columns.
 * Logs results to Activity Log as F2 Adjustments.
 *
 * Run this from the GAS script editor or set up a menu/button.
 */
function processAdjustments() {
  var sheet = getAdjustmentsSheet();
  var lastRow = sheet.getLastRow();

  if (lastRow < 4) {
    Logger.log('No adjustment rows to process');
    return;
  }

  var data = sheet.getRange(4, 1, lastRow - 3, 10).getValues();
  var pendingRows = [];

  // Find all PENDING rows
  for (var i = 0; i < data.length; i++) {
    var status = String(data[i][8]).trim();
    if (status === 'PENDING') {
      pendingRows.push(i);
    }
  }

  if (pendingRows.length === 0) {
    Logger.log('No PENDING adjustments found');
    return;
  }

  Logger.log('Processing ' + pendingRows.length + ' adjustments');

  var successCount = 0;
  var failCount = 0;
  var f2SubItems = [];
  var f2UomCache = {};
  var dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');

  for (var p = 0; p < pendingRows.length; p++) {
    var idx = pendingRows[p];
    var row = data[idx];
    var sheetRow = idx + 4; // 1-based sheet row (data starts row 4)

    var action = String(row[0]).trim().toUpperCase();
    var sku = String(row[1]).trim();
    var qty = parseFloat(row[2]) || 0;
    var site = String(row[3]).trim() || CONFIG.WASP_SITE;
    var location = String(row[4]).trim();
    var lot = String(row[5] || '').trim();
    var expiry = '';
    var notes = String(row[7] || '').trim() || ('Manual adjustment ' + dateStr);

    // Handle expiry — could be Date object or string
    expiry = normalizeBusinessDate_(row[6]);

    // Validate required fields
    if (!sku) {
      sheet.getRange(sheetRow, 9).setValue('Failed');
      sheet.getRange(sheetRow, 10).setValue('SKU is required');
      failCount++;
      continue;
    }
    if (!action || (action !== 'ADD' && action !== 'REMOVE' && action !== 'ADJUST')) {
      sheet.getRange(sheetRow, 9).setValue('Failed');
      sheet.getRange(sheetRow, 10).setValue('Action must be ADD, REMOVE, or ADJUST');
      failCount++;
      continue;
    }
    if (qty === 0 && action !== 'ADJUST') {
      sheet.getRange(sheetRow, 9).setValue('Failed');
      sheet.getRange(sheetRow, 10).setValue('Qty cannot be zero');
      failCount++;
      continue;
    }
    if (!location) {
      sheet.getRange(sheetRow, 9).setValue('Failed');
      sheet.getRange(sheetRow, 10).setValue('Location is required');
      failCount++;
      continue;
    }

    // Determine effective action for ADJUST type
    var effectiveAction = action;
    var effectiveQty = Math.abs(qty);
    if (action === 'ADJUST') {
      if (qty > 0) {
        effectiveAction = 'ADD';
      } else if (qty < 0) {
        effectiveAction = 'REMOVE';
      } else {
        sheet.getRange(sheetRow, 9).setValue('Skip');
        sheet.getRange(sheetRow, 10).setValue('Zero adjustment');
        continue;
      }
    }

    Logger.log('Adjustment: ' + effectiveAction + ' ' + sku + ' x' + effectiveQty + ' at ' + site + '/' + location + (lot ? ' lot=' + lot : ''));

    // Mark sync to prevent echo loops (WASP callout → F2)
    markSyncedToWasp(sku, location, effectiveAction === 'ADD' ? 'add' : 'remove');

    // Execute WASP API call
    var result;
    if (effectiveAction === 'ADD') {
      if (lot) {
        result = waspAddInventoryWithLot(sku, effectiveQty, location, lot, expiry, notes, site);
      } else {
        result = waspAddInventory(sku, effectiveQty, location, notes, site);
      }
    } else {
      if (lot) {
        result = waspRemoveInventoryWithLot(sku, effectiveQty, location, lot, notes, site, expiry);
      } else {
        result = waspRemoveInventory(sku, effectiveQty, location, notes, site);
      }
    }

    // Update sheet
    if (result.success) {
      sheet.getRange(sheetRow, 9).setValue('Complete');
      sheet.getRange(sheetRow, 10).setValue('');
      sheet.getRange(sheetRow, 1, 1, 10).setBackground('#e8f5e9'); // light green
      successCount++;
    } else {
      var errMsg = '';
      if (result.response) {
        errMsg = cleanErrorMessage(parseWaspError(result.response, effectiveAction, sku));
      } else {
        errMsg = result.error || 'Unknown error';
      }
      sheet.getRange(sheetRow, 9).setValue('Failed');
      sheet.getRange(sheetRow, 10).setValue(errMsg);
      sheet.getRange(sheetRow, 1, 1, 10).setBackground('#ffebee'); // light red
      failCount++;
    }

    // Adjustments Log — record every sync-sheet adjustment
    var adjDiff = effectiveAction === 'ADD' ? effectiveQty : -effectiveQty;
    var adjAdjStatus = result.success ? 'OK' : 'ERROR';
    var adjErrNote = result.success ? notes : (errMsg || result.error || 'WASP error');
    logAdjustment('Sync Sheet', effectiveAction === 'ADD' ? 'Add' : 'Remove', '', sku, '', site, location, lot, expiry, adjDiff, '', null, adjAdjStatus);

    // Build sub-item for Activity Log
    var subAction = buildActivityActionText_(
      getActivityDisplayLocation_(location) + '  ' + (effectiveAction === 'ADD' ? 'add' : 'remove'),
      lot,
      expiry
    );

    f2SubItems.push({
      sku: sku,
      qty: (effectiveAction === 'REMOVE' ? '-' : '') + effectiveQty,
      uom: resolveSkuUom(sku, f2UomCache),
      success: result.success,
      status: result.success ? (effectiveAction === 'ADD' ? 'Added' : 'Removed') : 'Failed',
      error: result.success ? '' : (result.response ? parseWaspError(result.response, effectiveAction, sku) : ''),
      action: subAction,
      qtyColor: effectiveAction === 'ADD' ? 'green' : 'red'
    });
  }

  // Log to Activity as F2 Adjustments
  var f2Status = failCount === 0 ? 'success' : successCount === 0 ? 'failed' : 'partial';
  var f2Detail = joinActivitySegments_(['Sync Sheet', buildActivityCountSummary_(pendingRows.length, 'adjustment', 'adjustments', '')]);

  logActivity('F2', f2Detail, f2Status, 'Sync Sheet -> WASP', f2SubItems);

  Logger.log('Adjustments complete: ' + successCount + ' success, ' + failCount + ' failed');
}

// parseWaspError() defined in 07_F3_AutoPoll.gs — reused here

/**
 * Quick-add an adjustment row from code (for retry scenarios).
 * Adds a PENDING row to the Adjustments sheet.
 */
function addAdjustmentRow(action, sku, qty, site, location, lot, expiry, notes) {
  var sheet = getAdjustmentsSheet();
  var newRow = sheet.getLastRow() + 1;
  sheet.getRange(newRow, 1, 1, 10).setValues([
    [action, sku, qty, site || CONFIG.WASP_SITE, location, lot || '', expiry || '', notes || '', 'PENDING', '']
  ]);
  return newRow;
}

// ============================================
// RETRY FROM ACTIVITY LOG (CHECKBOX IN COLUMN G)
// ============================================

/**
 * Process all checked Retry checkboxes in the Activity Log.
 * Supports all flows (F1-F6). Header-level retry: checking a header row
 * retries ALL its failed sub-items. Uses getRetryAction() for flow-aware
 * ADD/REMOVE determination and success status mapping.
 * Run from GAS menu: Katana-WASP > Retry Failed
 */
function retryMarkedItems() {
  var ss = SpreadsheetApp.openById(CONFIG.DEBUG_SHEET_ID);
  var sheet = ss.getSheetByName('Activity');
  if (!sheet) { Logger.log('No Activity sheet'); return; }

  var lastRow = sheet.getLastRow();
  if (lastRow < 4) { Logger.log('Empty Activity sheet'); return; }

  var data = sheet.getRange(1, 1, lastRow, 7).getValues();

  // Header-level: checking a WK-xxx row expands to all Failed/Skipped sub-items
  var checkedSet = {};
  for (var i = 3; i < data.length; i++) {
    if (data[i][6] !== true) continue;
    var colA = String(data[i][0]).trim();
    if (colA.indexOf('WK-') === 0) {
      sheet.getRange(i + 1, 7).setValue(false);
      for (var j = i + 1; j < data.length; j++) {
        var subColA = String(data[j][0]).trim();
        if (subColA.indexOf('WK-') === 0) break;
        var subSt = String(data[j][4]).trim();
        if (subSt === 'Failed' || subSt === 'Skipped') { checkedSet[j] = true; }
      }
    } else { checkedSet[i] = true; }
  }
  var checkedRows = [];
  for (var key in checkedSet) { checkedRows.push(parseInt(key, 10)); }
  checkedRows.sort(function(a, b) { return a - b; });
  if (checkedRows.length === 0) {
    Logger.log('No Retry checkboxes checked. Check column G on Failed/Skipped rows, then run again.');
    return;
  }
  Logger.log('Retrying ' + checkedRows.length + ' checked items');
  var successCount = 0;
  var failCount = 0;
  var f2SubItems = [];
  var retryUomCache = {};
  var touchedHeaders = {};

  for (var c = 0; c < checkedRows.length; c++) {
    var rowIdx = checkedRows[c];
    var sheetRow = rowIdx + 1; // 1-based
    var subLine = String(data[rowIdx][3]).trim();
    var subStatus = String(data[rowIdx][4]).trim();
    var subError = String(data[rowIdx][5]).trim();

    // Parse sub-item line
    var parsed = parseActivitySubItem(subLine);
    if (!parsed || !parsed.sku) {
      Logger.log('Row ' + sheetRow + ': Could not parse — skipping');
      sheet.getRange(sheetRow, 7).setValue(false);
      continue;
    }

    // Find parent header row (search upward for WK-ID)
    var parentHeader = findParentHeader(data, rowIdx);
    var moBatch = parentHeader.batch || '';
    var moExpiry = parentHeader.expiry || '';
    var outputSku = parentHeader.outputSku || '';
    var headerFlow = parentHeader.flow || '';

    // Track parent for post-retry header status update
    if (parentHeader.execId) { touchedHeaders[parentHeader.execId] = parentHeader; }

    // Determine action based on flow type
    var isOutput = (parsed.sku === outputSku);
    var retryInfo = getRetryAction(headerFlow, isOutput);
    var action = retryInfo.action;
    var successStatus = retryInfo.successStatus;
    var location = parsed.location || retryInfo.defaultLocation || 'PRODUCTION';
    var lot = parsed.lot || '';
    var expiry = parsed.expiry || '';
    var site = CONFIG.WASP_SITE;

    // For F4 output: use MO batch as lot
    if (headerFlow.indexOf('F4') >= 0 && isOutput && !lot) {
      lot = moBatch;
      if (!expiry) expiry = moExpiry;
    }

    // For items needing lot: look up from WASP
    if (!lot && subError.indexOf('Lot') >= 0) {
      Logger.log('Looking up WASP lot for ' + parsed.sku + ' at ' + location);
      var lotInfo = waspLookupItemLotAndDate(parsed.sku, location, site);
      if (lotInfo) {
        lot = lotInfo.lot || '';
        if (!expiry) expiry = lotInfo.dateCode || '';
        Logger.log('  Found lot=' + lot + ' exp=' + expiry);
      } else {
        Logger.log('  No lot found in WASP — will try without');
      }
    }

    // Mark sync to prevent echo
    markSyncedToWasp(parsed.sku, location, action === 'ADD' ? 'add' : 'remove');

    // Execute WASP API call
    var notes = 'Retry from Activity ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
    var result;

    Logger.log('RETRY: ' + action + ' ' + parsed.sku + ' x' + parsed.qty + ' at ' + site + '/' + location + (lot ? ' lot=' + lot : ''));

    if (action === 'ADD') {
      if (lot) {
        result = waspAddInventoryWithLot(parsed.sku, parsed.qty, location, lot, expiry, notes, site);
      } else {
        result = waspAddInventory(parsed.sku, parsed.qty, location, notes, site);
      }
    } else {
      if (lot) {
        result = waspRemoveInventoryWithLot(parsed.sku, parsed.qty, location, lot, notes, site, expiry);
      } else {
        result = waspRemoveInventory(parsed.sku, parsed.qty, location, notes, site);
      }
    }

    // Update the Activity row in-place
    var errMsg = '';
    if (result.success) {
      sheet.getRange(sheetRow, 5).setValue(successStatus);   // Status
      sheet.getRange(sheetRow, 6).setValue('');            // Clear error
      sheet.getRange(sheetRow, 7).setValue(false);         // Uncheck
      // Clear red/pink background
      sheet.getRange(sheetRow, 1, 1, 7).setBackground(null);
      successCount++;
      Logger.log('  OK → ' + successStatus);
    } else {
      errMsg = result.response ? parseWaspError(result.response, action, parsed.sku, location) : 'Unknown error';
      sheet.getRange(sheetRow, 6).setValue(cleanErrorMessage(errMsg)); // Update error
      sheet.getRange(sheetRow, 7).setValue(false);         // Uncheck
      failCount++;
      Logger.log('  FAIL: ' + errMsg);
    }

    // Build sub-item for F2 Activity entry
    var subAction = buildActivityActionText_(
      getActivityDisplayLocation_(location) + '  ' + (action === 'ADD' ? 'add' : 'remove'),
      lot,
      expiry
    );

    f2SubItems.push({
      sku: parsed.sku,
      qty: (action === 'REMOVE' ? '-' : '') + parsed.qty,
      uom: resolveSkuUom(parsed.sku, retryUomCache),
      success: result.success,
      status: result.success ? successStatus : 'Failed',
      error: result.success ? '' : errMsg,
      action: subAction,
      qtyColor: action === 'ADD' ? 'green' : 'red'
    });
  }

  // Update parent header statuses based on sub-item outcomes
  updateRetryHeaderStatuses(sheet, data, touchedHeaders);

  // Log retry batch to Activity as F2
  if (f2SubItems.length > 0) {
    var f2Status = failCount === 0 ? 'success' : successCount === 0 ? 'failed' : 'partial';
    var f2Detail = joinActivitySegments_(['Activity Retry', buildActivityCountSummary_(checkedRows.length, 'adjustment', 'adjustments', '')]);
    logActivity('F2', f2Detail, f2Status, 'Activity -> WASP', f2SubItems);
  }

  Logger.log('Retry complete: ' + successCount + ' success, ' + failCount + ' failed');
}

/**
 * Find the parent header row for a sub-item (search upward for WK-ID).
 * Returns { execId, flow, details, batch, expiry, outputSku }
 */
function findParentHeader(data, startIdx) {
  for (var i = startIdx - 1; i >= 3; i--) {
    var colA = String(data[i][0]).trim();
    if (colA.indexOf('WK-') === 0) {
      var details = String(data[i][3]).trim();
      var flow = String(data[i][2]).trim();

      // Parse batch from MO name: "MO-7226 (UFC410C 02/29) // FEB 24"
      var batch = '';
      var expiry = '';
      var bMatch = details.match(/\(([^)]+)\)/);
      if (bMatch) {
        var parts = bMatch[1].trim().split(/\s+/);
        batch = parts[0] || '';
        if (parts[1]) {
          var eParts = parts[1].split('/');
          if (eParts.length === 2) {
            var eMonth = eParts[0];
            var eYear = eParts[1].length === 2 ? '20' + eParts[1] : eParts[1];
            expiry = eYear + '-' + eMonth + '-01';
          }
        }
      }

      // Parse output SKU: "MO-7226 (UFC410C)// FEB 24  MS-IP-4 x481"
      var outputSku = '';
      // Match SKU pattern after batch/date info (before error count or arrow)
      var skuMatch = details.match(/(?:\/\/[^A-Z]*|^\S+\s+)([A-Z][A-Z0-9][\w-]+)\s+x[\d.]+/);
      if (skuMatch) {
        outputSku = skuMatch[1];
      }

      return {
        execId: colA,
        flow: flow,
        details: details,
        batch: batch,
        expiry: expiry,
        outputSku: outputSku
      };
    }
  }
  return { execId: '', flow: '', details: '', batch: '', expiry: '', outputSku: '' };
}

/**
 * Add Retry checkboxes to all existing Failed/Skipped rows in Activity Log.
 * Run once after deploying to add checkboxes to old rows.
 *
 * Usage: setupRetryCheckboxes()
 */
function setupRetryCheckboxes() {
  var ss = SpreadsheetApp.openById(CONFIG.DEBUG_SHEET_ID);
  var sheet = ss.getSheetByName('Activity');
  if (!sheet) { Logger.log('No Activity sheet'); return; }

  // Add header if column G is empty
  var headerVal = sheet.getRange(3, 7).getValue();
  if (!headerVal) {
    sheet.getRange(3, 7).setValue('Retry');
    sheet.setColumnWidth(7, 50);
  }

  var lastRow = sheet.getLastRow();
  if (lastRow < 4) return;

  var statuses = sheet.getRange(4, 5, lastRow - 3, 1).getValues();
  var added = 0;

  for (var i = 0; i < statuses.length; i++) {
    var st = String(statuses[i][0]).trim();
    if (st === 'Failed' || st === 'Skipped') {
      sheet.getRange(i + 4, 7).insertCheckboxes();
      added++;
    }
  }

  Logger.log('Added ' + added + ' Retry checkboxes to existing Failed/Skipped rows');
}

/**
 * Retry a specific failed MO. Legacy -- prefer retryMarkedItems().
 */
function retryFailedMO(searchTerm) {
  if (!searchTerm) {
    Logger.log('Usage: retryFailedMO("WK-1418") or retryFailedMO("MO-7226")');
    return;
  }

  var ss = SpreadsheetApp.openById(CONFIG.DEBUG_SHEET_ID);
  var actSheet = ss.getSheetByName('Activity');
  if (!actSheet) {
    Logger.log('Activity sheet not found');
    return;
  }

  var lastRow = actSheet.getLastRow();
  if (lastRow < 4) {
    Logger.log('Activity sheet is empty');
    return;
  }

  var data = actSheet.getRange(1, 1, lastRow, 6).getValues();

  // Find header row matching search term (WK-ID in col A or MO ref in col D)
  var headerIdx = -1;
  for (var i = data.length - 1; i >= 3; i--) {
    var colA = String(data[i][0]).trim();
    var colD = String(data[i][3]).trim();
    if (colA === searchTerm || colD.indexOf(searchTerm) >= 0) {
      if (colA.indexOf('WK-') === 0) {
        headerIdx = i;
        break;
      }
    }
  }

  if (headerIdx < 0) {
    Logger.log('No Activity entry found for: ' + searchTerm);
    return;
  }

  var headerRow = data[headerIdx];
  var execId = String(headerRow[0]).trim();
  var flow = String(headerRow[2]).trim();
  var details = String(headerRow[3]).trim();
  var status = String(headerRow[4]).trim();

  Logger.log('Found: ' + execId + ' | ' + flow + ' | ' + details + ' | ' + status);

  // Parse MO batch/expiry from header details
  // Pattern: "MO-7226 (UFC410C 02/29) // FEB 24  MS-IP-4 x481..."
  var moBatch = '';
  var moExpiry = '';
  var batchMatch = details.match(/\(([^)]+)\)/);
  if (batchMatch) {
    var bParts = batchMatch[1].trim().split(/\s+/);
    moBatch = bParts[0] || '';
    if (bParts[1]) {
      var eParts = bParts[1].split('/');
      if (eParts.length === 2) {
        var eMonth = eParts[0];
        var eYear = eParts[1].length === 2 ? '20' + eParts[1] : eParts[1];
        moExpiry = eYear + '-' + eMonth + '-01';
      }
    }
  }
  Logger.log('MO batch: ' + (moBatch || 'none') + ', expiry: ' + (moExpiry || 'none'));

  // Determine output SKU from header (pattern: "MO-XXXX ...  SKU xQTY")
  var outputSku = '';
  var outputMatch = details.match(/\s+([A-Z0-9][\w-]+)\s+x(\d+)/);
  if (outputMatch) {
    outputSku = outputMatch[1];
  }

  // Read sub-item rows (rows below header until next header or end)
  var subItems = [];
  for (var j = headerIdx + 1; j < data.length; j++) {
    var subId = String(data[j][0]).trim();
    // Stop at next header row (has WK-ID)
    if (subId.indexOf('WK-') === 0) break;

    var subDetail = String(data[j][3]).trim();
    var subStatus = String(data[j][4]).trim();
    var subError = String(data[j][5]).trim();

    // Only care about Failed or Skipped items
    if (subStatus !== 'Failed' && subStatus !== 'Skipped') continue;
    if (!subDetail) continue;

    // Parse sub-item: "    ├─ SKU xQTY  LOCATION  lot:XXX  exp:YYYY-MM-DD"
    var parsed = parseActivitySubItem(subDetail);
    if (!parsed || !parsed.sku) continue;

    parsed.status = subStatus;
    parsed.error = subError;
    subItems.push(parsed);
  }

  if (subItems.length === 0) {
    Logger.log('No failed/skipped items found for ' + execId);
    return;
  }

  Logger.log('Found ' + subItems.length + ' failed/skipped items to retry');

  // Generate adjustment rows
  var adjSheet = getAdjustmentsSheet();
  var created = 0;

  for (var k = 0; k < subItems.length; k++) {
    var item = subItems[k];
    var action = '';
    var lot = item.lot || '';
    var expiry = item.expiry || '';
    var notes = 'Retry ' + execId + ' (' + item.error + ')';

    if (item.sku === outputSku) {
      // Output item — needs ADD with MO batch
      action = 'ADD';
      if (!lot) lot = moBatch;
      if (!expiry) expiry = moExpiry;
      notes = 'Retry ' + execId + ' output ADD lot:' + lot;
    } else {
      // Ingredient — needs REMOVE (wasn't deducted)
      action = 'REMOVE';
      notes = 'Retry ' + execId + ' ingredient REMOVE (' + item.error + ')';
    }

    var location = item.location || 'PRODUCTION';
    addAdjustmentRow(action, item.sku, item.qty, CONFIG.WASP_SITE, location, lot, expiry, notes);
    created++;

    Logger.log('  ' + action + ' ' + item.sku + ' x' + item.qty + ' at ' + location + (lot ? ' lot=' + lot : '') + (expiry ? ' exp=' + expiry : ''));
  }

  Logger.log('Created ' + created + ' adjustment rows — review in Adjustments sheet, then run processAdjustments()');
}

/**
 * Parse an Activity Log sub-item line.
 * Input:  "    ├─ EO-LAV x300  PRODUCTION  lot:5005828385  exp:2029-07-01"
 * Output: { sku: 'EO-LAV', qty: 300, location: 'PRODUCTION', lot: '5005828385', expiry: '2029-07-01' }
 */
function parseActivitySubItem(line) {
  if (!line) return null;

  // Strip tree characters
  var clean = line.replace(/[├└─│\s]+/g, ' ').trim();
  if (!clean) return null;

  // Extract SKU (first word)
  var parts = clean.split(/\s+/);
  var sku = parts[0] || '';

  // Extract qty (pattern: xNNN)
  var qty = 0;
  var qtyMatch = clean.match(/x([\d.]+)/);
  if (qtyMatch) qty = parseFloat(qtyMatch[1]) || 0;

  // Extract location (PRODUCTION, PROD-RECEIVING, etc.)
  var location = '';
  var locMatch = clean.match(/\b(PRODUCTION|PROD-RECEIVING|RECEIVING-DOCK|SW-STORAGE|UNSORTED|SHOPIFY)\b/);
  if (locMatch) location = locMatch[1];

  // Extract lot (lot:XXX)
  var lot = '';
  var lotMatch = clean.match(/lot:(\S+)/);
  if (lotMatch) lot = lotMatch[1];

  // Extract expiry (exp:YYYY-MM-DD)
  var expiry = '';
  var expMatch = clean.match(/exp:(\S+)/);
  if (expMatch) expiry = expMatch[1];

  return {
    sku: sku,
    qty: qty,
    location: location,
    lot: lot,
    expiry: expiry
  };
}

/**
 * Determine WASP API action and success status based on flow type.
 */
function getRetryAction(flow, isOutput) {
  if (flow.indexOf('F1') >= 0) return { action: 'ADD', successStatus: 'Received', defaultLocation: 'RECEIVING-DOCK' };
  if (flow.indexOf('F3') >= 0) return { action: 'ADD', successStatus: 'Synced', defaultLocation: '' };
  if (flow.indexOf('F4') >= 0) return {
    action: isOutput ? 'ADD' : 'REMOVE',
    successStatus: isOutput ? 'Produced' : 'Consumed',
    defaultLocation: isOutput ? 'PROD-RECEIVING' : 'PRODUCTION'
  };
  if (flow.indexOf('F5') >= 0) return { action: 'REMOVE', successStatus: 'Deducted', defaultLocation: 'SHOPIFY' };
  if (flow.indexOf('F6') >= 0) return { action: 'ADD', successStatus: 'Staged', defaultLocation: 'PRODUCTION' };
  return { action: 'ADD', successStatus: 'Added', defaultLocation: 'PRODUCTION' };
}

/**
 * Update parent header statuses after retrying sub-items.
 */
function updateRetryHeaderStatuses(sheet, data, touchedHeaders) {
  for (var execId in touchedHeaders) {
    for (var i = 3; i < data.length; i++) {
      if (String(data[i][0]).trim() !== execId) continue;
      var headerRow = i + 1;
      var flow = String(data[i][2]).trim();
      var flowKey = flow.substring(0, 2);
      var allSuccess = true;
      var anySuccess = false;
      var hasSubItems = false;
      for (var j = i + 1; j < data.length; j++) {
        var subA = String(data[j][0]).trim();
        if (subA.indexOf('WK-') === 0) break;
        var subDetail = String(data[j][3]).trim();
        if (!subDetail) continue;
        hasSubItems = true;
        var curStatus = String(sheet.getRange(j + 1, 5).getValue()).trim();
        if (curStatus === 'Failed' || curStatus === 'Skipped') { allSuccess = false; }
        else if (curStatus) { anySuccess = true; }
      }
      if (!hasSubItems) break;
      if (allSuccess) {
        var flowStatus = FLOW_STATUS_TEXT[flowKey] || 'Complete';
        sheet.getRange(headerRow, 5).setValue(flowStatus);
        sheet.getRange(headerRow, 6).setValue('');
      } else if (anySuccess) {
        sheet.getRange(headerRow, 5).setValue('Partial');
      }
      break;
    }
  }
}

/**
 * Scan ALL partial/failed MO entries in Activity Log and list them.
 * Useful to see what needs fixing.
 *
 * Usage: listFailedMOs()
 */
function listFailedMOs() {
  var ss = SpreadsheetApp.openById(CONFIG.DEBUG_SHEET_ID);
  var actSheet = ss.getSheetByName('Activity');
  if (!actSheet) { Logger.log('No Activity sheet'); return; }

  var lastRow = actSheet.getLastRow();
  if (lastRow < 4) { Logger.log('Empty'); return; }

  var data = actSheet.getRange(1, 1, lastRow, 6).getValues();
  var found = [];

  for (var i = 3; i < data.length; i++) {
    var colA = String(data[i][0]).trim();
    var colC = String(data[i][2]).trim();
    var colE = String(data[i][4]).trim();

    if (colA.indexOf('WK-') === 0 && colC.indexOf('F4') >= 0) {
      if (colE === 'Partial' || colE === 'Failed') {
        var colD = String(data[i][3]).trim();
        var colF = String(data[i][5]).trim();
        found.push(colA + '  ' + colD.substring(0, 60) + '  [' + colE + ']  ' + colF);
      }
    }
  }

  if (found.length === 0) {
    Logger.log('No failed/partial F4 entries found');
  } else {
    Logger.log('Failed/Partial MOs (' + found.length + '):');
    for (var f = 0; f < found.length; f++) {
      Logger.log('  ' + found[f]);
    }
    Logger.log('\nTo retry: retryFailedMO("WK-XXXX")');
  }
}

// ============================================
// RETRY MENU
// ============================================

function onOpenRetryMenu() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Katana-WASP')
    .addItem('Retry Failed', 'retryMarkedItems')
    .addSeparator()
    .addItem('Setup Retry Checkboxes', 'setupRetryCheckboxes')
    .addItem('Process Adjustments', 'processAdjustments')
    .addItem('Run Activity Conformance Suite', 'runActivityConformanceSuite')
    .addItem('Run Real Activity Samples', 'runRealActivitySamples')
    .addItem('Debug MO Batch Data', 'debugMOBatchDataPrompt')
    .addItem('Repair MO From Debug', 'repairMOFromDebugPrompt')
    .addSeparator()
    .addItem('Build F5 Print Audit', 'buildF5PrintAuditToday')
    .addItem('Build F5 Manual Recovery', 'buildF5RecoveryReportToday')
    .addItem('Finalize F5 Manual Recovery', 'finalizeF5RecoveryReport')
    .addItem('Resume F5 Shipping Polling', 'resumeF5ShippingPolling')
    .addToUi();
}

/**
 * One-time setup: installs menu trigger + Activity edit trigger.
 * Run this once from the GAS editor after deploying.
 */
function installRetryMenu() {
  var ss = SpreadsheetApp.openById(CONFIG.DEBUG_SHEET_ID);
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'onOpenRetryMenu') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  ScriptApp.newTrigger('onOpenRetryMenu')
    .forSpreadsheet(ss)
    .onOpen()
    .create();
  Logger.log('Menu trigger installed');
}
