/**
 * 05_PushEngine.gs
 * Push pending changes from "Wasp" sheet to WASP API
 * Google Apps Script V8 compatible (var only, no ES6)
 */

/**
 * Main entry point for pushing pending changes to WASP.
 * Called from menu "Push to WASP".
 */
function pushToWasp() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var waspSheet = ss.getSheetByName('Wasp');

  if (!waspSheet) {
    SpreadsheetApp.getUi().alert('Wasp sheet not found.');
    return;
  }

  var startTime = new Date();
  var lastRow = waspSheet.getLastRow();

  if (lastRow < 2) {
    try {
      ss.toast('No data to push.', 'Push Complete', 5);
    } catch (e) {
      Logger.log('Toast failed: ' + e.message);
    }
    return;
  }

  var NUM_COLS = 18;

  // Read all data rows (A2:R{lastRow})
  var dataRange = waspSheet.getRange(2, 1, lastRow - 1, NUM_COLS);
  var data = dataRange.getValues();
  var backgrounds = dataRange.getBackgrounds();

  var pendingRows = [];
  var i;

  // Find rows with Status === 'Pending' (col P, index 15)
  var deleteRows = [];
  for (i = 0; i < data.length; i++) {
    var status = String(data[i][15]).trim();
    if (status === 'Pending') {
      pendingRows.push(i);
    } else if (status === 'DELETE') {
      deleteRows.push(i);
    }
  }

  // Detect deleted rows by comparing sheet against manifest
  var deletedActions = detectDeletedRows_(waspSheet, lastRow, NUM_COLS);

  Logger.log('Found ' + pendingRows.length + ' pending rows, ' + deletedActions.length + ' deleted rows.');

  var activeUser = '';
  try { activeUser = Session.getActiveUser().getEmail(); } catch (e) {}
  if (!activeUser) { try { activeUser = Session.getEffectiveUser().getEmail(); } catch (e) {} }
  if (!activeUser) { activeUser = PropertiesService.getScriptProperties().getProperty('USER_NAME') || ''; }

  var successCount = 0;
  var errorCount = 0;
  var errorDetails = [];
  var j;

  // Process each pending row
  for (j = 0; j < pendingRows.length; j++) {
    var rowIndex = pendingRows[j];
    var row = data[rowIndex];
    var sheetRowNum = rowIndex + 2; // Convert to 1-based sheet row

    Logger.log('Processing row ' + sheetRowNum + ': ' + row[0]);

    // Compute actions needed
    var actions = computeRowActions_(row, rowIndex);

    if (actions.length === 0) {
      Logger.log('Row ' + sheetRowNum + ': No actions needed.');
      data[rowIndex][14] = 'Synced ' + formatTimestamp();
      backgrounds[rowIndex][14] = '#ffffff';
      successCount++;
      continue;
    }

    // Execute each action
    var allSuccess = true;
    var errorMsg = '';
    var k;

    for (k = 0; k < actions.length; k++) {
      var action = actions[k];
      var actionDesc = action.type + (action.qty ? (' ' + action.qty) : '') + ' of ' + action.sku;
      Logger.log('Action ' + (k + 1) + '/' + actions.length + ': ' + actionDesc);

      var result = executeAction_(action);

      if (!result.success) {
        allSuccess = false;
        errorMsg = result.error;
        Logger.log('Action failed: ' + errorMsg);
        break;
      }
    }

    // Update row status
    if (allSuccess) {
      // Capture log values BEFORE modifying data[rowIndex] — row is same reference,
      // so mutations below (e.g. data[rowIndex][10] = ...) would corrupt log reads.
      var logSku = String(row[0]).trim();
      var logSite = String(row[3]).trim();
      var logLoc = String(row[4]).trim();
      var logLot = String(row[6]).trim();
      var logExpiry = (row[7] instanceof Date) ? Utilities.formatDate(row[7], 'America/Vancouver', 'yyyy-MM-dd') : String(row[7] || '').trim();
      var logOrigQty = (row[10] instanceof Date) ? 0 : (parseFloat(row[10]) || 0);
      var logNewQty = (row[11] instanceof Date) ? 0 : (parseFloat(row[11]) || 0);

      data[rowIndex][15] = 'Synced ' + formatTimestamp();

      // Clear yellow background from entire row
      for (var bg = 0; bg < NUM_COLS; bg++) {
        backgrounds[rowIndex][bg] = '#ffffff';
      }

      // Update Orig values to match new values
      data[rowIndex][10] = data[rowIndex][11];  // Orig WASP Qty = W.Qty
      data[rowIndex][12] = data[rowIndex][3];   // Orig Site = Site
      data[rowIndex][13] = data[rowIndex][4];   // Orig Location = Location
      data[rowIndex][16] = data[rowIndex][6];   // Orig Lot = Lot
      data[rowIndex][17] = data[rowIndex][7];   // Orig DateCode = DateCode

      successCount++;

      // Log to Adjustments Log tab
      var logDiff = Math.round((logNewQty - logOrigQty) * 100) / 100;
      if (Math.abs(logDiff) > 0.01) {
        var logAction = logDiff > 0 ? 'Add' : 'Remove';
        logAdjustment_(ss, 'Sync Sheet', logAction, activeUser, logSku, '', logSite, logLoc, logLot, logExpiry, logDiff, 'OK');
      }
    } else {
      data[rowIndex][15] = 'ERROR';

      // Set red background on entire row for errors
      for (var bg2 = 0; bg2 < NUM_COLS; bg2++) {
        backgrounds[rowIndex][bg2] = '#ffcdd2';
      }
      errorCount++;
      errorDetails.push('Row ' + sheetRowNum + ' (' + row[0] + '): ' + errorMsg);
    }

    // Rate limiting and UI refresh every 10 rows
    if ((j + 1) % 10 === 0) {
      Utilities.sleep(500);
      dataRange.setValues(data);
      dataRange.setBackgrounds(backgrounds);
      SpreadsheetApp.flush();
      Logger.log('Flushed after ' + (j + 1) + ' rows.');
    }
  }

  // Final write back for pending rows
  dataRange.setValues(data);
  dataRange.setBackgrounds(backgrounds);
  SpreadsheetApp.flush();

  // Process DELETE rows (user typed DELETE in Status column)
  // When any row for a SKU is marked DELETE, remove ALL rows for that SKU
  // across every location — not just the single marked row.
  var rowsToDelete = [];

  // Collect the set of SKUs marked for deletion
  var deleteSkus = {};
  for (var dri = 0; dri < deleteRows.length; dri++) {
    var dSku = String(data[deleteRows[dri]][0] || '').trim();
    if (dSku) deleteSkus[dSku] = true;
  }

  if (Object.keys(deleteSkus).length > 0) {
    var delDateStr = Utilities.formatDate(new Date(), 'America/Vancouver', 'yyyy-MM-dd');

    // For each DELETE SKU, find ALL rows for it in the sheet (any location/lot)
    for (var di2 = 0; di2 < data.length; di2++) {
      var dRow = data[di2];
      var dRowSku = String(dRow[0] || '').trim();
      if (!dRowSku || !deleteSkus[dRowSku]) continue;
      if (dRowSku.indexOf(' > ') > -1) continue; // skip separator rows

      var dSheetRowNum = di2 + 2;
      var dSite     = String(dRow[3] || '').trim();
      var dLocation = String(dRow[4] || '').trim();
      var dLot      = String(dRow[6] || '').trim();
      var dDateCode = '';
      if (dRow[7] instanceof Date) {
        var ddY = dRow[7].getFullYear(), ddM = dRow[7].getMonth() + 1, ddD = dRow[7].getDate();
        dDateCode = ddY + '-' + (ddM < 10 ? '0' + ddM : ddM) + '-' + (ddD < 10 ? '0' + ddD : ddD);
      } else {
        dDateCode = String(dRow[7] || '').trim();
        if (dDateCode.indexOf('T') > -1) dDateCode = dDateCode.split('T')[0];
      }
      // Use W.Qty (index 11) — the current sheet qty — not the Orig baseline (index 10)
      var dWQty = (dRow[11] instanceof Date) ? 0 : (parseFloat(dRow[11]) || 0);

      if (dWQty > 0 && dSite) {
        Logger.log('DELETE ' + dRowSku + ' row ' + dSheetRowNum + ': REMOVE ' + dWQty + ' from ' + dSite + '/' + dLocation);
        var dAction = {
          type: 'REMOVE',
          sku: dRowSku,
          qty: dWQty,
          site: dSite,
          location: dLocation,
          lot: dLot,
          dateCode: dDateCode,
          _rowLot: dLot,
          _rowDateCode: dDateCode,
          notes: 'Sheet DELETE ' + delDateStr
        };
        var dResult = executeAction_(dAction);
        if (dResult.success) {
          successCount++;
          rowsToDelete.push(dSheetRowNum);
          Logger.log('DELETE ' + dRowSku + ' row ' + dSheetRowNum + ' REMOVE success');
          logAdjustment_(ss, 'Sync Sheet', 'Remove', activeUser, dRowSku, '', dSite, dLocation, dLot, dDateCode, -dWQty, 'OK');
        } else {
          errorCount++;
          var dErrRaw = String(dResult.error || '');
          var dErrClean = dErrRaw;
          try {
            var dErrObj = JSON.parse(dErrRaw);
            var dMsgs = dErrObj.Messages || [];
            if (dMsgs.length > 0) dErrClean = dMsgs[0].Message + ' (' + dMsgs[0].ResultCode + ')';
          } catch (ep2) {}
          errorDetails.push(dRowSku + ' @ ' + dLocation + ': ' + dErrClean);
          Logger.log('DELETE ' + dRowSku + ' row ' + dSheetRowNum + ' FAILED: ' + dErrClean);
          logAdjustment_(ss, 'Sync Sheet', 'Remove', activeUser, dRowSku, '', dSite, dLocation, dLot, dDateCode, -dWQty, 'ERROR');
          data[di2][15] = 'ERROR';
          for (var dbg = 0; dbg < NUM_COLS; dbg++) {
            backgrounds[di2][dbg] = '#ffcdd2';
          }
        }
      } else {
        // W.Qty = 0 — no stock to remove, just mark row for sheet deletion
        successCount++;
        rowsToDelete.push(dSheetRowNum);
      }
    }

    // Note: WASP API does not expose a catalog delete endpoint (all ic/item/delete paths 404).
    // Stock removal + sheet row deletion is the complete DELETE behavior.
  }

  // Write back status for failed DELETEs
  if (deleteRows.length > 0) {
    dataRange.setValues(data);
    dataRange.setBackgrounds(backgrounds);
    SpreadsheetApp.flush();
  }

  // Delete successful DELETE rows from sheet (bottom-up to preserve row numbers)
  if (rowsToDelete.length > 0) {
    rowsToDelete.sort(function(a, b) { return b - a; }); // descending
    for (var rtd = 0; rtd < rowsToDelete.length; rtd++) {
      waspSheet.deleteRow(rowsToDelete[rtd]);
    }
    SpreadsheetApp.flush();
    Logger.log('Deleted ' + rowsToDelete.length + ' rows from Wasp sheet');
  }

  // Process deleted rows (detected from manifest comparison)
  var deleteSuccess = 0;
  var deleteError = 0;
  for (var di = 0; di < deletedActions.length; di++) {
    var delAction = deletedActions[di];
    Logger.log('Deleted row: REMOVE ' + delAction.qty + ' of ' + delAction.sku + ' from ' + delAction.site + '/' + delAction.location);

    var delResult = executeAction_(delAction);
    if (delResult.success) {
      deleteSuccess++;
      Logger.log('Deleted row REMOVE success: ' + delAction.sku);
      logAdjustment_(ss, 'Sync Sheet', 'Remove', activeUser, delAction.sku, '', delAction.site || '', delAction.location || '', delAction.lot || '', delAction.dateCode || '', -Math.abs(delAction.qty), 'OK');
    } else {
      deleteError++;
      errorDetails.push('Deleted row (' + delAction.sku + '): ' + delResult.error);
      Logger.log('Deleted row REMOVE failed: ' + delAction.sku + ' - ' + delResult.error);
      logAdjustment_(ss, 'Sync Sheet', 'Remove', activeUser, delAction.sku, '', delAction.site || '', delAction.location || '', delAction.lot || '', delAction.dateCode || '', -Math.abs(delAction.qty), 'ERROR');
    }

    if ((di + 1) % 10 === 0) Utilities.sleep(500);
  }
  successCount += deleteSuccess;
  errorCount += deleteError;

  // Update manifest after processing deletions
  if (deletedActions.length > 0) {
    updateManifestAfterDeletes_(deletedActions);
  }

  // ---- Also process Batch tab pending rows ----
  var batchResult = pushBatchPendingRows_(ss);
  successCount += batchResult.success;
  errorCount += batchResult.errors;
  if (batchResult.errorDetails.length > 0) {
    for (var bd = 0; bd < batchResult.errorDetails.length; bd++) {
      errorDetails.push(batchResult.errorDetails[bd]);
    }
  }

  // ---- Also process Katana tab Push rows ----
  var katanaResult = pushKatanaSyncRows_(ss);
  successCount += katanaResult.success;
  errorCount += katanaResult.errors;
  if (katanaResult.errorDetails.length > 0) {
    for (var kd = 0; kd < katanaResult.errorDetails.length; kd++) {
      errorDetails.push(katanaResult.errorDetails[kd]);
    }
  }
    // ---- Also process Katana tab qty adjustments ----
    var katanaAdjResult = pushKatanaAdjustments_(ss);
    successCount += katanaAdjResult.success;
    errorCount += katanaAdjResult.errors;
    if (katanaAdjResult.errorDetails.length > 0) {
        for (var kad = 0; kad < katanaAdjResult.errorDetails.length; kad++) {
            errorDetails.push(katanaAdjResult.errorDetails[kad]);
        }
    }

    var endTime = new Date();
  var duration = (endTime.getTime() - startTime.getTime()) / 1000;

  // Log to Sync History
  var summary = successCount + ' synced, ' + errorCount + ' errors';
  if (deleteSuccess > 0)          summary += ', ' + deleteSuccess + ' deleted';
  if (batchResult.success > 0)    summary += ', ' + batchResult.success + ' batch';
  if (katanaResult.success > 0)   summary += ', ' + katanaResult.success + ' katana';
  if (katanaAdjResult.success > 0) summary += ', ' + katanaAdjResult.success + ' katana adj';
  var errDetail = errorDetails.length > 0 ? errorDetails.join('; ') : '';

  var status = errorCount === 0 ? 'SUCCESS' : 'PARTIAL';
  logSyncHistory(ss, 'PUSH', summary, errorCount, errDetail, duration, status);

  // Show toast
  var message = 'Push complete: ' + successCount + ' succeeded, ' + errorCount + ' failed.';
  try {
    ss.toast(message, 'Push Complete', 5);
  } catch (e) {
    Logger.log('Toast failed: ' + e.message);
    Logger.log(message);
  }

  Logger.log('Push complete: ' + message + ' Duration: ' + duration + 's');
}

/**
 * Analyze a pending row and determine what WASP API actions are needed.
 * 16-col Wasp tab: A(0)SKU B(1)Name C(2)Site D(3)Location E(4)LotTracked
 *   F(5)Lot G(6)DateCode H(7)UOM I(8)OrigQty J(9)Qty K(10)OrigSite
 *   L(11)OrigLocation M(12)Match N(13)Status O(14)OrigLot P(15)OrigDateCode
 *
 * Two cases:
 *   Case 1: New row (no orig data) → CREATE + ADD
 *   Case 2: Any change → smart delta for qty-only, REMOVE old + ADD new for anything else
 *
 * @param {Array} row - 16-element array (0-indexed)
 * @param {number} rowIndex - 0-based index in data array (for logging)
 * @return {Array} Array of action objects
 */
function computeRowActions_(row, rowIndex) {
  var sku = String(row[0]).trim();
  var name = String(row[1]).trim();
  // row[2] = Type (Katana category, read-only)
  var site = String(row[3]).trim();
  var location = String(row[4]).trim();
  var lotTracked = String(row[5]).trim();
  var lot = String(row[6]).trim();
  var dateCode = '';
  var uom = String(row[8]).trim();
  // row[9] = K.Qty (read-only Katana qty, not used for push logic)
  var origQtyRaw = row[10];   // Orig WASP Qty (hidden baseline)
  var qtyRaw = row[11];       // W.Qty (editable)
  var origSite = String(row[12]).trim();
  var origLocation = String(row[13]).trim();
  // row[14] = Match, row[15] = Status
  var origLot = String(row[16]).trim();
  var origDateCode = '';

  // Handle Date objects in dateCode fields (Sheets auto-converts date-like text)
  if (row[7] instanceof Date) {
    var dcY = row[7].getFullYear();
    var dcM = row[7].getMonth() + 1;
    var dcD = row[7].getDate();
    dateCode = dcY + '-' + (dcM < 10 ? '0' + dcM : dcM) + '-' + (dcD < 10 ? '0' + dcD : dcD);
  } else {
    dateCode = String(row[7] || '').trim();
    if (dateCode.indexOf('T') > -1) dateCode = dateCode.split('T')[0];
  }

  if (row[17] instanceof Date) {
    var odcY = row[17].getFullYear();
    var odcM = row[17].getMonth() + 1;
    var odcD = row[17].getDate();
    origDateCode = odcY + '-' + (odcM < 10 ? '0' + odcM : odcM) + '-' + (odcD < 10 ? '0' + odcD : odcD);
  } else {
    origDateCode = String(row[17] || '').trim();
    if (origDateCode.indexOf('T') > -1) origDateCode = origDateCode.split('T')[0];
  }

  // Defend against Date objects in qty fields
  var origQty = (origQtyRaw instanceof Date) ? 0 : (parseFloat(origQtyRaw) || 0);
  var qty = (qtyRaw instanceof Date) ? 0 : (parseFloat(qtyRaw) || 0);

  var isLotTracked = (lotTracked === 'Yes' || lotTracked === 'TRUE' || lotTracked === 'true');
  var actions = [];
  var dateStr = Utilities.formatDate(new Date(), 'America/Vancouver', 'yyyy-MM-dd');

  // ---- Case 1: New row (no orig data) → CREATE + ADD ----
  if ((origQty === 0 || origQty === '') && origSite === '' && origLocation === '') {
    if (qty > 0) {
      actions.push({
        type: 'CREATE',
        sku: sku,
        name: name,
        site: site,
        location: location,
        lotTracking: isLotTracked,
        qty: qty,
        lot: lot,
        dateCode: dateCode,
        _rowLot: lot,
        _rowDateCode: dateCode,
        notes: 'Sheet push new item ' + dateStr
      });
      Logger.log('Row ' + (rowIndex + 2) + ': NEW row - CREATE + ADD ' + qty);
    }
    return actions;
  }

  // ---- Case 2: Existing row — detect what changed ----
  var qtyChanged = Math.abs(qty - origQty) > 0.01;
  var siteChanged = site !== origSite;
  var locationChanged = location !== origLocation;
  var lotChanged = lot !== origLot;
  var dateCodeChanged = dateCode !== origDateCode;

  var anythingBesidesQty = siteChanged || locationChanged || lotChanged || dateCodeChanged;

  // Sub-case 2a: ONLY qty changed — smart delta (ADD or REMOVE the difference)
  if (qtyChanged && !anythingBesidesQty) {
    var delta = qty - origQty;
    if (delta > 0) {
      actions.push({
        type: 'ADD',
        sku: sku,
        qty: delta,
        site: site,
        location: location,
        lot: lot,
        dateCode: dateCode,
        _rowLot: lot,
        _rowDateCode: dateCode,
        notes: 'Sheet push qty increase ' + dateStr
      });
      Logger.log('Row ' + (rowIndex + 2) + ': QTY increase - ADD ' + delta);
    } else {
      actions.push({
        type: 'REMOVE',
        sku: sku,
        qty: Math.abs(delta),
        site: site,
        location: location,
        lot: lot,
        dateCode: dateCode,
        _rowLot: lot,
        _rowDateCode: dateCode,
        notes: 'Sheet push qty decrease ' + dateStr
      });
      Logger.log('Row ' + (rowIndex + 2) + ': QTY decrease - REMOVE ' + Math.abs(delta));
    }
    return actions;
  }

  // Sub-case 2b: Site/Location/Lot/DateCode changed (with or without qty change)
  // → REMOVE all from old location/lot, ADD all to new location/lot
  if (anythingBesidesQty) {
    // REMOVE from old location with old lot
    if (origQty > 0) {
      actions.push({
        type: 'REMOVE',
        sku: sku,
        qty: origQty,
        site: origSite || site,
        location: origLocation || location,
        lot: origLot,
        dateCode: origDateCode,
        _rowLot: origLot,
        _rowDateCode: origDateCode,
        notes: 'Sheet push change from ' + (origSite || site) + '/' + (origLocation || location) +
               (lotChanged ? ' lot=' + origLot : '') + (dateCodeChanged ? ' dc=' + origDateCode : '')
      });
    }

    // ADD to new location with new lot
    if (qty > 0) {
      actions.push({
        type: 'ADD',
        sku: sku,
        qty: qty,
        site: site,
        location: location,
        lot: lot,
        dateCode: dateCode,
        _rowLot: lot,
        _rowDateCode: dateCode,
        notes: 'Sheet push change to ' + site + '/' + location +
               (lotChanged ? ' lot=' + lot : '') + (dateCodeChanged ? ' dc=' + dateCode : '')
      });
    }

    var changeDesc = [];
    if (siteChanged) changeDesc.push('site');
    if (locationChanged) changeDesc.push('location');
    if (lotChanged) changeDesc.push('lot');
    if (dateCodeChanged) changeDesc.push('dateCode');
    if (qtyChanged) changeDesc.push('qty');
    Logger.log('Row ' + (rowIndex + 2) + ': CHANGE (' + changeDesc.join('+') + ') - REMOVE old + ADD new');
    return actions;
  }

  // No changes detected (shouldn't happen for Pending rows, but safe fallback)
  return actions;
}

/**
 * Execute a single WASP API action.
 * @param {Object} action - Action object with type, sku, qty, site, location, lot, dateCode, notes
 * @return {Object} { success: boolean, error: string }
 */
function executeAction_(action) {
  // Note: enginPreMark_ is called inside waspAddInventory / waspRemoveInventory (and their lot variants),
  // so there is no need to call it here. Calling it twice adds ~7s of HTTP overhead per action.

  var notes = action.notes || ('Sheet push ' + Utilities.formatDate(new Date(), 'America/Vancouver', 'yyyy-MM-dd'));
  var site = action.site || WASP_DEFAULT_SITE;
  var defaultCreateLocation = site === 'MMH Mayfair' ? 'QA-Hold-1' : 'UNSORTED';
  var result;

  if (action.type === 'ADD') {
    if (action.lot) {
      result = waspAddInventoryWithLot(action.sku, action.qty, action.location, action.lot, action.dateCode || '', notes, site);
    } else {
      result = waspAddInventory(action.sku, action.qty, action.location, notes, site);
    }
  } else if (action.type === 'REMOVE') {
    if (action.lot) {
      result = waspRemoveInventoryWithLot(action.sku, action.qty, action.location, action.lot, notes, site, action.dateCode || '');
    } else {
      result = waspRemoveInventory(action.sku, action.qty, action.location, notes, site);
    }
  } else if (action.type === 'UPDATE_LOT_TRACKING') {
    result = waspUpdateItemLotTracking(action.sku, action.lotTracking);
  } else if (action.type === 'SKIP') {
    // Zero-qty manifest entry — nothing to remove from WASP, just clean the manifest
    result = { success: true, response: 'skipped (zero qty)' };
  } else if (action.type === 'CREATE') {
    // Create new item in WASP catalog, then add stock
    var createResult = waspCreateItem(action.sku, action.name || action.sku, action.site || WASP_DEFAULT_SITE, action.location || defaultCreateLocation, action.lotTracking || false, action.uom || '', action.salesPrice || 0);
    if (!createResult.success) {
      // Item likely already exists — always proceed to ADD
      // If item truly doesn't exist, the ADD will fail with a clear error
      var createErr = String(createResult.response || '').substring(0, 200);
      Logger.log('CREATE for ' + action.sku + ' failed: ' + createErr + ' — proceeding to ADD (item may already exist)');
    }
    // Now add stock — delegate to executeAction_ so -57041 retry works
    if (action.qty > 0) {
      var addAction = {
        type: 'ADD',
        sku: action.sku,
        qty: action.qty,
        site: action.site,
        location: action.location,
        lot: action.lot || '',
        dateCode: action.dateCode || '',
        _rowLot: action._rowLot || '',
        _rowDateCode: action._rowDateCode || '',
        notes: notes
      };
      return executeAction_(addAction);
    } else {
      return { success: true, error: '' };
    }
  } else {
    return { success: false, error: 'Unknown action type: ' + action.type };
  }

  if (result.success) {
    return { success: true, error: '' };
  }

  // Check for -46002 (insufficient qty) on REMOVE — retry with actual WASP qty
  var errorMsg = String(result.response || 'Unknown error');
  if (action.type === 'REMOVE' && errorMsg.indexOf('-46002') > -1 && !action._retried) {
    Logger.log('REMOVE -46002 for ' + action.sku + ' — fetching real WASP qty to retry');
    try {
      var searchResult = waspFetch('ic/item/advancedinventorysearch', {
        PageSize: 100,
        PageNumber: 1,
        ItemNumber: action.sku,
        SiteName: action.site
      });
      var realQty = 0;
      var items = searchResult.Data || searchResult.data || [];
      if (searchResult.Data && searchResult.Data.ResultList) items = searchResult.Data.ResultList;
      for (var ri = 0; ri < items.length; ri++) {
        var rItem = items[ri];
        if (rItem.ItemNumber === action.sku && rItem.LocationCode === action.location) {
          if (action.lot) {
            if (rItem.Lot === action.lot) { realQty = rItem.QtyAvailable || 0; break; }
          } else {
            realQty += (rItem.QtyAvailable || 0);
          }
        }
      }
      if (realQty > 0) {
        Logger.log('Real WASP qty for ' + action.sku + ' at ' + action.location + ': ' + realQty + ' (was trying to remove ' + action.qty + ')');
        var retryAction = {
          type: 'REMOVE',
          sku: action.sku,
          qty: realQty,
          site: action.site,
          location: action.location,
          lot: action.lot || '',
          dateCode: action.dateCode || '',
          notes: action.notes,
          _retried: true
        };
        return executeAction_(retryAction);
      }
      // realQty === 0 at assumed location — scan all locations for any available stock
      var foundLoc = '';
      var foundSite = action.site;
      var foundLocQty = 0;
      for (var ri2 = 0; ri2 < items.length; ri2++) {
        var rItem2 = items[ri2];
        if (String(rItem2.ItemNumber || '').trim() === action.sku) {
          var availQty = parseFloat(rItem2.QtyAvailable || 0);
          if (availQty > 0) {
            foundLoc = String(rItem2.LocationCode || '').trim();
            foundSite = String(rItem2.SiteName || action.site).trim();
            foundLocQty = availQty;
            break;
          }
        }
      }
      if (foundLoc && foundLocQty > 0) {
        Logger.log('REMOVE -46002: ' + action.sku + ' not at ' + action.location + ', found stock at ' + foundSite + '/' + foundLoc + ' qty=' + foundLocQty);
        var retryAction2 = {
          type: 'REMOVE',
          sku: action.sku,
          qty: Math.min(action.qty, foundLocQty),
          site: foundSite,
          location: foundLoc,
          lot: action.lot || '',
          dateCode: action.dateCode || '',
          notes: action.notes,
          _retried: true
        };
        return executeAction_(retryAction2);
      }
    } catch (retryErr) {
      Logger.log('Retry fetch failed: ' + retryErr.message);
    }
  }

  // Check for -57041 (lot/dateCode missing) — item is lot-tracked in WASP but action had no lot info
  if ((action.type === 'ADD' || action.type === 'REMOVE') && errorMsg.indexOf('-57041') > -1 && !action._lotRetried) {
    Logger.log(action.type + ' -57041 for ' + action.sku + ' — lot/dateCode missing, retrying with lot info');

    if (action.type === 'ADD') {
      // ADD: use row's lot/dateCode first, fallback to NO-LOT + future date
      var retryLot = (action._rowLot && action._rowLot !== '') ? action._rowLot : 'NO-LOT';
      var retryDC = '';
      if (action._rowDateCode && action._rowDateCode !== '') {
        retryDC = action._rowDateCode;
      } else {
        var futureDate = new Date();
        futureDate.setFullYear(futureDate.getFullYear() + 2);
        retryDC = futureDate.toISOString().slice(0, 10);
      }
      var addRetry = {
        type: 'ADD',
        sku: action.sku,
        qty: action.qty,
        site: action.site,
        location: action.location,
        lot: retryLot,
        dateCode: retryDC,
        notes: action.notes,
        _rowLot: action._rowLot,
        _rowDateCode: action._rowDateCode,
        _lotRetried: true
      };
      Logger.log('ADD retry with lot=' + retryLot + ' dateCode=' + retryDC);
      return executeAction_(addRetry);
    }

    if (action.type === 'REMOVE') {
      // REMOVE: look up actual lot from WASP inventory
      try {
        var lotSearch = waspFetch('ic/item/advancedinventorysearch', {
          PageSize: 100, PageNumber: 1,
          ItemNumber: action.sku, SiteName: action.site
        });
        var lotItems = lotSearch.Data || [];
        for (var li = 0; li < lotItems.length; li++) {
          var lItem = lotItems[li];
          if (lItem.ItemNumber === action.sku && lItem.LocationCode === action.location) {
            var foundLot = lItem.Lot || lItem.LotNumber || '';
            var foundDC = lItem.DateCode || '';
            if (foundLot) {
              var removeRetry = {
                type: 'REMOVE',
                sku: action.sku,
                qty: action.qty,
                site: action.site,
                location: action.location,
                lot: foundLot,
                dateCode: foundDC,
                notes: action.notes,
                _lotRetried: true
              };
              Logger.log('REMOVE retry with lot=' + foundLot + ' dateCode=' + foundDC);
              return executeAction_(removeRetry);
            }
          }
        }
        Logger.log('No lot found in WASP for ' + action.sku + ' at ' + action.location);
      } catch (lotErr) {
        Logger.log('Lot lookup failed: ' + lotErr.message);
      }
    }
  }

  if (errorMsg.length > 200) {
    errorMsg = errorMsg.substring(0, 200) + '...';
  }
  return { success: false, error: errorMsg };
}

// ============================================
// DELETED ROW DETECTION (manifest comparison)
// ============================================

/**
 * Compares current Wasp sheet rows against saved manifest.
 * Returns REMOVE actions for rows that were deleted from the sheet.
 *
 * @param {Sheet} sheet - The Wasp sheet
 * @param {number} lastRow - Last row in sheet
 * @param {number} numCols - Number of columns
 * @returns {Array} Array of REMOVE action objects for deleted rows
 */
function detectDeletedRows_(sheet, lastRow, numCols) {
  var manifest = loadWaspManifest_();
  var manifestKeys = Object.keys(manifest);
  if (manifestKeys.length === 0) return [];

  // Build current key set from sheet
  var currentKeys = {};
  if (lastRow >= 2) {
    var data = sheet.getRange(2, 1, lastRow - 1, numCols).getValues();
    for (var i = 0; i < data.length; i++) {
      var sku = String(data[i][0]).trim();
      if (!sku) continue; // separator row
      var site = String(data[i][3]).trim();
      var loc = String(data[i][4]).trim();
      var lot = String(data[i][6]).trim();
      var key = sku + '|' + site + '|' + loc + '|' + lot;
      currentKeys[key] = true;
    }
  }

  // Find keys in manifest but not in sheet
  var deleted = [];
  for (var m = 0; m < manifestKeys.length; m++) {
    var mKey = manifestKeys[m];
    if (currentKeys[mKey]) continue; // still in sheet

    var entry = manifest[mKey];
    var parts = mKey.split('|');

    // If entry has qty > 0: generate REMOVE action
    // If entry has qty === 0: nothing to remove from WASP — skip API call, just clean manifest
    var actionType = (entry.q && entry.q > 0) ? 'REMOVE' : 'SKIP';

    // Always include lot/dateCode from manifest — don't gate on lt flag
    // (item may have become lot-tracked in WASP since manifest was saved)
    deleted.push({
      type: actionType,
      sku: parts[0],
      qty: entry.q || 0,
      site: parts[1],
      location: parts[2],
      lot: parts[3] || '',
      dateCode: entry.dc || '',
      _rowLot: parts[3] || '',
      _rowDateCode: entry.dc || '',
      notes: 'Sheet push row deleted ' + Utilities.formatDate(new Date(), 'America/Vancouver', 'yyyy-MM-dd')
    });
  }

  if (deleted.length > 0) {
    Logger.log('Detected ' + deleted.length + ' deleted rows from manifest comparison');
  }

  return deleted;
}

/**
 * Removes successfully-deleted entries from the manifest.
 * Called after deleted row REMOVE actions succeed.
 *
 * @param {Array} deletedActions - Array of action objects that were processed
 */
function updateManifestAfterDeletes_(deletedActions) {
  var manifest = loadWaspManifest_();
  var changed = false;

  for (var i = 0; i < deletedActions.length; i++) {
    var action = deletedActions[i];
    var key = action.sku + '|' + action.site + '|' + action.location + '|' + (action.lot || '');
    if (manifest[key]) {
      delete manifest[key];
      changed = true;
    }
  }

  if (changed) {
    // Re-save manifest
    var json = JSON.stringify(manifest);
    var props = PropertiesService.getScriptProperties();

    var oldCount = parseInt(props.getProperty('WASP_MANIFEST_COUNT') || '0', 10);
    for (var ci = 0; ci < oldCount; ci++) {
      props.deleteProperty('WASP_MANIFEST_' + ci);
    }

    var CHUNK_SIZE = 8000;
    var chunks = Math.ceil(json.length / CHUNK_SIZE);
    for (var c = 0; c < chunks; c++) {
      props.setProperty('WASP_MANIFEST_' + c, json.substring(c * CHUNK_SIZE, (c + 1) * CHUNK_SIZE));
    }
    props.setProperty('WASP_MANIFEST_COUNT', String(chunks));
    Logger.log('Manifest updated: removed ' + deletedActions.length + ' deleted entries');
  }
}

// ============================================
// BATCH TAB PUSH (lot-tracked items)
// ============================================

/**
 * Push pending changes from Batch sheet to WASP API.
 * Batch rows are always lot-tracked.
 *
 * Batch tab layout (18 cols):
 *   A(1)=SKU, B(2)=K.Location, C(3)=K.Qty, D(4)=K.UOM, E(5)=K.Batch, F(6)=K.Expiry
 *   G(7)=W.Site, H(8)=W.Location, I(9)=W.Qty, J(10)=W.UOM, K(11)=W.Lot, L(12)=W.DateCode
 *   M(13)=Diff, N(14)=Status, O(15)=Orig W.Qty, P(16)=Orig W.Site, Q(17)=Orig W.Location
 *   R(18)=Sync Status
 *
 * @param {Spreadsheet} ss - Active spreadsheet
 * @returns {Object} { success: number, errors: number, errorDetails: Array }
 */
function pushBatchPendingRows_(ss) {
  var result = { success: 0, errors: 0, errorDetails: [] };
  var sheet = ss.getSheetByName('Batch');
  if (!sheet) return result;

  var activeUser = '';
  try { activeUser = Session.getActiveUser().getEmail(); } catch (e) {}
  if (!activeUser) { try { activeUser = Session.getEffectiveUser().getEmail(); } catch (e) {} }
  if (!activeUser) { activeUser = PropertiesService.getScriptProperties().getProperty('USER_NAME') || ''; }

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return result;

  var BATCH_COLS = 20;
  var dataRange = sheet.getRange(2, 1, lastRow - 1, BATCH_COLS);
  var data = dataRange.getValues();
  var backgrounds = dataRange.getBackgrounds();

  var pendingRows = [];
  var deleteRows = [];
  for (var i = 0; i < data.length; i++) {
    var syncStatus = String(data[i][17]).trim(); // col R(18), idx 17
    if (syncStatus === 'Pending') {
      pendingRows.push(i);
    } else if (syncStatus === 'DELETE') {
      deleteRows.push(i);
    }
  }

  if (pendingRows.length === 0 && deleteRows.length === 0) return result;

  Logger.log('Batch push: found ' + pendingRows.length + ' pending rows');

  for (var j = 0; j < pendingRows.length; j++) {
    var rowIndex = pendingRows[j];
    var row = data[rowIndex];
    var sheetRowNum = rowIndex + 2;

    // Skip separator rows
    var sku = String(row[0]).trim();
    if (!sku || sku.indexOf(' > ') > -1) continue;

    Logger.log('Batch push row ' + sheetRowNum + ': ' + sku);

    var actions = computeBatchRowActions_(row, rowIndex);

    if (actions.length === 0) {
      Logger.log('Batch row ' + sheetRowNum + ': No actions needed.');
      data[rowIndex][17] = 'Synced ' + formatTimestamp();
      for (var bg = 0; bg < BATCH_COLS; bg++) {
        backgrounds[rowIndex][bg] = '#ffffff';
      }
      result.success++;
      continue;
    }

    // Execute each action
    var allSuccess = true;
    var errorMsg = '';

    for (var k = 0; k < actions.length; k++) {
      var action = actions[k];
      var actionDesc = action.type + ' ' + action.qty + ' of ' + action.sku + ' lot=' + action.lot;
      Logger.log('Batch action ' + (k + 1) + '/' + actions.length + ': ' + actionDesc);

      var actionResult = executeAction_(action);

      if (!actionResult.success) {
        allSuccess = false;
        errorMsg = actionResult.error;
        Logger.log('Batch action failed: ' + errorMsg);
        break;
      }
    }

    // Update row status
    if (allSuccess) {
      data[rowIndex][17] = 'Synced ' + formatTimestamp();

      // Clear yellow background
      for (var bg2 = 0; bg2 < BATCH_COLS; bg2++) {
        backgrounds[rowIndex][bg2] = '#ffffff';
      }

      // Update orig values to match new values
      data[rowIndex][14] = data[rowIndex][8];  // Orig W.Qty = W.Qty
      data[rowIndex][15] = data[rowIndex][6];  // Orig W.Site = W.Site
      data[rowIndex][16] = data[rowIndex][7];  // Orig W.Location = W.Location
      data[rowIndex][18] = data[rowIndex][10]; // Orig W.Lot = W.Lot
      data[rowIndex][19] = data[rowIndex][11]; // Orig W.DateCode = W.DateCode            result.success++;
            // Log to Adjustments tab
            var bLogSku = String(row[0]).trim();
            var bLogSite = String(row[6]).trim();
            var bLogLoc = String(row[7]).trim();
            var bLogLot = String(row[10]).trim();
            var bLogOrigQty = (row[14] instanceof Date) ? 0 : (parseFloat(row[14]) || 0);
            var bLogNewQty = (row[8] instanceof Date) ? 0 : (parseFloat(row[8]) || 0);
            if (Math.abs(bLogNewQty - bLogOrigQty) > 0.01) {
                logAdjustment_(ss, 'Google Sheet', (bLogNewQty > bLogOrigQty ? 'Add' : 'Remove'), activeUser, bLogSku, '', bLogSite, bLogLoc, bLogLot, '', Math.round((bLogNewQty - bLogOrigQty) * 100) / 100, 'OK');
            }
        } else {
      data[rowIndex][17] = 'ERROR';

      // Red background for errors
      for (var bg3 = 0; bg3 < BATCH_COLS; bg3++) {
        backgrounds[rowIndex][bg3] = '#ffcdd2';
      }

      result.errors++;
      result.errorDetails.push('Batch row ' + sheetRowNum + ' (' + sku + '): ' + errorMsg);
    }

    // Rate limiting every 10 rows
    if ((j + 1) % 10 === 0) {
      Utilities.sleep(500);
      dataRange.setValues(data);
      dataRange.setBackgrounds(backgrounds);
      SpreadsheetApp.flush();
    }
  }

  // Final write back
  dataRange.setValues(data);
  dataRange.setBackgrounds(backgrounds);
  SpreadsheetApp.flush();

  // Process DELETE rows (user typed DELETE in Sync Status column)
  var batchRowsToDelete = [];
  for (var dri = 0; dri < deleteRows.length; dri++) {
    var delRowIndex = deleteRows[dri];
    var delRow = data[delRowIndex];
    var delSheetRowNum = delRowIndex + 2;
    var delSku = String(delRow[0]).trim();
    if (!delSku || delSku.indexOf(' > ') > -1) continue;

    var delSite = String(delRow[6]).trim();       // W.Site
    var delLocation = String(delRow[7]).trim();    // W.Location
    var delQty = (delRow[8] instanceof Date) ? 0 : (parseFloat(delRow[8]) || 0);  // W.Qty
    var delLot = String(delRow[10]).trim();        // W.Lot
    var delDateCode = '';
    if (delRow[11] instanceof Date) {
      var ddcY = delRow[11].getFullYear();
      var ddcM = delRow[11].getMonth() + 1;
      var ddcD = delRow[11].getDate();
      delDateCode = ddcY + '-' + (ddcM < 10 ? '0' + ddcM : ddcM) + '-' + (ddcD < 10 ? '0' + ddcD : ddcD);
    } else {
      delDateCode = String(delRow[11] || '').trim();
      if (delDateCode.indexOf('T') > -1) delDateCode = delDateCode.split('T')[0];
    }

    if (delQty > 0 && delSku && delSite && delSite !== '(unknown)') {
      Logger.log('Batch DELETE row ' + delSheetRowNum + ': REMOVE ' + delQty + ' of ' + delSku + ' from ' + delSite + '/' + delLocation);
      var delAction = {
        type: 'REMOVE',
        sku: delSku,
        qty: delQty,
        site: delSite,
        location: delLocation,
        lot: delLot,
        dateCode: delDateCode,
        _rowLot: delLot,
        _rowDateCode: delDateCode,
        notes: 'Batch DELETE ' + Utilities.formatDate(new Date(), 'America/Vancouver', 'yyyy-MM-dd')
      };
      var delResult = executeAction_(delAction);
      if (delResult.success) {
        result.success++;
        batchRowsToDelete.push(delSheetRowNum);
        Logger.log('Batch DELETE row ' + delSheetRowNum + ' success');
        logAdjustment_(ss, 'Google Sheet', 'Remove', activeUser, delSku, '', delSite, delLocation, delLot, delDateCode, -delQty, 'OK');
      } else {
        result.errors++;
        result.errorDetails.push('Batch row ' + delSheetRowNum + ' (' + delSku + '): DELETE failed - ' + delResult.error);
        Logger.log('Batch DELETE row ' + delSheetRowNum + ' failed: ' + delResult.error);
        data[delRowIndex][17] = 'ERROR';
        for (var dbg = 0; dbg < BATCH_COLS; dbg++) {
          backgrounds[delRowIndex][dbg] = '#ffcdd2';
        }
      }
    } else {
      // No stock to remove (qty=0 or unknown site), just delete the row
      batchRowsToDelete.push(delSheetRowNum);
      result.success++;
    }

    if ((dri + 1) % 10 === 0) Utilities.sleep(500);
  }

  // Write back status for failed DELETEs
  if (deleteRows.length > 0) {
    dataRange.setValues(data);
    dataRange.setBackgrounds(backgrounds);
    SpreadsheetApp.flush();
  }

  // Delete successful DELETE rows from sheet (bottom-up to preserve row numbers)
  if (batchRowsToDelete.length > 0) {
    batchRowsToDelete.sort(function(a, b) { return b - a; });
    for (var rtd = 0; rtd < batchRowsToDelete.length; rtd++) {
      sheet.deleteRow(batchRowsToDelete[rtd]);
    }
    SpreadsheetApp.flush();
  }

  Logger.log('Batch push complete: ' + result.success + ' success, ' + result.errors + ' errors');
  return result;
}

/**
 * Compute WASP API actions needed for a pending Batch row.
 * All Batch rows are lot-tracked, so lot+dateCode always included.
 *
 * Index map (20 cols):
 *   0:SKU 6:W.Site 7:W.Location 8:W.Qty 10:W.Lot 11:W.DateCode
 *   14:Orig W.Qty 15:Orig W.Site 16:Orig W.Location 17:Sync Status
 *   18:Orig W.Lot 19:Orig W.DateCode
 *
 * @param {Array} row - 20-element array (0-indexed)
 * @param {number} rowIndex - 0-based index (for logging)
 * @returns {Array} Array of action objects
 */
function computeBatchRowActions_(row, rowIndex) {
  var sku = String(row[0]).trim();
  var wSite = String(row[6]).trim();
  var wLocation = String(row[7]).trim();
  var wQtyRaw = row[8];
  var wLot = String(row[10]).trim();
  var wDateCode = '';
  var origQtyRaw = row[14];
  var origSite = String(row[15]).trim();
  var origLocation = String(row[16]).trim();
  var origLot = String(row[18] || '').trim();
  var origDateCode = '';

  // Handle Date objects in dateCode fields
  if (row[11] instanceof Date) {
    var dcY = row[11].getFullYear();
    var dcM = row[11].getMonth() + 1;
    var dcD = row[11].getDate();
    wDateCode = dcY + '-' + (dcM < 10 ? '0' + dcM : dcM) + '-' + (dcD < 10 ? '0' + dcD : dcD);
  } else {
    wDateCode = String(row[11] || '').trim();
    if (wDateCode.indexOf('T') > -1) wDateCode = wDateCode.split('T')[0];
  }

  if (row[19] instanceof Date) {
    var odcY = row[19].getFullYear();
    var odcM = row[19].getMonth() + 1;
    var odcD = row[19].getDate();
    origDateCode = odcY + '-' + (odcM < 10 ? '0' + odcM : odcM) + '-' + (odcD < 10 ? '0' + odcD : odcD);
  } else {
    origDateCode = String(row[19] || '').trim();
    if (origDateCode.indexOf('T') > -1) origDateCode = origDateCode.split('T')[0];
  }

  // Defend against Date objects in qty fields
  var wQty = (wQtyRaw instanceof Date) ? 0 : (parseFloat(wQtyRaw) || 0);
  var origQty = (origQtyRaw instanceof Date) ? 0 : (parseFloat(origQtyRaw) || 0);

  var qtyChanged = Math.abs(wQty - origQty) > 0.01;
  var siteChanged = wSite !== origSite && origSite !== '';
  var locationChanged = wLocation !== origLocation && origLocation !== '';
  var lotChanged = wLot !== origLot && origLot !== '';
  var dateCodeChanged = wDateCode !== origDateCode && origDateCode !== '';

  var anythingBesidesQty = siteChanged || locationChanged || lotChanged || dateCodeChanged;

  var actions = [];
  var dateStr = Utilities.formatDate(new Date(), 'America/Vancouver', 'yyyy-MM-dd');

  // Case 1: New row (no orig data — user filled in WASP columns on a KATANA ONLY row)
  if (origQty === 0 && origSite === '' && origLocation === '') {
    if (wQty > 0 && wSite) {
      actions.push({
        type: 'ADD',
        sku: sku,
        qty: wQty,
        site: wSite,
        location: wLocation,
        lot: wLot,
        dateCode: wDateCode,
        _rowLot: wLot,
        _rowDateCode: wDateCode,
        notes: 'Batch push new ' + dateStr
      });
      Logger.log('Batch row ' + (rowIndex + 2) + ': NEW - ADD ' + wQty + ' lot=' + wLot);
    }
    return actions;
  }

  // Case 2a: Only Qty changed — smart delta
  if (qtyChanged && !anythingBesidesQty) {
    var delta = wQty - origQty;
    if (delta > 0) {
      actions.push({
        type: 'ADD',
        sku: sku,
        qty: delta,
        site: wSite,
        location: wLocation,
        lot: wLot,
        dateCode: wDateCode,
        _rowLot: wLot,
        _rowDateCode: wDateCode,
        notes: 'Batch push qty increase ' + dateStr
      });
      Logger.log('Batch row ' + (rowIndex + 2) + ': ADD ' + delta + ' lot=' + wLot);
    } else {
      actions.push({
        type: 'REMOVE',
        sku: sku,
        qty: Math.abs(delta),
        site: wSite,
        location: wLocation,
        lot: wLot,
        dateCode: wDateCode,
        _rowLot: wLot,
        _rowDateCode: wDateCode,
        notes: 'Batch push qty decrease ' + dateStr
      });
      Logger.log('Batch row ' + (rowIndex + 2) + ': REMOVE ' + Math.abs(delta) + ' lot=' + wLot);
    }
    return actions;
  }

  // Case 2b: Site/Location/Lot/DateCode changed (with or without qty)
  if (anythingBesidesQty) {
    if (origQty > 0) {
      actions.push({
        type: 'REMOVE',
        sku: sku,
        qty: origQty,
        site: origSite || wSite,
        location: origLocation || wLocation,
        lot: origLot || wLot,
        dateCode: origDateCode || wDateCode,
        _rowLot: origLot,
        _rowDateCode: origDateCode,
        notes: 'Batch push change from ' + (origSite || wSite) + '/' + (origLocation || wLocation)
      });
    }
    if (wQty > 0) {
      actions.push({
        type: 'ADD',
        sku: sku,
        qty: wQty,
        site: wSite,
        location: wLocation,
        lot: wLot,
        dateCode: wDateCode,
        _rowLot: wLot,
        _rowDateCode: wDateCode,
        notes: 'Batch push change to ' + wSite + '/' + wLocation
      });
    }
    var changeDesc = [];
    if (siteChanged) changeDesc.push('site');
    if (locationChanged) changeDesc.push('location');
    if (lotChanged) changeDesc.push('lot');
    if (dateCodeChanged) changeDesc.push('dateCode');
    if (qtyChanged) changeDesc.push('qty');
    Logger.log('Batch row ' + (rowIndex + 2) + ': CHANGE (' + changeDesc.join('+') + ')');
    return actions;
  }

  return actions;
}

// ============================================
// KATANA TAB SYNC PUSH (create items in WASP)
// ============================================

/**
 * Push Katana tab rows with WASP Status = "SYNC" to WASP.
 * Creates items in WASP at MMH Kelowna / UNSORTED, then adds stock.
 * Lot-tracked items get per-batch ADD with lot + expiry (dateCode).
 * Non-lot items get a single ADD with total qty.
 *
 * Katana tab layout (11 cols):
 *   A(0)=SKU B(1)=Name C(2)=Location D(3)=Type E(4)=Lot Tracked
 *   F(5)=Batch G(6)=Expiry H(7)=UOM I(8)=Qty J(9)=Match K(10)=WASP Status
 *
 * @param {Spreadsheet} ss - Active spreadsheet
 * @returns {Object} { success: number, errors: number, errorDetails: Array }
 */
function pushKatanaSyncRows_(ss) {
  var result = { success: 0, errors: 0, errorDetails: [] };
  var sheet = ss.getSheetByName('Katana');
  if (!sheet) return result;

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return result;

  var KATANA_COLS = 15;
  var dataRange = sheet.getRange(2, 1, lastRow - 1, KATANA_COLS);
  var data = dataRange.getValues();
  var backgrounds = dataRange.getBackgrounds();

  // Group rows by SKU — ONLY 'Push' status (user must explicitly select from dropdown)
  // 'NEW' is the auto-assigned default for KATANA ONLY items — does NOT trigger push
  var syncSkus = {};  // sku -> { name, type, lotTracked, site, location, catalogOnly, rows: [] }
  for (var i = 0; i < data.length; i++) {
    var statusVal = String(data[i][11]).trim();
    if (statusVal !== 'Push') continue;
    var isCatalogOnly = false;  // Push creates item AND adds stock with batch/lot tracking

    var sku = String(data[i][0]).trim();
    if (!sku || sku.indexOf(' > ') > -1) continue; // skip separators

    // Smart location routing: type + Katana location
    var katanaLoc = String(data[i][2]).trim();  // col C
    var itemType = String(data[i][3]).trim();    // col D
    var targetSite = WASP_DEFAULT_SITE;
    var targetLocation = 'QA-Hold-1';

    if (katanaLoc === 'Storage Warehouse') {
      targetSite = 'Storage Warehouse';
      targetLocation = 'SW-STORAGE';
    } else if (katanaLoc === 'MMH Mayfair') {
      targetSite = 'MMH Mayfair';
      targetLocation = 'QA-Hold-1';
    } else if (itemType === 'Material' || itemType === 'Intermediate') {
      targetLocation = 'PRODUCTION';
    } else {
      targetLocation = 'UNSORTED';
    }

    if (!syncSkus[sku]) {
      syncSkus[sku] = {
        name: String(data[i][1]).trim(),
        type: itemType,
        lotTracked: String(data[i][4]).trim() === 'Yes',
        site: targetSite,
        location: targetLocation,
        catalogOnly: isCatalogOnly,
        rows: []
      };
    }

    // Parse expiry — handle Date objects from Sheets
    var expiry = '';
    if (data[i][6] instanceof Date) {
      var eY = data[i][6].getFullYear();
      var eM = data[i][6].getMonth() + 1;
      var eD = data[i][6].getDate();
      expiry = eY + '-' + (eM < 10 ? '0' + eM : eM) + '-' + (eD < 10 ? '0' + eD : eD);
    } else {
      expiry = String(data[i][6] || '').trim();
      if (expiry.indexOf('T') > -1) expiry = expiry.split('T')[0];
    }

    var qty = parseFloat(data[i][8]) || 0;

    syncSkus[sku].rows.push({
      rowIndex: i,
      qty: qty,
      batch: String(data[i][5] || '').trim(),
      expiry: expiry,
      site: targetSite,
      location: targetLocation
    });
  }

  var skuList = Object.keys(syncSkus);
  if (skuList.length === 0) return result;

  Logger.log('Katana push: found ' + skuList.length + ' SKUs with Push status');

  // Load SKU Lookup for UOM + Price + Cost
  var skuLookup = {};  // sku -> { uom, price, cost }
  var lookupSheet = ss.getSheetByName('SKU Lookup');
  if (lookupSheet && lookupSheet.getLastRow() > 1) {
    var lookupData = lookupSheet.getRange(2, 1, lookupSheet.getLastRow() - 1, 7).getValues();
    for (var li = 0; li < lookupData.length; li++) {
      var lSku = String(lookupData[li][0]).trim();
      if (lSku) {
        skuLookup[lSku] = {
          uom: String(lookupData[li][2] || 'ea').trim(),
          price: parseFloat(lookupData[li][5]) || 0,
          cost: parseFloat(lookupData[li][6]) || 0
        };
      }
    }
  }

  // Fetch purchase prices + categories from Katana API (for richer WASP items)
  var katanaPrices = {};   // sku -> purchasePrice
  var katanaCategories = {};  // sku -> category
  try {
    var allVariants = katanaFetchAllPages('/variants');
    var allProducts = katanaFetchAllPages('/products');
    var allMaterials = katanaFetchAllPages('/materials');

    // Build product category map
    var prodCatMap = {};
    for (var pi = 0; pi < allProducts.length; pi++) {
      var p = allProducts[pi];
      prodCatMap[p.id] = p.category || p.category_name || p.product_type || '';
    }
    // Build material category map
    var matCatMap = {};
    for (var mi = 0; mi < allMaterials.length; mi++) {
      var m = allMaterials[mi];
      matCatMap[m.id] = m.category || m.category_name || '';
    }
    // Build SKU -> purchasePrice + category from variants
    for (var vi = 0; vi < allVariants.length; vi++) {
      var v = allVariants[vi];
      if (!v.sku) continue;
      katanaPrices[v.sku] = v.purchase_price || v.default_purchase_price || v.cost || 0;
      if (v.product_id && prodCatMap[v.product_id]) {
        katanaCategories[v.sku] = prodCatMap[v.product_id];
      } else if (v.material_id && matCatMap[v.material_id]) {
        katanaCategories[v.sku] = matCatMap[v.material_id];
      }
    }
    Logger.log('Katana enrichment: ' + Object.keys(katanaPrices).length + ' prices, ' + Object.keys(katanaCategories).length + ' categories');
  } catch (e) {
    Logger.log('Katana enrichment failed (non-blocking): ' + e.message);
  }

  var dateStr = Utilities.formatDate(new Date(), 'America/Vancouver', 'yyyy-MM-dd');

  for (var si = 0; si < skuList.length; si++) {
    var sku = skuList[si];
    var info = syncSkus[sku];

    // Look up UOM + Price + Cost from SKU Lookup sheet
    var lookup = skuLookup[sku] || { uom: 'ea', price: 0, cost: 0 };

    // Step 1: Create item in WASP — INLINE (no external function dependency)
    // Normalize Katana UOM to WASP abbreviation (from WASP UoM settings)
    // WASP abbreviations: EA, dz, in, cu.in, lb, KG, PCS, BOX, g, pack, ROLLS
    var rawUom = String(lookup.uom || 'ea').trim().toLowerCase();
    var waspUom = 'EA';
    if (rawUom === 'each' || rawUom === 'ea' || rawUom === 'unit' || rawUom === '1' || rawUom === 'bucket') { waspUom = 'EA'; }
    else if (rawUom === 'pack' || rawUom === 'pk' || rawUom === '6pack') { waspUom = 'pack'; }
    else if (rawUom === 'pcs' || rawUom === 'pc') { waspUom = 'PCS'; }
    else if (rawUom === 'box') { waspUom = 'BOX'; }
    else if (rawUom === 'kg') { waspUom = 'KG'; }
    else if (rawUom === 'g' || rawUom === 'grams' || rawUom === 'gram') { waspUom = 'g'; }
    else if (rawUom === 'rolls' || rawUom === 'roll') { waspUom = 'ROLLS'; }
    else if (rawUom === 'dozen' || rawUom === 'dz') { waspUom = 'dz'; }
    else if (rawUom === 'pound' || rawUom === 'lb') { waspUom = 'lb'; }

    var createItem = {
      ItemNumber: sku,
      ItemDescription: (info.name || sku),
      StockingUnit: waspUom,
      PurchaseUnit: waspUom,
      SalesUnit: waspUom
    };
    // Cost: prefer Katana purchase_price, fall back to SKU Lookup cost
    var purchasePrice = katanaPrices[sku] || 0;
    if (purchasePrice > 0) {
      createItem.Cost = purchasePrice;
    } else if (lookup.cost > 0) {
      createItem.Cost = lookup.cost;
    }
    if (lookup.price > 0) createItem.SalesPrice = lookup.price;
    // Category from Katana (e.g. NON-INVENTORY, FINISHED GOOD, etc.)
    var katCategory = katanaCategories[sku] || '';
    if (katCategory) createItem.CategoryDescription = katCategory;
    if (info.lotTracked) {
      createItem.TrackbyInfo = { TrackedByLot: true, TrackedByDateCode: true };
    }
    // DimensionInfo — must set UOMs to prevent WASP from validating empty defaults
    createItem.DimensionInfo = { DimensionUnit: 'in', WeightUnit: 'lb', VolumeUnit: 'cu.in' };
    var cSite = info.site || WASP_DEFAULT_SITE;
    var cLoc = info.location || '';
    if (cLoc) {
      createItem.ItemLocationSettings = [{ SiteName: cSite, LocationCode: cLoc, PrimaryLocation: true }];
      createItem.ItemSiteSettings = [{ SiteName: cSite }];
    }
    Logger.log('WASP CREATE: ' + sku + ' UOM=' + waspUom + ' rawUom=' + rawUom + ' cost=' + (createItem.Cost || 0) + ' cat=' + (katCategory || 'none') + ' site=' + cSite + '/' + cLoc);
    var createResult = waspApiCall(getWaspBase() + 'ic/item/createInventoryItems', [createItem]);
    var itemCreated = createResult.success;
    var createErr = '';
    if (!createResult.success) {
      createErr = String(createResult.response || '').substring(0, 200);
      // Item may already exist — log but proceed to ADD stock
      Logger.log('Katana CREATE for ' + sku + ' result: ' + createErr + ' — proceeding to ADD stock');
      // Treat "already exists" as success for ADD purposes
      if (createErr.indexOf('already exist') > -1 || createErr.indexOf('-57') > -1) {
        itemCreated = true;
      }
    }

    // Step 2: Add stock per batch row with lot/batch tracking
    var skuSuccess = true;
    var skuError = '';
    var addAttempted = false;

    for (var ri = 0; ri < info.rows.length; ri++) {
      var batchRow = info.rows[ri];
      if (batchRow.qty <= 0) continue;

      addAttempted = true;

      Logger.log('WASP ADD STOCK: ' + sku + ' qty=' + batchRow.qty + ' batch=' + (batchRow.batch || 'none') + ' expiry=' + (batchRow.expiry || 'none') + ' site=' + batchRow.site + '/' + batchRow.location);
      var addAction = {
        type: 'ADD',
        sku: sku,
        qty: batchRow.qty,
        site: batchRow.site,
        location: batchRow.location,
        lot: batchRow.batch,
        dateCode: batchRow.expiry,
        _rowLot: batchRow.batch,
        _rowDateCode: batchRow.expiry,
        notes: 'Katana push ' + dateStr
      };

      var addResult = executeAction_(addAction);

      if (!addResult.success) {
        skuSuccess = false;
        skuError = addResult.error;
        Logger.log('Katana ADD failed for ' + sku + ' batch=' + batchRow.batch + ': ' + skuError);
        break;
      }
    }

    // If CREATE failed and no ADD was attempted (0-qty item), mark as ERROR
    if (!itemCreated && !addAttempted) {
      skuSuccess = false;
      skuError = createErr || 'Item creation failed';
    }

    // Step 3: Update status for all rows of this SKU
    for (var ui = 0; ui < info.rows.length; ui++) {
      var idx = info.rows[ui].rowIndex;
      if (skuSuccess) {
        data[idx][11] = 'Synced';
        for (var bg = 0; bg < KATANA_COLS; bg++) {
          backgrounds[idx][bg] = '#ffffff';
        }
      } else {
        data[idx][11] = 'ERROR';
        for (var bg2 = 0; bg2 < KATANA_COLS; bg2++) {
          backgrounds[idx][bg2] = '#ffcdd2';
        }
      }
    }

    if (skuSuccess) {
      result.success++;
      Logger.log('Katana push success: ' + sku);
    } else {
      result.errors++;
      result.errorDetails.push('Katana SKU ' + sku + ': ' + skuError);
    }

    // Rate limiting every 5 SKUs
    if ((si + 1) % 5 === 0) {
      Utilities.sleep(500);
      dataRange.setValues(data);
      dataRange.setBackgrounds(backgrounds);
      SpreadsheetApp.flush();
    }
  }

  // Final write back
  dataRange.setValues(data);
  dataRange.setBackgrounds(backgrounds);
  SpreadsheetApp.flush();

  Logger.log('Katana push complete: ' + result.success + ' SKUs success, ' + result.errors + ' errors');
  return result;
}

// ============================================
// KATANA TAB QTY ADJUSTMENTS (set WASP to match edited Katana qty)
// ============================================

/**
 * Push Katana tab rows with Adj Status = "Pending" to WASP.
 * For each pending row: computes delta = editedQty - waspQty
 *   delta > 0 â†’ WASP ADD
 *   delta < 0 â†’ WASP REMOVE
 * Uses smart location routing (same as pushKatanaSyncRows_).
 *
 * Katana tab layout (14 cols):
 *     A(0)=SKU B(1)=Name C(2)=Location D(3)=Type E(4)=Lot Tracked
 *     F(5)=Batch G(6)=Expiry H(7)=UOM I(8)=Qty J(9)=WASP Qty
 *     K(10)=Match L(11)=WASP Status M(12)=Orig Qty N(13)=Adj Status
 *
 * @param {Spreadsheet} ss - Active spreadsheet
 * @returns {Object} { success: number, errors: number, errorDetails: Array }
 */
function pushKatanaAdjustments_(ss) {
    var result = { success: 0, errors: 0, errorDetails: [] };
    var sheet = ss.getSheetByName('Katana');
    if (!sheet) return result;

    var activeUser = '';
    try { activeUser = Session.getActiveUser().getEmail(); } catch (e) {}
    if (!activeUser) { try { activeUser = Session.getEffectiveUser().getEmail(); } catch (e) {} }
    if (!activeUser) { activeUser = PropertiesService.getScriptProperties().getProperty('USER_NAME') || ''; }

    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return result;

    var KATANA_COLS = 15;
    var dataRange = sheet.getRange(2, 1, lastRow - 1, KATANA_COLS);
    var data = dataRange.getValues();
    var backgrounds = dataRange.getBackgrounds();

    // Find rows with Adj Status === 'Pending' or 'ERROR' (col 14, index 13)
    // ERROR rows are retried on the next push attempt just like Pending rows.
    var pendingRows = [];
    for (var i = 0; i < data.length; i++) {
        var adjStatus = String(data[i][13]).trim();
        if (adjStatus === 'Pending' || adjStatus === 'ERROR') {
            pendingRows.push(i);
        }
    }

    if (pendingRows.length === 0) return result;

    Logger.log('Katana adj push: found ' + pendingRows.length + ' pending rows');

    var dateStr = Utilities.formatDate(new Date(), 'America/Vancouver', 'yyyy-MM-dd');

    for (var j = 0; j < pendingRows.length; j++) {
        var rowIndex = pendingRows[j];
        var row = data[rowIndex];
        var sheetRowNum = rowIndex + 2;

        var sku = String(row[0]).trim();
        if (!sku || sku.indexOf(' > ') > -1) continue;  // skip separators

        var name = String(row[1]).trim();
        var katanaLoc = String(row[2]).trim();
        // col D (row[3]) = Katana category (PACKAGING, RAW MATER, etc.) — for display only
        // col O (row[14]) = Katana item type (Material, Intermediate, Product) — used for WASP location routing
        var itemType = String(row[14] || row[3] || '').trim();
        var lotTracked = String(row[4]).trim() === 'Yes';
        var batch = String(row[5]).trim();
        var expiry = '';
        if (row[6] instanceof Date) {
            var eY = row[6].getFullYear();
            var eM = row[6].getMonth() + 1;
            var eD = row[6].getDate();
            expiry = eY + '-' + (eM < 10 ? '0' + eM : eM) + '-' + (eD < 10 ? '0' + eD : eD);
        } else {
            expiry = String(row[6] || '').trim();
            if (expiry.indexOf('T') > -1) expiry = expiry.split('T')[0];
        }

        var editedQty = (row[9] instanceof Date) ? 0 : (parseFloat(row[9]) || 0);
        // WASP Qty is now the editable column (row[9]); orig WASP value is in row[12]
        var origQty = (row[12] instanceof Date) ? 0 : (parseFloat(row[12]) || 0);

        // Smart location routing (default)
        var targetSite = WASP_DEFAULT_SITE;
        var targetLocation = 'QA-Hold-1';
        if (katanaLoc === 'Storage Warehouse') {
            targetSite = 'Storage Warehouse';
            targetLocation = 'SW-STORAGE';
        } else if (katanaLoc === 'MMH Mayfair') {
            targetSite = 'MMH Mayfair';
            targetLocation = 'QA-Hold-1';
        } else if (itemType === 'Material' || itemType === 'Intermediate') {
            targetLocation = 'PRODUCTION';
        } else {
            targetLocation = 'UNSORTED';
        }

        // For lot-tracked rows, look up actual WASP location (smart routing may guess wrong)
        if (lotTracked && batch) {
            try {
                var searchResult = waspFetch('ic/item/advancedinventorysearch', {
                    ItemNumber: sku, Lot: batch, PageSize: 20, PageNumber: 1
                });
                var sItems = searchResult.Items || searchResult.items || [];
                for (var si = 0; si < sItems.length; si++) {
                    var sItem = sItems[si];
                    var sLot = String(sItem.Lot || sItem.lot || '').trim();
                    var sQty = parseFloat(sItem.TotalInHouse || sItem.Quantity || sItem.TotalAvailable || 0);
                    if (sLot === batch && sQty > 0) {
                        targetSite = sItem.SiteName || sItem.siteName || targetSite;
                        targetLocation = sItem.LocationCode || sItem.locationCode || targetLocation;
                        Logger.log('Katana adj: found ' + sku + '/' + batch + ' at ' + targetSite + '/' + targetLocation + ' qty=' + sQty);
                        break;
                    }
                }
            } catch (lookupErr) {
                Logger.log('Katana adj: WASP lookup failed for ' + sku + '/' + batch + ': ' + lookupErr.message);
            }
        }

        // For REMOVE deltas: look up actual WASP location regardless of lot-tracking.
        // Smart routing (PACKAGING→UNSORTED, etc.) may guess wrong, causing -46002.
        // This runs after the lot-based lookup above so it can override if needed.
        var deltaCheck = editedQty - origQty;
        if (deltaCheck < 0) {
            try {
                var preSearch = waspFetch('ic/item/advancedinventorysearch', {
                    ItemNumber: sku, PageSize: 50, PageNumber: 1
                });
                var preList = (preSearch.Data && preSearch.Data.ResultList) ? preSearch.Data.ResultList :
                              (preSearch.Items || preSearch.items || []);
                for (var pli = 0; pli < preList.length; pli++) {
                    var pItem = preList[pli];
                    var pLotOk = !batch || (String(pItem.Lot || '').trim() === batch);
                    var pQty = parseFloat(pItem.QtyAvailable || 0);
                    if (String(pItem.ItemNumber || '').trim() === sku && pQty > 0 && pLotOk) {
                        targetSite = String(pItem.SiteName || targetSite).trim();
                        targetLocation = String(pItem.LocationCode || targetLocation).trim();
                        Logger.log('Katana adj REMOVE pre-lookup: ' + sku + ' found at ' + targetSite + '/' + targetLocation + ' qty=' + pQty);
                        break;
                    }
                }
            } catch (preErr) {
                Logger.log('Katana adj REMOVE pre-lookup failed for ' + sku + ': ' + preErr.message);
            }
        }

        // Compute delta: editedQty vs origQty (original WASP value)
        var delta = editedQty - origQty;

        if (Math.abs(delta) < 0.01) {
            // No actual change needed
            data[rowIndex][13] = 'Pushed ' + formatTimestamp();
            data[rowIndex][12] = editedQty;  // update orig qty
            for (var bg = 0; bg < KATANA_COLS; bg++) {
                backgrounds[rowIndex][bg] = '#ffffff';
            }
            result.success++;
            logAdjustment_(ss, 'Google Sheet', 'Sync', activeUser, sku, '', targetSite, targetLocation, batch, expiry, Math.round((editedQty - origQty) * 100) / 100, 'OK');
            continue;
        }

        Logger.log('Katana adj row ' + sheetRowNum + ': ' + sku + ' delta=' + delta);

        var action;
        if (delta > 0) {
            action = {
                type: 'ADD',
                sku: sku,
                qty: delta,
                site: targetSite,
                location: targetLocation,
                lot: batch,
                dateCode: expiry,
                _rowLot: batch,
                _rowDateCode: expiry,
                notes: 'Katana adj push increase ' + dateStr
            };
        } else {
            action = {
                type: 'REMOVE',
                sku: sku,
                qty: Math.abs(delta),
                site: targetSite,
                location: targetLocation,
                lot: batch,
                dateCode: expiry,
                _rowLot: batch,
                _rowDateCode: expiry,
                notes: 'Katana adj push decrease ' + dateStr
            };
        }

        var actionResult = executeAction_(action);

        if (actionResult.success) {
            data[rowIndex][13] = 'Pushed ' + formatTimestamp();
            data[rowIndex][12] = editedQty;  // update orig qty to new value
            for (var bg2 = 0; bg2 < KATANA_COLS; bg2++) {
                backgrounds[rowIndex][bg2] = '#ffffff';
            }
            result.success++;
            logAdjustment_(ss, 'Google Sheet', (action.type === 'ADD' ? 'Add' : 'Remove'), activeUser, sku, '', targetSite, targetLocation, batch, expiry, Math.round((editedQty - origQty) * 100) / 100, 'OK');
        } else {
            data[rowIndex][13] = 'ERROR';
            for (var bg3 = 0; bg3 < KATANA_COLS; bg3++) {
                backgrounds[rowIndex][bg3] = '#ffcdd2';
            }
            result.errors++;
            result.errorDetails.push('Katana adj row ' + sheetRowNum + ' (' + sku + '): ' + actionResult.error);
            logAdjustment_(ss, 'Google Sheet', (action.type === 'ADD' ? 'Add' : 'Remove'), activeUser, sku, '', targetSite, targetLocation, batch, expiry, Math.round((editedQty - origQty) * 100) / 100, 'ERROR');
        }

        // Rate limiting every 10 rows
        if ((j + 1) % 10 === 0) {
            Utilities.sleep(500);
            dataRange.setValues(data);
            dataRange.setBackgrounds(backgrounds);
            SpreadsheetApp.flush();
        }
    }

    // Final write back
    dataRange.setValues(data);
    dataRange.setBackgrounds(backgrounds);
    SpreadsheetApp.flush();

    Logger.log('Katana adj push complete: ' + result.success + ' success, ' + result.errors + ' errors');
    return result;
}

// ============================================
// WASP CATALOG ITEM CREATION (self-contained)
// ============================================

/**
 * Normalize Katana UOM to valid WASP UOM.
 * WASP only accepts UOMs that exist in their system.
 */
function normalizeUomForWasp_(uom) {
  if (!uom) return 'EA';
  var u = String(uom).trim().toLowerCase();
  if (u === 'each' || u === 'ea' || u === 'unit' || u === '1' || u === 'bucket') return 'EA';
  if (u === 'pack' || u === 'pk' || u === '6pack') return 'pack';
  if (u === 'pcs' || u === 'pc') return 'PCS';
  if (u === 'box') return 'BOX';
  if (u === 'kg') return 'KG';
  if (u === 'g' || u === 'grams' || u === 'gram') return 'g';
  if (u === 'rolls' || u === 'roll') return 'ROLLS';
  if (u === 'dozen' || u === 'dz') return 'dz';
  if (u === 'pound' || u === 'lb') return 'lb';
  return 'EA';
}

/**
 * Create a WASP inventory item with ALL required fields.
 * Self-contained — does not depend on 04_SyncEngine.
 *
 * Required per WASP API docs:
 *   ItemNumber, StockingUnit, PurchaseUnit, SalesUnit, Cost, CostMethod, TaxCode
 */
function createWaspCatalogItem_(sku, name, site, location, lotTracking, uom, salesPrice, cost) {
  var url = getWaspBase() + 'ic/item/createInventoryItems';
  var normalizedUom = normalizeUomForWasp_(uom);

  var item = {
    ItemNumber: sku,
    ItemDescription: name || sku,
    StockingUnit: normalizedUom,
    PurchaseUnit: normalizedUom,
    SalesUnit: normalizedUom
  };

  // Cost — send from Katana avg cost, default 0
  if (cost !== undefined && cost !== null && cost > 0) {
    item.Cost = cost;
  }

  if (salesPrice !== undefined && salesPrice !== null && salesPrice > 0) {
    item.SalesPrice = salesPrice;
  }

  if (lotTracking) {
    item.TrackbyInfo = {
      TrackedByLot: true,
      TrackedByDateCode: true
    };
  }

  var targetSite = site || WASP_DEFAULT_SITE;
  var targetLoc = location || (targetSite === 'MMH Mayfair' ? 'QA-Hold-1' : '');
  if (targetLoc) {
    item.ItemLocationSettings = [{
      SiteName: targetSite,
      LocationCode: targetLoc,
      PrimaryLocation: true
    }];
    item.ItemSiteSettings = [{
      SiteName: targetSite
    }];
  }

  Logger.log('createWaspCatalogItem_: ' + sku + ' UOM=' + normalizedUom + ' site=' + targetSite + '/' + targetLoc);

  var payload = [item];
  return waspApiCall(url, payload);
}

// ============================================
// AUTO-ADJUST WASP (ADJUSTABLE rows only)
// ============================================

/**
 * Automatically adjust WASP inventory for all ADJUSTABLE rows to match Katana.
 * ADJUSTABLE = truly mismatched, non-lot-tracked, no UOM conversion.
 *
 * ADD delta  → add to dominant location (highest W.Qty).
 * REMOVE delta → drain from locations sorted highest-to-lowest until delta satisfied.
 *   This avoids -46002 "insufficient qty" errors when stock is spread across locations.
 *
 * Called from menu: Katana Sync > Auto-Adjust WASP
 */
function autoAdjustWasp() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Wasp');
  if (!sheet) {
    SpreadsheetApp.getUi().alert('Wasp sheet not found.');
    return;
  }

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    ss.toast('No data in Wasp sheet.', 'Auto-Adjust', 5);
    return;
  }

  var NUM_COLS = 18;
  var data = sheet.getRange(2, 1, lastRow - 1, NUM_COLS).getValues();

  var activeUser = '';
  try { activeUser = Session.getActiveUser().getEmail(); } catch (e) {}
  if (!activeUser) { try { activeUser = Session.getEffectiveUser().getEmail(); } catch (e) {} }
  if (!activeUser) { activeUser = PropertiesService.getScriptProperties().getProperty('USER_NAME') || ''; }

  var dateStr = Utilities.formatDate(new Date(), 'America/Vancouver', 'yyyy-MM-dd');

  // Pass 1: build SKU+site groups from ADJUSTABLE rows only
  var groups = {};
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var sku = String(row[0] || '').trim();
    var match = String(row[14] || '').trim();
    if (!sku || sku.indexOf(' > ') > -1 || match !== 'ADJUSTABLE') continue;

    var site = String(row[3] || '').trim();
    var location = String(row[4] || '').trim();
    var kQty = (row[9] instanceof Date) ? 0 : (parseFloat(row[9]) || 0);
    var wQty = (row[11] instanceof Date) ? 0 : (parseFloat(row[11]) || 0);

    var key = sku + '|' + site;
    if (!groups[key]) {
      groups[key] = { sku: sku, site: site, kQty: kQty, wTotal: 0, rows: [] };
    }
    groups[key].wTotal += wQty;
    groups[key].rows.push({ wQty: wQty, location: location });
  }

  var groupKeys = Object.keys(groups);
  if (groupKeys.length === 0) {
    ss.toast('No ADJUSTABLE rows found. Run Sync Now first.', 'Auto-Adjust', 5);
    return;
  }

  var successCount = 0;
  var errorCount = 0;
  var errorDetails = [];
  var skippedCount = 0;
  var startTime = new Date();
  var TIME_LIMIT_MS = 330000; // 5.5 minutes — safe under GAS 6-min limit

  Logger.log('autoAdjustWasp: starting ' + groupKeys.length + ' groups');
  try { ss.toast('Starting ' + groupKeys.length + ' adjustments...', 'Auto-Adjust', 5); } catch (e) {}

  // Pass 2: apply adjustments per group
  for (var g = 0; g < groupKeys.length; g++) {
    // Time-limit safety: stop before GAS kills the execution
    if ((new Date() - startTime) > TIME_LIMIT_MS) {
      var remaining = groupKeys.length - g;
      Logger.log('autoAdjustWasp: time limit reached with ' + remaining + ' groups remaining');
      try {
        ss.toast(successCount + ' done, ' + remaining + ' remaining — re-run to continue.', 'Auto-Adjust Paused', 10);
      } catch (e) {}
      return;
    }

    var grp = groups[groupKeys[g]];
    var delta = Math.round((grp.kQty - grp.wTotal) * 10000) / 10000;

    if (Math.abs(delta) < 0.01) {
      skippedCount++;
      continue;
    }

    // Progress toast every 10 items
    if (g > 0 && g % 10 === 0) {
      try { ss.toast(successCount + ' done, processing ' + (g + 1) + '/' + groupKeys.length + '...', 'Auto-Adjust', 4); } catch (e) {}
    }

    // Find dominant (highest W.Qty) — used for ADD target
    var dominantRow = grp.rows[0];
    for (var r = 1; r < grp.rows.length; r++) {
      if (grp.rows[r].wQty > dominantRow.wQty) dominantRow = grp.rows[r];
    }

    var groupSuccess = true;
    var groupError = '';

    if (delta > 0) {
      // ADD: put the full delta into the dominant location
      var addAction = {
        type: 'ADD',
        sku: grp.sku,
        qty: delta,
        site: grp.site,
        location: dominantRow.location,
        lot: '', dateCode: '', _rowLot: '', _rowDateCode: '',
        notes: 'Auto-adjust add ' + dateStr
      };
      Logger.log('[' + (g + 1) + '/' + groupKeys.length + '] ADD ' + delta +
                 ' of ' + grp.sku + ' @ ' + grp.site + '/' + dominantRow.location +
                 ' (K=' + grp.kQty + ' W=' + grp.wTotal + ')');
      var addResult = executeAction_(addAction);
      if (!addResult.success) { groupSuccess = false; groupError = addResult.error; }

    } else {
      // REMOVE: drain from locations sorted highest-to-lowest until delta is satisfied.
      // This avoids trying to remove more than any single location holds.
      var rowsSorted = grp.rows.slice().sort(function(a, b) { return b.wQty - a.wQty; });
      var toRemove = Math.abs(delta);

      Logger.log('[' + (g + 1) + '/' + groupKeys.length + '] REMOVE ' + toRemove +
                 ' of ' + grp.sku + ' @ ' + grp.site +
                 ' across ' + rowsSorted.length + ' location(s)' +
                 ' (K=' + grp.kQty + ' W=' + grp.wTotal + ')');

      for (var rd = 0; rd < rowsSorted.length && toRemove > 0.01; rd++) {
        var locRow = rowsSorted[rd];
        if (locRow.wQty <= 0) continue;

        var removeQty = Math.min(
          Math.round(locRow.wQty * 10000) / 10000,
          Math.round(toRemove * 10000) / 10000
        );

        var removeAction = {
          type: 'REMOVE',
          sku: grp.sku,
          qty: removeQty,
          site: grp.site,
          location: locRow.location,
          lot: '', dateCode: '', _rowLot: '', _rowDateCode: '',
          notes: 'Auto-adjust remove ' + dateStr
        };
        Logger.log('  -> REMOVE ' + removeQty + ' from ' + locRow.location + ' (has ' + locRow.wQty + ')');
        var remResult = executeAction_(removeAction);
        if (remResult.success) {
          toRemove = Math.round((toRemove - removeQty) * 10000) / 10000;
        } else {
          groupSuccess = false;
          // Clean up the error message — strip raw JSON, show just the WASP message
          var rawErr = String(remResult.error || '');
          var cleanErr = rawErr;
          try {
            var errObj = JSON.parse(rawErr);
            var msgs = (errObj.Messages || []);
            if (msgs.length > 0) cleanErr = msgs[0].Message + ' (' + msgs[0].ResultCode + ')';
          } catch (ep) { /* keep rawErr as-is if not JSON */ }
          groupError = cleanErr;
          Logger.log('  -> FAILED at ' + locRow.location + ': ' + cleanErr);
          break;
        }
      }
    }

    if (groupSuccess) {
      successCount++;
      var logAction = delta > 0 ? 'Add' : 'Remove';
      logAdjustment_(ss, 'Auto-Adjust', logAction, activeUser, grp.sku, '',
                     grp.site, dominantRow.location, '', '',
                     Math.round(delta * 100) / 100, 'OK');
    } else {
      errorCount++;
      errorDetails.push(grp.sku + ': ' + groupError);
      Logger.log('autoAdjustWasp ERROR: ' + grp.sku + ' — ' + groupError);
      logAdjustment_(ss, 'Auto-Adjust', delta > 0 ? 'Add' : 'Remove', activeUser,
                     grp.sku, '', grp.site, dominantRow.location, '', '',
                     Math.round(delta * 100) / 100, 'ERROR');
    }
  }

  var duration = Math.round((new Date() - startTime) / 1000);
  var msg = successCount + ' adjusted';
  if (skippedCount > 0) msg += ', ' + skippedCount + ' skipped';
  if (errorCount > 0)   msg += ', ' + errorCount + ' errors';
  msg += ' (' + duration + 's)';
  Logger.log('autoAdjustWasp complete: ' + msg);

  try {
    ss.toast(msg + '. Run Sync Now to refresh.', 'Auto-Adjust Complete', 10);
  } catch (e) {
    Logger.log('Toast failed: ' + e.message);
  }

  if (errorDetails.length > 0 && errorDetails.length <= 10) {
    SpreadsheetApp.getUi().alert('Auto-Adjust errors:\n' + errorDetails.join('\n'));
  }
}
