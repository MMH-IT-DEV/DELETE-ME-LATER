// ============================================
// 97_WebhookAudit.gs - WEBHOOK CROSS-CHECK
// ============================================
// Pulls recent events from Katana API and compares
// against Activity tab to find missed webhooks.
// DELETE THIS FILE when no longer needed.
// ============================================

/**
 * Main function — run this to audit webhooks.
 * Checks MOs, STs, SOs, POs from the last N days
 * and flags any that are missing from the Activity tab.
 */
function webhookAudit_RUN() {
  var DAYS_BACK = 2;
  var cutoff = new Date();
  cutoff.setDate(cutoff.getDate() - DAYS_BACK);
  var cutoffStr = cutoff.toISOString().split('T')[0];
  var recentCutoff = new Date();
  recentCutoff.setDate(recentCutoff.getDate() - 1);
  var recentStr = recentCutoff.toISOString().split('T')[0];

  Logger.log('===== WEBHOOK AUDIT =====');
  Logger.log('Checking events since: ' + cutoffStr);
  Logger.log('');

  // Read Activity tab to build a set of known IDs
  var knownIds = readActivityIds();
  Logger.log('Activity tab entries scanned: ' + Object.keys(knownIds).length);
  Logger.log('');

  // Write results to the comparison sheet
  var ss = SpreadsheetApp.openById(COMPARISON_SHEET_ID);
  var tab = ss.getSheetByName('Webhook Audit');
  if (!tab) {
    tab = ss.insertSheet('Webhook Audit');
  } else {
    tab.clear();
  }

  var headers = ['Type', 'Katana ID', 'Reference', 'Status', 'Updated At', 'In Activity?', 'Recent?', 'Notes'];
  tab.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
  tab.setFrozenRows(1);

  var allRows = [];
  var missing = 0;
  var missingRecent = 0;
  var found = 0;

  // Helper: lowercase status check
  function statusIn(val, list) {
    var lower = String(val || '').toLowerCase();
    for (var x = 0; x < list.length; x++) {
      if (lower === list[x]) return true;
    }
    return false;
  }

  // Helper: log unique statuses from first batch
  function logStatuses(label, items) {
    var seen = {};
    for (var x = 0; x < items.length; x++) {
      var s = String(items[x].status || 'null');
      seen[s] = (seen[s] || 0) + 1;
    }
    var parts = [];
    for (var key in seen) {
      parts.push(key + ':' + seen[key]);
    }
    Logger.log('  Statuses: ' + parts.join(', '));
  }

  // Helper: log first item's keys for debugging
  function logSampleKeys(label, items) {
    if (items.length > 0) {
      var keys = [];
      for (var key in items[0]) {
        keys.push(key);
      }
      Logger.log('  Sample keys: ' + keys.join(', '));
      Logger.log('  Sample: id=' + items[0].id + ' order_no=' + (items[0].order_no || items[0].stock_transfer_number || '') + ' status=' + items[0].status + ' updated_at=' + items[0].updated_at);
    }
  }

  // 1. Manufacturing Orders
  Logger.log('--- Manufacturing Orders ---');
  var mos = fetchKatanaEventsPaginated('manufacturing_orders?per_page=50&sort=-updated_at');
  logSampleKeys('MO', mos);
  logStatuses('MO', mos);
  for (var i = 0; i < mos.length; i++) {
    var mo = mos[i];
    var moUpdated = String(mo.updated_at || mo.updatedAt || '');
    if (moUpdated && moUpdated < cutoffStr) break;
    var moNum = String(mo.order_no || '');
    var moId = String(mo.id || '');
    var inActivity = knownIds['MO-' + moId] || knownIds[moNum] || knownIds['MO-' + moNum] || false;
    var status = String(mo.status || '').toLowerCase();
    // Include done and resource_completed (webhook-triggering statuses)
    if (statusIn(status, ['done', 'completed', 'resource_completed', 'in_progress'])) {
      var isRecent = moUpdated >= recentStr;
      var row = ['MO', moId, moNum, status, moUpdated, inActivity ? 'YES' : 'MISSING', isRecent ? 'RECENT' : '', ''];
      allRows.push(row);
      if (!inActivity && (status === 'done' || status === 'completed' || status === 'resource_completed')) {
        missing++;
        if (isRecent) missingRecent++;
      } else {
        found++;
      }
    }
  }
  Logger.log('  Matched: ' + allRows.length + ' of ' + mos.length);

  // 2. Stock Transfers
  var moCount = allRows.length;
  Logger.log('--- Stock Transfers ---');
  var sts = fetchKatanaEventsPaginated('stock_transfers?per_page=50&sort=-updated_at');
  logSampleKeys('ST', sts);
  logStatuses('ST', sts);
  for (var j = 0; j < sts.length; j++) {
    var st = sts[j];
    var stUpdated = String(st.updated_at || st.updatedAt || '');
    if (stUpdated && stUpdated < cutoffStr) break;
    var stNum = String(st.stock_transfer_number || '');
    var stId = String(st.id || '');
    var stInActivity = knownIds['ST-' + stId] || knownIds['#' + stId] || knownIds[stNum] || knownIds['#' + stNum] || false;
    var stStatus = String(st.status || '').toLowerCase();
    if (statusIn(stStatus, ['completed', 'done', 'in_progress', 'received'])) {
      var stRecent = stUpdated >= recentStr;
      var stRow = ['ST', stId, stNum, stStatus, stUpdated, stInActivity ? 'YES' : 'MISSING', stRecent ? 'RECENT' : '', ''];
      allRows.push(stRow);
      if (!stInActivity && (stStatus === 'completed' || stStatus === 'done' || stStatus === 'received')) {
        missing++;
        if (stRecent) missingRecent++;
      } else {
        found++;
      }
    }
  }
  Logger.log('  Matched: ' + (allRows.length - moCount) + ' of ' + sts.length);

  // 3. Sales Orders
  var stCount = allRows.length;
  Logger.log('--- Sales Orders ---');
  var sos = fetchKatanaEventsPaginated('sales_orders?per_page=50&sort=-updated_at');
  logSampleKeys('SO', sos);
  logStatuses('SO', sos);
  for (var k = 0; k < sos.length; k++) {
    var so = sos[k];
    var soUpdated = String(so.updated_at || so.updatedAt || '');
    if (soUpdated && soUpdated < cutoffStr) break;
    var soNum = String(so.order_no || '');
    var soId = String(so.id || '');
    var soInActivity = knownIds['SO-' + soId] || knownIds['#' + soId] || knownIds[soNum] || knownIds['#' + soNum] || false;
    var soStatus = String(so.status || '').toLowerCase();
    if (statusIn(soStatus, ['fulfilled', 'delivered', 'shipped', 'done', 'completed'])) {
      var soRecent = soUpdated >= recentStr;
      var soRow = ['SO', soId, soNum, soStatus, soUpdated, soInActivity ? 'YES' : 'MISSING', soRecent ? 'RECENT' : '', ''];
      allRows.push(soRow);
      if (!soInActivity) {
        missing++;
        if (soRecent) missingRecent++;
      } else {
        found++;
      }
    }
  }
  Logger.log('  Matched: ' + (allRows.length - stCount) + ' of ' + sos.length);

  // 4. Purchase Orders
  var soCount = allRows.length;
  Logger.log('--- Purchase Orders ---');
  var pos = fetchKatanaEventsPaginated('purchase_orders?per_page=50&sort=-updated_at');
  logSampleKeys('PO', pos);
  logStatuses('PO', pos);
  for (var l = 0; l < pos.length; l++) {
    var po = pos[l];
    var poUpdated = String(po.updated_at || po.updatedAt || '');
    if (poUpdated && poUpdated < cutoffStr) break;
    var poNum = String(po.order_no || '');
    var poId = String(po.id || '');
    var poInActivity = knownIds['PO-' + poId] || knownIds['PO-' + poNum] || false;
    var poStatus = String(po.status || '').toLowerCase();
    if (statusIn(poStatus, ['received', 'partially_received', 'done', 'completed'])) {
      var poRecent = poUpdated >= recentStr;
      var poRow = ['PO', poId, poNum, poStatus, poUpdated, poInActivity ? 'YES' : 'MISSING', poRecent ? 'RECENT' : '', ''];
      allRows.push(poRow);
      if (!poInActivity) {
        missing++;
        if (poRecent) missingRecent++;
      } else {
        found++;
      }
    }
  }
  Logger.log('  Matched: ' + (allRows.length - soCount) + ' of ' + pos.length);

  // Write all rows
  if (allRows.length > 0) {
    tab.getRange(2, 1, allRows.length, headers.length).setValues(allRows);
  }

  // Summary row
  var summaryRow = allRows.length + 3;
  tab.getRange(summaryRow, 1, 5, 2).setValues([
    ['SUMMARY', ''],
    ['Total events checked', allRows.length],
    ['Found in Activity', found],
    ['MISSING (total)', missing],
    ['MISSING (last 24h)', missingRecent]
  ]);

  Logger.log('');
  Logger.log('===== AUDIT COMPLETE =====');
  Logger.log('Total events: ' + allRows.length);
  Logger.log('Found in Activity: ' + found);
  Logger.log('MISSING (total): ' + missing);
  Logger.log('MISSING (last 24h): ' + missingRecent);
}

/**
 * Read Activity tab and build a lookup of all known event references.
 * Extracts IDs like MO-15260040, #91092, PO-542, ST-224 from the Details column.
 */
function readActivityIds() {
  var ss = SpreadsheetApp.openById(CONFIG.DEBUG_SHEET_ID);
  var activitySheet = ss.getSheetByName('Activity');
  if (!activitySheet) return {};

  var data = activitySheet.getDataRange().getValues();
  var ids = {};

  for (var i = 3; i < data.length; i++) {
    var details = String(data[i][3] || '');

    // Extract MO references: "MO-7143", "MO-15260040"
    var moMatches = details.match(/MO-\d+/g);
    if (moMatches) {
      for (var m = 0; m < moMatches.length; m++) {
        ids[moMatches[m]] = true;
      }
    }

    // Extract # references: "#91092", "#90930"
    var hashMatches = details.match(/#\d+/g);
    if (hashMatches) {
      for (var h = 0; h < hashMatches.length; h++) {
        ids[hashMatches[h]] = true;
      }
    }

    // Extract PO references: "PO-542"
    var poMatches = details.match(/PO-\d+/g);
    if (poMatches) {
      for (var p = 0; p < poMatches.length; p++) {
        ids[poMatches[p]] = true;
      }
    }

    // Extract ST references: "ST-224"
    var stMatches = details.match(/ST-\d+/g);
    if (stMatches) {
      for (var s = 0; s < stMatches.length; s++) {
        ids[stMatches[s]] = true;
      }
    }

    // Extract SO references from SO# patterns
    var soMatches = details.match(/SO-\d+/g);
    if (soMatches) {
      for (var so = 0; so < soMatches.length; so++) {
        ids[soMatches[so]] = true;
      }
    }

    // Also store raw Katana internal IDs from MO lines like "MO-15260040"
    var katanaIdMatches = details.match(/MO-\d{8}/g);
    if (katanaIdMatches) {
      for (var ki = 0; ki < katanaIdMatches.length; ki++) {
        ids[katanaIdMatches[ki]] = true;
      }
    }
  }

  return ids;
}

/**
 * Backfill missed SOs — re-process SO deliveries that were dropped by the old doPost()
 * Reads the "Webhook Audit" tab, finds rows where Type=SO, In Activity?=MISSING, Recent?=RECENT
 * Builds a fake webhook payload for each and calls routeWebhook()
 * Run once after deploying the queue system to catch up the missed SOs
 */
function backfillMissedSOs() {
  var ss = SpreadsheetApp.openById(COMPARISON_SHEET_ID);
  var tab = ss.getSheetByName('Webhook Audit');
  if (!tab) {
    Logger.log('No "Webhook Audit" tab found — run webhookAudit_RUN() first');
    return;
  }

  var lastRow = tab.getLastRow();
  if (lastRow <= 1) {
    Logger.log('No data in Webhook Audit tab');
    return;
  }

  // Headers: Type | Katana ID | Reference | Status | Updated At | In Activity? | Recent? | Notes
  var data = tab.getRange(2, 1, lastRow - 1, 8).getValues();
  var backfilled = 0;
  var skipped = 0;
  var errors = 0;

  Logger.log('===== BACKFILL MISSED SOs =====');

  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var type = String(row[0] || '');
    var katanaId = String(row[1] || '');
    var reference = String(row[2] || '');
    var status = String(row[3] || '');
    var inActivity = String(row[5] || '');
    var recent = String(row[6] || '');

    // Only process SO rows that are MISSING and RECENT
    if (type !== 'SO') continue;
    if (inActivity !== 'MISSING') continue;

    // Build fake webhook payload matching Katana format
    var fakePayload = {
      action: 'sales_order.delivered',
      object: { id: parseInt(katanaId, 10) }
    };

    Logger.log('Backfilling SO ' + reference + ' (ID: ' + katanaId + ', status: ' + status + ')');

    try {
      var result = routeWebhook(fakePayload);
      var resultStatus = result ? (result.status || 'unknown') : 'null';

      // Mark the audit row as BACKFILLED in the Notes column (col 8)
      var sheetRow = i + 2;
      tab.getRange(sheetRow, 8).setValue('BACKFILLED — ' + resultStatus);

      if (resultStatus === 'error' || resultStatus === 'failed') {
        errors++;
        Logger.log('  ERROR: ' + JSON.stringify(result));
      } else {
        backfilled++;
        Logger.log('  OK: ' + resultStatus);
      }

      // 500ms pause between API calls
      Utilities.sleep(500);

    } catch (err) {
      errors++;
      var errRow = i + 2;
      tab.getRange(errRow, 8).setValue('BACKFILL ERROR: ' + err.message);
      Logger.log('  EXCEPTION: ' + err.message);
    }
  }

  Logger.log('');
  Logger.log('===== BACKFILL COMPLETE =====');
  Logger.log('Backfilled: ' + backfilled);
  Logger.log('Skipped: ' + skipped);
  Logger.log('Errors: ' + errors);
}

/**
 * Fetch paginated results from Katana API.
 * Returns all items from all pages (up to 500 max to avoid timeout).
 */
function fetchKatanaEventsPaginated(endpoint) {
  var allItems = [];
  var page = 1;
  var hasMore = true;
  var maxItems = 500;

  while (hasMore && allItems.length < maxItems) {
    var separator = endpoint.indexOf('?') > -1 ? '&' : '?';
    var result = katanaApiCall(endpoint + separator + 'page=' + page);

    if (!result) {
      hasMore = false;
      break;
    }

    var items = result.data || result;
    if (!items || !items.length || items.length === 0) {
      hasMore = false;
      break;
    }

    for (var i = 0; i < items.length; i++) {
      allItems.push(items[i]);
    }

    // Stop if fewer than requested
    if (items.length < 50) {
      hasMore = false;
    } else {
      page++;
      Utilities.sleep(300);
    }
  }

  return allItems;
}
