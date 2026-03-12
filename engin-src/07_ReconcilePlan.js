/**
 * Safe reconciliation lane for engin-src.
 *
 * This does not modify the existing sync or push flows.
 * It builds a separate plan sheet from the current Wasp tab and only executes
 * rows the operator explicitly approves.
 */

var RECONCILE_PLAN_SHEET = 'Reconcile Plan';
var RECONCILE_PLAN_HEADERS = [
  'Run ID', 'SKU', 'Site', 'Location', 'Lot', 'DateCode',
  'Katana Qty', 'WASP Qty', 'Delta', 'Action',
  'Safety', 'Source Match', 'Approval', 'Result', 'Note'
];

function buildSafeReconcilePlan() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var pending = getReconcilePendingWarnings_(ss);
  if (pending.length > 0) {
    SpreadsheetApp.getUi().alert(
      'Reconcile plan blocked.\n\n' +
      'Clear pending edits first:\n- ' + pending.join('\n- ') + '\n\n' +
      'Run Sync Now again after the sheets are stable.'
    );
    return;
  }

  var waspSheet = ss.getSheetByName('Wasp');
  if (!waspSheet || waspSheet.getLastRow() < 2) {
    SpreadsheetApp.getUi().alert('Wasp sheet not found or empty. Run Sync Now first.');
    return;
  }

  var lastRow = waspSheet.getLastRow();
  var data = waspSheet.getRange(2, 1, lastRow - 1, 19).getValues();
  var runId = formatTimestamp();
  var planRows = [];
  var safeCount = 0;
  var reviewCount = 0;
  var infoCount = 0;
  var groups = {};

  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var sku = String(row[0] || '').trim();
    if (!sku || sku.indexOf(' > ') >= 0) continue;

    var site = String(row[3] || '').trim();
    var location = String(row[4] || '').trim();
    var lotTracked = String(row[5] || '').trim();
    var lot = String(row[6] || '').trim();
    var dateCode = reconcileFormatDate_(row[7]);
    var katanaQty = reconcileToNumber_(row[9]);
    var waspQty = reconcileToNumber_(row[11]);
    var match = String(row[14] || '').trim().toUpperCase();
    var key = sku + '|' + site;

    if (match === 'ADJUSTABLE') {
      if (!groups[key]) {
        groups[key] = {
          sku: sku,
          site: site,
          katanaQty: katanaQty,
          waspTotal: 0,
          rows: []
        };
      }
      groups[key].katanaQty = katanaQty;
      groups[key].waspTotal += waspQty;
      groups[key].rows.push({
        location: location,
        lot: lot,
        dateCode: dateCode,
        waspQty: waspQty
      });
      continue;
    }

    if (match === 'MISMATCH' || match === 'WASP ONLY') {
      planRows.push([
        runId,
        sku,
        site,
        location,
        lot,
        dateCode,
        katanaQty,
        waspQty,
        reconcileRound_(katanaQty - waspQty),
        '',
        'REVIEW',
        match,
        'REVIEW',
        '',
        buildReconcileReviewNote_(match, lotTracked, lot)
      ]);
      reviewCount++;
      continue;
    }

    if (match === 'SPLIT MATCH') {
      planRows.push([
        runId,
        sku,
        site,
        location,
        lot,
        dateCode,
        katanaQty,
        waspQty,
        0,
        '',
        'INFO',
        match,
        'SKIP',
        'NO ACTION',
        'Site total already matches; stock is only split across sub-locations.'
      ]);
      infoCount++;
    }
  }

  var groupKeys = Object.keys(groups);
  for (var g = 0; g < groupKeys.length; g++) {
    var grp = groups[groupKeys[g]];
    var delta = reconcileRound_(grp.katanaQty - grp.waspTotal);
    if (Math.abs(delta) < 0.01) continue;

    var rowsSorted = grp.rows.slice().sort(function(a, b) {
      return b.waspQty - a.waspQty;
    });
    var dominantRow = rowsSorted.length > 0 ? rowsSorted[0] : { location: 'UNSORTED', waspQty: 0 };

    if (delta > 0) {
      planRows.push([
        runId,
        grp.sku,
        grp.site,
        dominantRow.location,
        '',
        '',
        grp.katanaQty,
        grp.waspTotal,
        delta,
        'ADD',
        'SAFE',
        'ADJUSTABLE',
        'PENDING',
        '',
        'Safe site-level add based on ADJUSTABLE rows.'
      ]);
      safeCount++;
      continue;
    }

    var toRemove = Math.abs(delta);
    for (var r = 0; r < rowsSorted.length && toRemove > 0.01; r++) {
      var locRow = rowsSorted[r];
      if (locRow.waspQty <= 0) continue;
      var removeQty = reconcileRound_(Math.min(locRow.waspQty, toRemove));
      if (removeQty <= 0) continue;
      planRows.push([
        runId,
        grp.sku,
        grp.site,
        locRow.location,
        '',
        '',
        grp.katanaQty,
        grp.waspTotal,
        -removeQty,
        'REMOVE',
        'SAFE',
        'ADJUSTABLE',
        'PENDING',
        '',
        'Safe drain from highest-WASP-qty location.'
      ]);
      safeCount++;
      toRemove = reconcileRound_(toRemove - removeQty);
    }
  }

  var sheet = ensureReconcilePlanSheet_(ss);
  clearReconcilePlanSheet_(sheet);
  if (planRows.length > 0) {
    sheet.getRange(2, 1, planRows.length, RECONCILE_PLAN_HEADERS.length).setValues(planRows);
    applyReconcilePlanFormatting_(sheet, planRows.length);
  }

  var duration = 0;
  var summary = 'Safe=' + safeCount + ', Review=' + reviewCount + ', Info=' + infoCount;
  logSyncHistory(ss, 'RECONCILE', 'Plan built: ' + summary, 0, '', duration, 'SUCCESS');

  SpreadsheetApp.getUi().alert(
    'Reconcile plan built.\n\n' +
    'Safe rows: ' + safeCount + '\n' +
    'Review rows: ' + reviewCount + '\n' +
    'Info rows: ' + infoCount + '\n\n' +
    'Next step: review the "' + RECONCILE_PLAN_SHEET + '" tab, then run "Approve Safe Reconcile".'
  );
}

function approveSafeReconcileActions() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(RECONCILE_PLAN_SHEET);
  if (!sheet || sheet.getLastRow() < 2) {
    SpreadsheetApp.getUi().alert('No reconcile plan found. Build the plan first.');
    return;
  }

  var lastRow = sheet.getLastRow();
  var range = sheet.getRange(2, 1, lastRow - 1, RECONCILE_PLAN_HEADERS.length);
  var data = range.getValues();
  var updated = 0;

  for (var i = 0; i < data.length; i++) {
    var action = String(data[i][9] || '').trim();
    var safety = String(data[i][10] || '').trim().toUpperCase();
    var approval = String(data[i][12] || '').trim().toUpperCase();
    var result = String(data[i][13] || '').trim().toUpperCase();
    if (action && safety === 'SAFE' && approval !== 'APPROVED' && result !== 'DONE') {
      data[i][12] = 'APPROVED';
      updated++;
    }
  }

  range.setValues(data);
  SpreadsheetApp.getUi().alert('Approved ' + updated + ' safe reconcile actions.');
}

function executeSafeReconcilePlan() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(RECONCILE_PLAN_SHEET);
  if (!sheet || sheet.getLastRow() < 2) {
    SpreadsheetApp.getUi().alert('No reconcile plan found. Build the plan first.');
    return;
  }

  var lastRow = sheet.getLastRow();
  var range = sheet.getRange(2, 1, lastRow - 1, RECONCILE_PLAN_HEADERS.length);
  var data = range.getValues();
  var activeUser = '';
  try { activeUser = Session.getActiveUser().getEmail(); } catch (e) {}
  if (!activeUser) { try { activeUser = Session.getEffectiveUser().getEmail(); } catch (e2) {} }
  if (!activeUser) { activeUser = PropertiesService.getScriptProperties().getProperty('USER_NAME') || ''; }

  var successCount = 0;
  var errorCount = 0;
  var skippedCount = 0;
  var errorDetails = [];
  var startTime = new Date();

  for (var i = 0; i < data.length; i++) {
    var action = String(data[i][9] || '').trim().toUpperCase();
    var approval = String(data[i][12] || '').trim().toUpperCase();
    var result = String(data[i][13] || '').trim().toUpperCase();
    if (!action || approval !== 'APPROVED' || result === 'DONE') {
      skippedCount++;
      continue;
    }

    var sku = String(data[i][1] || '').trim();
    var site = String(data[i][2] || '').trim();
    var location = String(data[i][3] || '').trim();
    var lot = String(data[i][4] || '').trim();
    var dateCode = reconcileFormatDate_(data[i][5]);
    var delta = reconcileToNumber_(data[i][8]);
    var qty = Math.abs(delta);

    if (!sku || !site || !location || qty <= 0) {
      data[i][13] = 'ERROR';
      data[i][14] = 'Missing SKU/site/location/qty.';
      errorCount++;
      errorDetails.push('Invalid row ' + (i + 2));
      continue;
    }

    var actionObj = {
      type: action,
      sku: sku,
      qty: qty,
      site: site,
      location: location,
      lot: lot,
      dateCode: dateCode,
      notes: 'Reconcile Plan ' + Utilities.formatDate(new Date(), 'America/Vancouver', 'yyyy-MM-dd')
    };

    var execResult = executeAction_(actionObj);
    if (execResult.success) {
      data[i][13] = 'DONE';
      data[i][14] = 'Executed ' + formatTimestamp();
      successCount++;
      logAdjustment_(
        ss,
        'Reconcile Plan',
        action === 'ADD' ? 'Add' : 'Remove',
        activeUser,
        sku,
        '',
        site,
        location,
        lot,
        dateCode,
        delta,
        'OK'
      );
    } else {
      data[i][13] = 'ERROR';
      data[i][14] = String(execResult.error || 'Unknown error').substring(0, 200);
      errorCount++;
      errorDetails.push(sku + ': ' + data[i][14]);
    }
  }

  range.setValues(data);
  applyReconcilePlanFormatting_(sheet, data.length);

  var duration = Math.round((new Date() - startTime) / 1000);
  var status = errorCount === 0 ? 'SUCCESS' : successCount === 0 ? 'ERROR' : 'PARTIAL';
  var summary = 'Executed=' + successCount + ', Errors=' + errorCount + ', Skipped=' + skippedCount;
  logSyncHistory(ss, 'RECONCILE', summary, errorCount, errorDetails.join('; '), duration, status);

  SpreadsheetApp.getUi().alert(
    'Reconcile execution complete.\n\n' +
    'Executed: ' + successCount + '\n' +
    'Errors: ' + errorCount + '\n' +
    'Skipped: ' + skippedCount + '\n\n' +
    'Run Sync Now to refresh the Katana/Wasp tabs.'
  );
}

function getReconcilePendingWarnings_(ss) {
  var warnings = [];
  var waspSheet = ss.getSheetByName('Wasp');
  if (waspSheet && waspSheet.getLastRow() > 1) {
    var waspData = waspSheet.getRange(2, 1, waspSheet.getLastRow() - 1, 16).getValues();
    for (var i = 0; i < waspData.length; i++) {
      var sku = String(waspData[i][0] || '').trim();
      if (!sku || sku.indexOf(' > ') >= 0) continue;
      var status = String(waspData[i][15] || '').trim().toUpperCase();
      if (status === 'PENDING' || status === 'DELETE') {
        warnings.push('Wasp tab has pending sheet changes');
        break;
      }
    }
  }

  var katanaSheet = ss.getSheetByName('Katana');
  if (katanaSheet && katanaSheet.getLastRow() > 1) {
    var katanaData = katanaSheet.getRange(2, 1, katanaSheet.getLastRow() - 1, 14).getValues();
    for (var k = 0; k < katanaData.length; k++) {
      var kSku = String(katanaData[k][0] || '').trim();
      if (!kSku || kSku.indexOf(' > ') >= 0) continue;
      var waspStatus = String(katanaData[k][11] || '').trim().toUpperCase();
      var adjStatus = String(katanaData[k][13] || '').trim().toUpperCase();
      if (waspStatus === 'PUSH' || waspStatus === 'ERROR' || adjStatus === 'PENDING' || adjStatus === 'ERROR') {
        warnings.push('Katana tab has pending sync/adjustment changes');
        break;
      }
    }
  }

  return warnings;
}

function ensureReconcilePlanSheet_(ss) {
  var sheet = ss.getSheetByName(RECONCILE_PLAN_SHEET);
  if (!sheet) sheet = ss.insertSheet(RECONCILE_PLAN_SHEET);
  if (sheet.getRange(1, 1, 1, RECONCILE_PLAN_HEADERS.length).getValues()[0][0] !== RECONCILE_PLAN_HEADERS[0]) {
    sheet.clear();
    sheet.getRange(1, 1, 1, RECONCILE_PLAN_HEADERS.length).setValues([RECONCILE_PLAN_HEADERS]);
    sheet.getRange(1, 1, 1, RECONCILE_PLAN_HEADERS.length)
      .setBackground('#263238')
      .setFontColor('#ffffff')
      .setFontWeight('bold');
    sheet.setFrozenRows(1);
    var widths = [155, 120, 120, 140, 110, 110, 90, 90, 80, 90, 90, 120, 100, 100, 320];
    for (var i = 0; i < widths.length; i++) {
      sheet.setColumnWidth(i + 1, widths[i]);
    }
  }
  return sheet;
}

function clearReconcilePlanSheet_(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, Math.max(sheet.getLastColumn(), RECONCILE_PLAN_HEADERS.length)).clearContent().clearFormat();
  }
  sheet.getRange(1, 1, 1, RECONCILE_PLAN_HEADERS.length)
    .setBackground('#263238')
    .setFontColor('#ffffff')
    .setFontWeight('bold');
}

function applyReconcilePlanFormatting_(sheet, rowCount) {
  if (rowCount <= 0) return;
  var fullRange = sheet.getRange(2, 1, rowCount, RECONCILE_PLAN_HEADERS.length);
  var safetyRange = sheet.getRange(2, 11, rowCount, 1);
  var approvalRange = sheet.getRange(2, 13, rowCount, 1);
  var resultRange = sheet.getRange(2, 14, rowCount, 1);

  var rules = [];
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('SAFE').setBackground('#e8f5e9').setRanges([safetyRange]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('REVIEW').setBackground('#fff3e0').setRanges([safetyRange]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('INFO').setBackground('#e3f2fd').setRanges([safetyRange]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('APPROVED').setBackground('#c8e6c9').setRanges([approvalRange]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('PENDING').setBackground('#fff8e1').setRanges([approvalRange]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('REVIEW').setBackground('#ffe0b2').setRanges([approvalRange]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('DONE').setBackground('#c8e6c9').setRanges([resultRange]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('ERROR').setBackground('#ffcdd2').setRanges([resultRange]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('NO ACTION').setBackground('#e3f2fd').setRanges([resultRange]).build());
  sheet.setConditionalFormatRules(rules);

  var approvalRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['PENDING', 'APPROVED', 'REVIEW', 'SKIP'], true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, 13, rowCount, 1).setDataValidation(approvalRule);
  fullRange.setVerticalAlignment('middle');
}

function buildReconcileReviewNote_(match, lotTracked, lot) {
  if (match === 'WASP ONLY') return 'WASP-only stock remains untouched by design.';
  if (lotTracked === 'Yes' || lot) return 'Lot-tracked mismatch requires manual review.';
  return 'Mismatch is not auto-safe; review UOM, routing, or location assumptions.';
}

function reconcileToNumber_(value) {
  if (value instanceof Date) return 0;
  return parseFloat(value) || 0;
}

function reconcileRound_(value) {
  return Math.round((parseFloat(value) || 0) * 100) / 100;
}

function reconcileFormatDate_(value) {
  if (!value) return '';
  if (value instanceof Date) {
    return Utilities.formatDate(value, 'America/Vancouver', 'yyyy-MM-dd');
  }
  var text = String(value).trim();
  if (text.indexOf('T') >= 0) return text.split('T')[0];
  return text;
}
