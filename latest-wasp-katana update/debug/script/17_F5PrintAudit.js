// ============================================
// 17_F5PrintAudit.gs - F5 PRINT / SHIPMENT AUDIT
// ============================================
// Builds a read-only audit sheet for today's ShipStation shipments so ops can
// see what made it through F5, what needs manual recovery, and what still
// needs review.
// ============================================

var F5_PRINT_AUDIT_CONFIG = {
  SHEET_NAME: 'F5 Print Audit'
};

function buildF5PrintAuditToday() {
  return buildF5PrintAuditForDate_(new Date());
}

function buildF5PrintAuditYesterday() {
  var d = new Date();
  d.setDate(d.getDate() - 1);
  return buildF5PrintAuditForDate_(d);
}

function buildF5PrintAuditForDate_(dateObj) {
  var ss = getDebugSpreadsheet_();
  var reportDate = Utilities.formatDate(dateObj, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  var f5Map = getF5RecoveryOrderStatusMap_(ss, reportDate);
  var shipments = fetchShipStationShipmentsForDate_(reportDate);
  var shipmentLookupCache = {};
  shipments = buildF5PrintAuditShipmentSet_(shipments, f5Map, reportDate, shipmentLookupCache);
  var ledgerState = typeof loadF5ShipmentLedgerState_ === 'function'
    ? loadF5ShipmentLedgerState_(ss, reportDate)
    : { byShipmentId: {}, byOrderDate: {} };
  var recoveryState = loadF5ManualRecoveryState_(ss);
  var cache = CacheService.getScriptCache();
  var shipmentCounts = getF5RecoveryShipmentCounts_(shipments);

  var rows = [];
  var summary = {
    reportDate: reportDate,
    totalShipments: 0,
    handled: 0,
    actionNeeded: 0,
    review: 0,
    missing: 0,
    manualRecovered: 0
  };

  for (var i = 0; i < shipments.length; i++) {
    var shipment = shipments[i];
    if (shipment.voidDate) continue;

    var orderKey = normalizeF5RecoveryOrder_(shipment.orderNumber);
    var shipmentId = String(shipment.shipmentId || '').trim();
    var f5Info = f5Map[orderKey] || null;
    var ledgerEntry = ledgerState.byShipmentId[shipmentId] || null;
    var ledgerClass = (ledgerEntry && typeof classifyF5ShipmentLedgerEntry_ === 'function')
      ? classifyF5ShipmentLedgerEntry_(ledgerEntry)
      : { mode: 'none' };
    var recovery = findF5ManualRecoveryRecord_(recoveryState, shipmentId, orderKey, reportDate);
    var cacheHit = cache.get('ss_shipped_' + shipmentId) ? 'Yes' : '';
    var expandedItems = expandF5RecoveryItems_(shipment.shipmentItems || shipment.items || shipment.orderItems || []);
    var itemsText = summarizeF5RecoveryItems_(expandedItems, 180);
    var shipmentSpecificMissing = [];
    var action;
    var issue = '';
    var ledgerText = '';
    var f5Status = f5Info ? String(f5Info.status || '') : '';
    var multiShipment = (shipmentCounts[orderKey] || 0) > 1;

    if (f5Info && (f5Status === 'Partial' || f5Status === 'Failed')) {
      shipmentSpecificMissing = deriveF5RecoveryMissingItemsFromShipment_(shipment, f5Info);
      if (!shipmentSpecificMissing.length) {
        shipmentSpecificMissing = filterF5RecoveryFailedItemsToShipment_(f5Info.failedItems || [], shipment);
      }
    }

    if (ledgerClass.mode === 'manual') {
      ledgerText = summarizeF5ShipmentLedgerLines_(ledgerClass.lines || [], 160);
    } else if (ledgerEntry && ledgerEntry.lines) {
      ledgerText = summarizeF5ShipmentLedgerLines_(ledgerEntry.lines || [], 160);
    }

    if (recovery) {
      action = 'Manual Recovered';
      issue = recovery.note || 'Already manually recovered';
      summary.manualRecovered++;
    } else if (ledgerClass.mode === 'handled') {
      action = 'Handled';
      issue = 'Ledger shows shipment already handled';
      summary.handled++;
    } else if (ledgerClass.mode === 'manual') {
      action = 'Manual Recovery Needed';
      issue = summarizeF5RecoveryItems_(shipmentSpecificMissing.length ? shipmentSpecificMissing : (ledgerClass.lines || []), 180);
      summary.actionNeeded++;
    } else if (ledgerClass.mode === 'review') {
      action = 'Review';
      issue = ledgerClass.reason || 'Ledger requires review';
      summary.review++;
    } else if (f5Status === 'Shipped' || f5Status === 'Voided' || f5Status === 'Returned') {
      action = 'Handled';
      issue = 'Visible F5 row shows ' + f5Status;
      summary.handled++;
    } else if (f5Status === 'Partial' || f5Status === 'Failed') {
      if (shipmentSpecificMissing.length > 0) {
        action = 'Manual Recovery Needed';
        issue = summarizeF5RecoveryItems_(shipmentSpecificMissing, 180);
        summary.actionNeeded++;
      } else {
        action = 'Review';
        issue = 'F5 header is ' + f5Status + ' but shipment-specific missing items were not resolved';
        summary.review++;
      }
    } else if (cacheHit) {
      action = 'Review';
      issue = 'Dedup cache hit but no visible F5 or ledger result';
      summary.review++;
    } else {
      action = 'Missing from F5';
      issue = 'ShipStation shipment has no matching F5 result yet';
      summary.missing++;
    }

    if (multiShipment) {
      issue = issue ? (issue + ' | Multi-shipment order') : 'Multi-shipment order';
    }

    rows.push([
      '#' + orderKey,
      shipmentId,
      getF5PrintAuditDisplayDateTime_(shipment, f5Info, reportDate),
      getF5RecoveryShipToName_(shipment),
      itemsText,
      f5Status || '',
      ledgerClass.mode === 'none' ? '' : ledgerClass.mode,
      cacheHit,
      issue,
      action
    ]);
    summary.totalShipments++;
  }

  rows.sort(function(aRow, bRow) {
    var actionCmp = String(aRow[9] || '').localeCompare(String(bRow[9] || ''));
    if (actionCmp !== 0) return actionCmp;
    return String(aRow[0] || '').localeCompare(String(bRow[0] || ''));
  });

  writeF5PrintAuditSheet_(ss, summary, rows);
  return summary;
}

function buildF5PrintAuditShipmentSet_(shipments, f5Map, reportDate, shipmentLookupCache) {
  shipments = shipments || [];
  f5Map = f5Map || {};
  shipmentLookupCache = shipmentLookupCache || {};

  var results = [];
  var seen = {};
  var seenOrders = {};

  for (var i = 0; i < shipments.length; i++) {
    var shipment = shipments[i];
    var orderKey = normalizeF5RecoveryOrder_(shipment.orderNumber);
    var dedupKey = String(shipment.shipmentId || '').trim() || (orderKey + '|' + formatF5RecoveryDateTime_(shipment.shipDate));
    if (seen[dedupKey]) continue;
    seen[dedupKey] = true;
    if (orderKey) seenOrders[orderKey] = true;
    results.push(shipment);
  }

  var orderKeys = Object.keys(f5Map);
  for (var j = 0; j < orderKeys.length; j++) {
    var orderKey = orderKeys[j];
    if (!orderKey || seenOrders[orderKey]) continue;

    var lookupShipment = findF5RecoveryShipmentByOrder_(orderKey, reportDate, shipmentLookupCache);
    if (lookupShipment) {
      var lookupKey = String(lookupShipment.shipmentId || '').trim() || (orderKey + '|' + formatF5RecoveryDateTime_(lookupShipment.shipDate));
      if (!seen[lookupKey]) {
        seen[lookupKey] = true;
        seenOrders[orderKey] = true;
        results.push(lookupShipment);
      }
      continue;
    }

    results.push(buildF5PrintAuditSyntheticShipment_(orderKey, f5Map[orderKey]));
    seenOrders[orderKey] = true;
  }

  return results;
}

function buildF5PrintAuditSyntheticShipment_(orderKey, f5Info) {
  f5Info = f5Info || {};
  var items = [];
  var allItems = f5Info.allItems || [];

  for (var i = 0; i < allItems.length; i++) {
    items.push({
      sku: allItems[i].sku || '',
      quantity: Number(allItems[i].qty || 0) || 0
    });
  }

  return {
    orderNumber: orderKey,
    shipmentId: '',
    shipDate: f5Info.time || '',
    shipTo: { name: f5Info.shipTo || '' },
    shipmentItems: items,
    __auditSynthetic: true
  };
}

function getF5PrintAuditDisplayDateTime_(shipment, f5Info, reportDate) {
  if (f5Info && f5Info.time) {
    return String(f5Info.time).trim();
  }

  shipment = shipment || {};
  var rawValue = shipment.shipDate || shipment.createDate || shipment.shipmentDate || '';
  var formatted = formatF5RecoveryDateTime_(rawValue);
  var formattedDate = formatF5RecoveryDate_(rawValue);

  if (reportDate && formattedDate && formattedDate !== reportDate) {
    return reportDate + ' batch';
  }

  if (formatted) return formatted;
  return reportDate || '';
}

function getF5PrintAuditSheet_(ss, createIfMissing) {
  ss = ss || getDebugSpreadsheet_();
  var sheet = ss.getSheetByName(F5_PRINT_AUDIT_CONFIG.SHEET_NAME);
  if (sheet) return sheet;
  if (!createIfMissing) return null;
  return ss.insertSheet(F5_PRINT_AUDIT_CONFIG.SHEET_NAME);
}

function writeF5PrintAuditSheet_(ss, summary, rows) {
  var sheet = getF5PrintAuditSheet_(ss, true);
  var totalRows = Math.max(8 + Math.max(rows.length, 1), 12);

  sheet.getRange('A1:J4').breakApart();
  sheet.getRange(1, 1, Math.max(sheet.getMaxRows(), totalRows), 10).clearContent().clearFormat();

  sheet.getRange('A1:J1').merge();
  sheet.getRange('A1').setValue('F5 Print Audit');
  sheet.getRange('A1').setBackground('#d9edf7').setFontWeight('bold').setFontSize(14);

  sheet.getRange('A2:J3').setValues([
    ['Report Date', summary.reportDate, 'Generated', Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss'), 'Shipments', summary.totalShipments, 'Handled', summary.handled, 'Manual Recovered', summary.manualRecovered],
    ['Action Needed', summary.actionNeeded, 'Review', summary.review, 'Missing from F5', summary.missing, '', '', '', '']
  ]);
  sheet.getRange('A2:J3').setBackground('#eef7fb').setFontWeight('bold');

  sheet.getRange('A4:J4').merge();
  sheet.getRange('A4').setValue('Source = ShipStation shipments for the report date. Action shows whether F5 handled the shipment, missed it, or still needs manual recovery/review.');
  sheet.getRange('A4').setBackground('#f7f7f7');

  var headerRow = 6;
  var headers = [['Order #', 'Shipment ID', 'Ship Date', 'Ship To', 'ShipStation Items', 'F5 Status', 'Ledger', 'Cache', 'Issue', 'Action']];
  sheet.getRange(headerRow, 1, 1, 10).setValues(headers);
  sheet.getRange(headerRow, 1, 1, 10).setBackground('#fce5cd').setFontWeight('bold');

  if (rows.length > 0) {
    sheet.getRange(headerRow + 1, 1, rows.length, 10).setValues(rows);
  } else {
    sheet.getRange(headerRow + 1, 1, 1, 10).setValues([['No shipments found for this report date.', '', '', '', '', '', '', '', '', '']]);
  }

  sheet.setFrozenRows(headerRow);
  sheet.setColumnWidths(1, 10, 120);
  sheet.setColumnWidth(1, 90);
  sheet.setColumnWidth(2, 110);
  sheet.setColumnWidth(3, 140);
  sheet.setColumnWidth(4, 160);
  sheet.setColumnWidth(5, 260);
  sheet.setColumnWidth(6, 90);
  sheet.setColumnWidth(7, 90);
  sheet.setColumnWidth(8, 60);
  sheet.setColumnWidth(9, 320);
  sheet.setColumnWidth(10, 170);

  if (rows.length > 0) {
    var actionRange = sheet.getRange(headerRow + 1, 10, rows.length, 1);
    var actionValues = actionRange.getValues();
    for (var i = 0; i < actionValues.length; i++) {
      var action = String(actionValues[i][0] || '');
      var bg = '#ffffff';
      if (action === 'Handled' || action === 'Manual Recovered') bg = '#d4edda';
      else if (action === 'Manual Recovery Needed') bg = '#fff3cd';
      else if (action === 'Missing from F5') bg = '#f8d7da';
      else if (action === 'Review') bg = '#fce5cd';
      sheet.getRange(headerRow + 1 + i, 10).setBackground(bg);
    }
  }

  ss.setActiveSheet(sheet);
  return sheet;
}
