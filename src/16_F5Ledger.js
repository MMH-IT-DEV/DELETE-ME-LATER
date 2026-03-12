// ============================================
// 16_F5Ledger.gs - F5 SHIPMENT LINE LEDGER
// ============================================
// Durable line-level ledger for F5 shipments.
// Writes "Pending" before WASP remove calls, then updates each line to
// "Deducted" or "Failed" so outage recovery can use recorded state.
// ============================================

var F5_LEDGER_CONFIG = {
  SHEET_NAME: 'F5 Shipment Ledger'
};

function isF5LedgerHandledStatus_(status) {
  status = String(status || '').trim();
  return status === 'Deducted' || status === 'Voided' || status === 'Manual Recovered';
}

function isF5LedgerVoidEligibleStatus_(status) {
  status = String(status || '').trim();
  return status === 'Deducted' || status === 'Manual Recovered' || status === 'Void Failed';
}

function indexF5ShipmentLedgerRowsByLineKey_(sheet) {
  sheet = sheet || ensureF5LedgerSheet_();
  var index = {};
  if (!sheet || sheet.getLastRow() < 2) return index;

  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 14).getValues();
  for (var i = 0; i < data.length; i++) {
    var lineKey = String(data[i][0] || '').trim();
    if (!lineKey) continue;
    index[lineKey] = {
      rowNumber: i + 2,
      shipmentId: String(data[i][1] || '').trim(),
      orderKey: normalizeF5RecoveryOrder_(data[i][2]),
      shipDate: formatF5RecoveryDate_(data[i][3]),
      sourceSku: String(data[i][5] || '').trim(),
      sku: String(data[i][6] || '').trim(),
      qty: Number(data[i][7] || 0) || 0,
      bundleSku: String(data[i][8] || '').trim(),
      status: String(data[i][9] || '').trim(),
      evidence: String(data[i][10] || '').trim(),
      execId: String(data[i][11] || '').trim(),
      updatedAt: String(data[i][12] || '').trim(),
      note: String(data[i][13] || '').trim()
    };
  }

  return index;
}

function hydrateF5ShipmentLedgerLines_(linePlans) {
  linePlans = linePlans || [];
  if (!linePlans.length) return linePlans;

  var existingByLineKey = indexF5ShipmentLedgerRowsByLineKey_();
  for (var i = 0; i < linePlans.length; i++) {
    var line = linePlans[i];
    var existing = existingByLineKey[String(line.lineKey || '').trim()];
    if (!existing) continue;
    line.ledgerRow = existing.rowNumber;
    line.previousLedgerStatus = existing.status;
    line.ledgerStatus = existing.status;
  }

  return linePlans;
}

function getF5ShipmentLedgerEntryByShipmentId_(shipmentId) {
  shipmentId = String(shipmentId || '').trim();
  if (!shipmentId) return null;
  var state = loadF5ShipmentLedgerState_(null, '');
  return state.byShipmentId[shipmentId] || null;
}

function ensureF5LedgerSheet_(ss) {
  ss = ss || getDebugSpreadsheet_();
  var sheet = ss.getSheetByName(F5_LEDGER_CONFIG.SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(F5_LEDGER_CONFIG.SHEET_NAME);
    sheet.getRange(1, 1, 1, 14).setValues([[
      'Line Key', 'Shipment ID', 'Order #', 'Ship Date', 'Ship To', 'Source SKU',
      'SKU', 'Qty', 'Bundle SKU', 'Status', 'Evidence', 'Exec ID', 'Updated At', 'Note'
    ]]);
    sheet.hideSheet();
  }
  return sheet;
}

function buildF5ShipmentLedgerLines_(shipment, shipDateText) {
  var shipmentId = String(shipment && shipment.shipmentId || '').trim();
  var orderKey = normalizeF5RecoveryOrder_(shipment && shipment.orderNumber);
  var shipTo = typeof getF5RecoveryShipToName_ === 'function'
    ? getF5RecoveryShipToName_(shipment)
    : (((shipment && shipment.shipTo) || {}).name || '');
  var shipmentItems = shipment && (shipment.shipmentItems || shipment.items || shipment.orderItems || []);
  var lines = [];

  for (var i = 0; i < shipmentItems.length; i++) {
    var item = shipmentItems[i] || {};
    var sourceSku = String(item.sku || item.SKU || '').trim();
    var quantity = Number(item.quantity || item.Quantity || 0) || 0;
    if (!sourceSku || quantity <= 0) continue;

    if (BUNDLE_MAP[sourceSku]) {
      var components = BUNDLE_MAP[sourceSku];
      for (var b = 0; b < components.length; b++) {
        lines.push({
          lineKey: (shipmentId || (orderKey + '|' + shipDateText)) + '|' + i + '|' + b,
          shipmentId: shipmentId,
          orderKey: orderKey,
          shipDate: shipDateText,
          shipTo: shipTo,
          sourceSku: sourceSku,
          sku: components[b].sku,
          qty: (Number(components[b].qty) || 0) * quantity,
          bundleSku: sourceSku,
          note: 'Bundle expansion'
        });
      }
    } else {
      lines.push({
        lineKey: (shipmentId || (orderKey + '|' + shipDateText)) + '|' + i + '|0',
        shipmentId: shipmentId,
        orderKey: orderKey,
        shipDate: shipDateText,
        shipTo: shipTo,
        sourceSku: sourceSku,
        sku: sourceSku,
        qty: quantity,
        bundleSku: '',
        note: ''
      });
    }
  }

  return lines;
}

function registerF5ShipmentLedgerPending_(shipment, execId, linePlans) {
  linePlans = linePlans || [];
  if (!linePlans.length) return linePlans;

  var sheet = ensureF5LedgerSheet_();
  var nowText = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  var useRetryGuard = !!(typeof HOTFIX_FLAGS !== 'undefined' && HOTFIX_FLAGS.F5_LEDGER_RETRY_GUARD);
  var existingByLineKey = useRetryGuard ? indexF5ShipmentLedgerRowsByLineKey_(sheet) : {};
  var startRow = sheet.getLastRow() + 1;
  var values = [];
  var pendingUpdates = [];

  for (var i = 0; i < linePlans.length; i++) {
    var line = linePlans[i];
    var existing = existingByLineKey[String(line.lineKey || '').trim()];

    if (existing) {
      line.ledgerRow = existing.rowNumber;
      line.previousLedgerStatus = existing.status;
      line.ledgerStatus = existing.status;

      if (!isF5LedgerHandledStatus_(existing.status)) {
        pendingUpdates.push({
          rowNumber: existing.rowNumber,
          values: [[
            'Pending',
            'ShipStation line queued for WASP remove',
            execId || '',
            nowText,
            line.note || existing.note || ''
          ]]
        });
        line.ledgerStatus = 'Pending';
      }
      continue;
    }

    values.push([
      line.lineKey,
      line.shipmentId || '',
      line.orderKey ? ('#' + line.orderKey) : '',
      line.shipDate || '',
      line.shipTo || '',
      line.sourceSku || '',
      line.sku || '',
      Number(line.qty) || 0,
      line.bundleSku || '',
      'Pending',
      'ShipStation line queued for WASP remove',
      execId || '',
      nowText,
      line.note || ''
    ]);
    line.ledgerRow = startRow + values.length - 1;
    line.previousLedgerStatus = '';
    line.ledgerStatus = 'Pending';
  }

  if (values.length > 0) {
    sheet.getRange(startRow, 1, values.length, 14).setValues(values);
  }

  for (var u = 0; u < pendingUpdates.length; u++) {
    sheet.getRange(pendingUpdates[u].rowNumber, 10, 1, 5).setValues(pendingUpdates[u].values);
  }

  return linePlans;
}

function applyF5ShipmentLedgerResults_(linePlans, itemResults, execId) {
  linePlans = linePlans || [];
  itemResults = itemResults || [];
  if (!linePlans.length) return;

  var sheet = ensureF5LedgerSheet_();
  var nowText = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  var updates = [];

  for (var i = 0; i < linePlans.length; i++) {
    var line = linePlans[i] || {};
    var result = itemResults[i] || {};
    var status;
    var evidence;
    var noteParts = [];

    if (result.preserveLedgerStatus && line.previousLedgerStatus) {
      status = line.previousLedgerStatus;
      evidence = 'Ledger state preserved on retry';
    } else {
      status = result.success ? 'Deducted' : 'Failed';
      evidence = result.success ? 'WASP remove success' : 'WASP remove failed';
    }

    if (result.action) noteParts.push(result.action);
    if (result.error) noteParts.push(result.error);
    if (line.note) noteParts.push(line.note);

    updates.push([
      status,
      evidence,
      execId || '',
      nowText,
      noteParts.join(' | ')
    ]);
  }

  var startRow = linePlans[0].ledgerRow;
  var isContiguous = true;
  for (var j = 1; j < linePlans.length; j++) {
    if (linePlans[j].ledgerRow !== startRow + j) {
      isContiguous = false;
      break;
    }
  }

  if (isContiguous) {
    sheet.getRange(startRow, 10, updates.length, 5).setValues(updates);
    return;
  }

  for (var k = 0; k < linePlans.length; k++) {
    if (!linePlans[k].ledgerRow) continue;
    sheet.getRange(linePlans[k].ledgerRow, 10, 1, 5).setValues([updates[k]]);
  }
}

function applyF5ShipmentLedgerVoidResults_(linePlans, itemResults, execId) {
  linePlans = linePlans || [];
  itemResults = itemResults || [];
  if (!linePlans.length) return;

  var sheet = ensureF5LedgerSheet_();
  var nowText = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');

  for (var i = 0; i < linePlans.length; i++) {
    var line = linePlans[i] || {};
    var result = itemResults[i] || {};
    if (!line.ledgerRow) continue;

    var status;
    var evidence;
    if (result.preserveLedgerStatus && line.previousLedgerStatus) {
      status = line.previousLedgerStatus;
      evidence = 'Ledger state preserved on void retry';
    } else {
      status = result.success ? 'Voided' : 'Void Failed';
      evidence = result.success ? 'WASP add success (void)' : 'WASP add failed (void)';
    }

    var noteParts = [];
    if (result.action) noteParts.push(result.action);
    if (result.error) noteParts.push(result.error);
    if (line.note) noteParts.push(line.note);

    sheet.getRange(line.ledgerRow, 10, 1, 5).setValues([[
      status,
      evidence,
      execId || '',
      nowText,
      noteParts.join(' | ')
    ]]);
  }
}

function loadF5ShipmentLedgerState_(ss, reportDate) {
  ss = ss || getDebugSpreadsheet_();
  var sheet = ss.getSheetByName(F5_LEDGER_CONFIG.SHEET_NAME);
  var state = {
    byShipmentId: {},
    byOrderDate: {}
  };
  if (!sheet || sheet.getLastRow() < 2) return state;

  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 14).getValues();
  var byLineKey = {};

  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var lineKey = String(row[0] || '').trim();
    var shipmentId = String(row[1] || '').trim();
    var orderKey = normalizeF5RecoveryOrder_(row[2]);
    var shipDate = formatF5RecoveryDate_(row[3]);
    if (!lineKey || !shipmentId) continue;
    if (reportDate && shipDate !== reportDate) continue;

    byLineKey[lineKey] = {
      lineKey: lineKey,
      shipmentId: shipmentId,
      orderKey: orderKey,
      shipDate: shipDate,
      shipTo: String(row[4] || '').trim(),
      sourceSku: String(row[5] || '').trim(),
      sku: String(row[6] || '').trim(),
      qty: Number(row[7] || 0) || 0,
      bundleSku: String(row[8] || '').trim(),
      status: String(row[9] || '').trim(),
      evidence: String(row[10] || '').trim(),
      execId: String(row[11] || '').trim(),
      updatedAt: String(row[12] || '').trim(),
      note: String(row[13] || '').trim(),
      rowNumber: i + 2
    };
  }

  var lineKeys = Object.keys(byLineKey);
  for (var j = 0; j < lineKeys.length; j++) {
    var line = byLineKey[lineKeys[j]];
    if (!state.byShipmentId[line.shipmentId]) {
      state.byShipmentId[line.shipmentId] = {
        shipmentId: line.shipmentId,
        orderKey: line.orderKey,
        shipDate: line.shipDate,
        shipTo: line.shipTo,
        lines: []
      };
    }
    state.byShipmentId[line.shipmentId].lines.push(line);

    if (line.orderKey && line.shipDate) {
      var orderDateKey = line.orderKey + '|' + line.shipDate;
      if (!state.byOrderDate[orderDateKey]) {
        state.byOrderDate[orderDateKey] = [];
      }
      state.byOrderDate[orderDateKey].push(state.byShipmentId[line.shipmentId]);
    }
  }

  return state;
}

function classifyF5ShipmentLedgerEntry_(entry) {
  if (!entry || !entry.lines || !entry.lines.length) {
    return { mode: 'none' };
  }

  var failed = [];
  var pending = [];
  var unknown = [];
  var handledCount = 0;

  for (var i = 0; i < entry.lines.length; i++) {
    var line = entry.lines[i];
    var status = String(line.status || '');
    if (status === 'Failed') {
      failed.push(line);
    } else if (status === 'Pending' || status === 'Void Failed') {
      pending.push(line);
    } else if (status === 'Deducted' || status === 'Voided' || status === 'Manual Recovered') {
      handledCount++;
    } else {
      unknown.push(line);
    }
  }

  if (pending.length > 0 || unknown.length > 0) {
    return {
      mode: 'review',
      reason: 'Ledger has pending or unknown lines',
      items: summarizeF5ShipmentLedgerLines_(entry.lines, 220)
    };
  }

  if (failed.length > 0) {
    return {
      mode: 'manual',
      lines: failed
    };
  }

  if (handledCount === entry.lines.length) {
    return { mode: 'handled' };
  }

  return {
    mode: 'review',
    reason: 'Ledger status mix requires review',
    items: summarizeF5ShipmentLedgerLines_(entry.lines, 220)
  };
}

function summarizeF5ShipmentLedgerLines_(lines, maxLen) {
  lines = lines || [];
  maxLen = maxLen || 200;
  var parts = [];

  for (var i = 0; i < lines.length; i++) {
    var line = lines[i];
    parts.push(
      (line.sku || '?') + ' x' + (Number(line.qty) || 0) +
      (line.status ? ' [' + line.status + ']' : '')
    );
  }

  var text = parts.join(', ');
  if (text.length > maxLen) return text.substring(0, maxLen - 3) + '...';
  return text;
}

function markF5ShipmentLedgerManualRecovered_(shipmentId, orderKey, reportDate, note) {
  var sheet = ensureF5LedgerSheet_();
  if (sheet.getLastRow() < 2) return 0;

  shipmentId = String(shipmentId || '').trim();
  orderKey = normalizeF5RecoveryOrder_(orderKey);
  reportDate = String(reportDate || '').trim();
  note = String(note || 'Manual recovery finalized');

  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 14).getValues();
  var nowText = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  var updated = 0;

  for (var i = 0; i < data.length; i++) {
    var rowNumber = i + 2;
    var rowShipmentId = String(data[i][1] || '').trim();
    var rowOrderKey = normalizeF5RecoveryOrder_(data[i][2]);
    var rowShipDate = formatF5RecoveryDate_(data[i][3]);
    var rowStatus = String(data[i][9] || '').trim();

    if (reportDate && rowShipDate !== reportDate) continue;
    if (shipmentId) {
      if (rowShipmentId !== shipmentId) continue;
    } else if (orderKey) {
      if (rowOrderKey !== orderKey) continue;
    } else {
      continue;
    }

    if (rowStatus === 'Deducted' || rowStatus === 'Voided' || rowStatus === 'Manual Recovered') {
      continue;
    }

    sheet.getRange(rowNumber, 10, 1, 5).setValues([[
      'Manual Recovered',
      'Manual WASP deduction confirmed',
      'F5 Recovery Finalize',
      nowText,
      note
    ]]);
    updated++;
  }

  return updated;
}
