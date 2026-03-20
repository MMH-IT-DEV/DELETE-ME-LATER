// ============================================
// 15_F5Recovery.gs - F5 MANUAL RECOVERY
// ============================================
// Builds a sheet-based manual recovery tab for same-day F5 outages.
// Primary output: "F5 Manual Recovery" tab with manual WASP deduction lines.
// Finalize step stores manual recovery state so live F5/backfill skips them.
// ============================================

var F5_RECOVERY_CONFIG = {
  SHEET_NAME: 'F5 Manual Recovery',
  LEGACY_SHEET_NAMES: ['F5 Recovery'],
  STATE_SHEET_NAME: 'F5 Recovery State',
  WASP_SITE: 'MMH Kelowna',
  WASP_LOCATION: 'SHOPIFY',
  PAGE_SIZE: 100,
  INCLUDE_UOM: false,
  PAUSE_ACTIVE_PROP: 'F5_MANUAL_RECOVERY_ACTIVE',
  PAUSE_STARTED_AT_PROP: 'F5_MANUAL_RECOVERY_STARTED_AT',
  PAUSE_REPORT_DATE_PROP: 'F5_MANUAL_RECOVERY_REPORT_DATE',
  PAUSE_REASON_PROP: 'F5_MANUAL_RECOVERY_REASON',
  STALE_CLEANUP_ORDERS_PROP: 'F5_MANUAL_RECOVERY_STALE_CLEANUP_ORDERS',
  PAUSE_MAX_MINUTES: 120
};

function getDebugSpreadsheet_() {
  var active = SpreadsheetApp.getActiveSpreadsheet();
  if (active && active.getId && active.getId() === CONFIG.DEBUG_SHEET_ID) {
    return active;
  }
  return SpreadsheetApp.openById(CONFIG.DEBUG_SHEET_ID);
}

/**
 * Build a same-day F5 outage recovery report in the Debug Log sheet.
 * Safe/read-only: ShipStation read only, no WASP inventory calls.
 */
function buildF5RecoveryReportToday() {
  return buildF5RecoveryReportForDate_(new Date(), false);
}

/**
 * Optional helper for manual back-office recovery on the next morning.
 */
function buildF5RecoveryReportYesterday() {
  var d = new Date();
  d.setDate(d.getDate() - 1);
  return buildF5RecoveryReportForDate_(d, false);
}

function getF5ManualRecoveryModeInfo_() {
  var props = PropertiesService.getScriptProperties();
  var active = props.getProperty(F5_RECOVERY_CONFIG.PAUSE_ACTIVE_PROP) === '1';
  var startedAt = String(props.getProperty(F5_RECOVERY_CONFIG.PAUSE_STARTED_AT_PROP) || '').trim();
  var reportDate = String(props.getProperty(F5_RECOVERY_CONFIG.PAUSE_REPORT_DATE_PROP) || '').trim();
  var reason = String(props.getProperty(F5_RECOVERY_CONFIG.PAUSE_REASON_PROP) || '').trim();
  if (!active) {
    return { active: false, startedAt: '', reportDate: '', reason: '' };
  }

  var startedDate = startedAt ? new Date(startedAt) : null;
  var ageMinutes = 0;
  if (startedDate && !isNaN(startedDate.getTime())) {
    ageMinutes = Math.floor((Date.now() - startedDate.getTime()) / 60000);
    if (ageMinutes > F5_RECOVERY_CONFIG.PAUSE_MAX_MINUTES) {
      clearF5ManualRecoveryMode_('Expired stale manual recovery pause');
      return { active: false, startedAt: '', reportDate: '', reason: '', expired: true };
    }
  }

  return {
    active: true,
    startedAt: startedAt,
    reportDate: reportDate,
    reason: reason || 'F5 manual recovery active',
    ageMinutes: ageMinutes
  };
}

function activateF5ManualRecoveryMode_(reportDate, reason) {
  var props = PropertiesService.getScriptProperties();
  var nowIso = new Date().toISOString();
  props.setProperties({
    F5_MANUAL_RECOVERY_ACTIVE: '1',
    F5_MANUAL_RECOVERY_STARTED_AT: nowIso,
    F5_MANUAL_RECOVERY_REPORT_DATE: String(reportDate || '').trim(),
    F5_MANUAL_RECOVERY_REASON: String(reason || 'F5 manual recovery active').trim()
  }, false);
  logToSheet('SS_SHIP_POLL_PAUSE_ON', {
    reportDate: reportDate || '',
    startedAt: nowIso
  }, { reason: reason || 'F5 manual recovery active' });
  return getF5ManualRecoveryModeInfo_();
}

function clearF5ManualRecoveryMode_(reason) {
  var props = PropertiesService.getScriptProperties();
  var info = {
    reportDate: String(props.getProperty(F5_RECOVERY_CONFIG.PAUSE_REPORT_DATE_PROP) || '').trim(),
    startedAt: String(props.getProperty(F5_RECOVERY_CONFIG.PAUSE_STARTED_AT_PROP) || '').trim()
  };
  props.deleteProperty(F5_RECOVERY_CONFIG.PAUSE_ACTIVE_PROP);
  props.deleteProperty(F5_RECOVERY_CONFIG.PAUSE_STARTED_AT_PROP);
  props.deleteProperty(F5_RECOVERY_CONFIG.PAUSE_REPORT_DATE_PROP);
  props.deleteProperty(F5_RECOVERY_CONFIG.PAUSE_REASON_PROP);
  props.deleteProperty(F5_RECOVERY_CONFIG.STALE_CLEANUP_ORDERS_PROP);
  logToSheet('SS_SHIP_POLL_PAUSE_OFF', {
    reportDate: info.reportDate || '',
    startedAt: info.startedAt || ''
  }, { reason: reason || 'F5 manual recovery cleared' });
  return true;
}

function resumeF5ShippingPolling() {
  clearF5ManualRecoveryMode_('Manual resume from menu');
  return { success: true, resumed: true };
}

function saveF5RecoveryStaleCleanupOrders_(orderKeys) {
  var props = PropertiesService.getScriptProperties();
  orderKeys = orderKeys || [];
  if (!orderKeys.length) {
    props.deleteProperty(F5_RECOVERY_CONFIG.STALE_CLEANUP_ORDERS_PROP);
    return;
  }
  props.setProperty(
    F5_RECOVERY_CONFIG.STALE_CLEANUP_ORDERS_PROP,
    orderKeys.map(normalizeF5RecoveryOrder_).filter(Boolean).join(',')
  );
}

function loadF5RecoveryStaleCleanupOrders_() {
  var props = PropertiesService.getScriptProperties();
  var raw = String(props.getProperty(F5_RECOVERY_CONFIG.STALE_CLEANUP_ORDERS_PROP) || '').trim();
  if (!raw) return [];
  var parts = raw.split(',');
  var seen = {};
  var result = [];
  for (var i = 0; i < parts.length; i++) {
    var orderKey = normalizeF5RecoveryOrder_(parts[i]);
    if (!orderKey || seen[orderKey]) continue;
    seen[orderKey] = true;
    result.push(orderKey);
  }
  return result;
}

/**
 * Finalize the current recovery sheet after manual WASP deductions are done.
 * Stores shipment/order recovery state permanently so F5/backfill skips them later.
 */
function finalizeF5RecoveryReport() {
  var ss = getDebugSpreadsheet_();
  var sheet = getF5RecoverySheet_(ss, false);
  if (!sheet) throw new Error(F5_RECOVERY_CONFIG.SHEET_NAME + ' tab not found');

  var reportDate = String(sheet.getRange('B2').getValue() || '').trim();
  if (!reportDate) throw new Error('Report date missing from ' + F5_RECOVERY_CONFIG.SHEET_NAME + ' sheet');
  var hiddenReviewCount = getF5RecoveryHiddenReviewCount_(sheet);
  if (hiddenReviewCount > 0) {
    throw new Error(
      'Cannot finalize while Hidden Review = ' + hiddenReviewCount +
      '. Review those shipments first or keep F5 manual recovery mode active.'
    );
  }

  var rows = sheet.getDataRange().getValues();
  var currentSection = '';
  var saved = 0;
  var ledgerUpdated = 0;
  var visibleUpdated = 0;
  var staleCleanupUpdated = 0;
  var compactedRows = 0;
  var uniqueKeys = {};
  var manualLinesByOrder = {};

  for (var i = 0; i < rows.length; i++) {
    var row = rows[i];
    var colA = String(row[0] || '').trim();
    var colB = String(row[1] || '').trim();
    var colE = String(row[4] || '').trim();
    var colF = String(row[5] || '').trim();
    var colG = String(row[6] || '').trim();

    if (colA === 'SKU' && colB === 'Qty' && colE === 'Order #' && colF === 'Shipment ID' && colG === 'Ship Date') {
      currentSection = 'manual_lines';
      continue;
    }
    if (colA === 'Manual Lines To Deduct') {
      currentSection = '';
      continue;
    }
    if (currentSection !== 'manual_lines' || !colA || !(Number(row[1] || 0) > 0)) {
      continue;
    }

    var sku = String(row[0] || '').trim();
    var qty = Number(row[1] || 0) || 0;
    var orderKey = normalizeF5RecoveryOrder_(row[4]);
    var shipmentId = String(row[5] || '').trim();

    if (!manualLinesByOrder[orderKey]) {
      manualLinesByOrder[orderKey] = [];
    }
    if (sku && qty > 0) {
      manualLinesByOrder[orderKey].push({ sku: sku, qty: qty, shipmentId: shipmentId });
    }

    var dedupKey = (shipmentId || ('ORDER|' + orderKey + '|' + reportDate));
    if (uniqueKeys[dedupKey]) continue;
    uniqueKeys[dedupKey] = true;

    if (recordF5ManualRecovery_({
      shipmentId: shipmentId,
      orderKey: orderKey,
      reportDate: reportDate,
      source: 'F5 Recovery Finalize',
      note: 'Manual line recovered from ' + F5_RECOVERY_CONFIG.SHEET_NAME + ' tab'
    })) {
      saved++;
    }
    if (typeof markF5ShipmentLedgerManualRecovered_ === 'function') {
      ledgerUpdated += markF5ShipmentLedgerManualRecovered_(
        shipmentId,
        orderKey,
        reportDate,
        'Manual line recovered from ' + F5_RECOVERY_CONFIG.SHEET_NAME + ' tab'
      );
    }
  }

  if (typeof finalizeF5ManualRecoveryBatch === 'function') {
    var batchResult = finalizeF5ManualRecoveryBatch(manualLinesByOrder);
    if (batchResult && batchResult.success) {
      visibleUpdated = Number(batchResult.repairedOrders || 0) + Number(batchResult.loggedOrders || 0);
    }
  } else {
    var orderKeys = Object.keys(manualLinesByOrder);
    for (var k = 0; k < orderKeys.length; k++) {
      if (typeof finalizeF5ManualRecoveryRows !== 'function') break;
      var orderResult = finalizeF5ManualRecoveryRows(orderKeys[k], manualLinesByOrder[orderKeys[k]]);
      if (orderResult && orderResult.success) {
        visibleUpdated++;
      }
    }
  }

  var staleCleanupOrders = loadF5RecoveryStaleCleanupOrders_();
  if (staleCleanupOrders.length && typeof normalizeF5ManualRecoveryBatchOrders === 'function') {
    var cleanupResult = normalizeF5ManualRecoveryBatchOrders(staleCleanupOrders);
    if (cleanupResult && cleanupResult.success) {
      staleCleanupUpdated = Number(cleanupResult.normalizedOrders || 0);
      compactedRows += Number(cleanupResult.compactedRows || 0);
    }
  }
  if (typeof compactBlankF5ManualRecoverySheets_ === 'function') {
    var compactOnlyResult = compactBlankF5ManualRecoverySheets_();
    if (compactOnlyResult && compactOnlyResult.success) {
      compactedRows += Number(compactOnlyResult.compactedRows || 0);
    }
  }

  var finalizedAt = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  sheet.getRange('A4:J4').clearContent().merge();
  sheet.getRange('A4').setValue(
    'Finalized at ' + finalizedAt + ' | Stored ' + saved + ' shipment/order recovery marks, ' + ledgerUpdated + ' ledger line updates, ' + visibleUpdated + ' visible F5 row repairs, and ' + compactedRows + ' blank rows compacted. Safe to delete this tab.'
  );
  sheet.getRange('A4').setBackground('#d4edda').setFontWeight('bold');

  var activitySheet = ss.getSheetByName('Activity');
  if (activitySheet) {
    ss.setActiveSheet(activitySheet);
  }
  if (typeof ensureF5RecoveryStateSheet_ === 'function') {
    var stateSheet = ensureF5RecoveryStateSheet_(ss);
    if (stateSheet && !stateSheet.isSheetHidden()) stateSheet.hideSheet();
  }
  if (typeof ensureF5LedgerSheet_ === 'function') {
    var ledgerSheet = ensureF5LedgerSheet_(ss);
    if (ledgerSheet && !ledgerSheet.isSheetHidden()) ledgerSheet.hideSheet();
  }
  if (!sheet.isSheetHidden()) {
    sheet.hideSheet();
  }
  clearF5ManualRecoveryMode_('Manual recovery finalized for ' + reportDate);

  Logger.log('F5 recovery finalized | reportDate=' + reportDate + ' | saved=' + saved + ' | ledgerUpdated=' + ledgerUpdated + ' | visibleUpdated=' + visibleUpdated);
  return {
    success: true,
    reportDate: reportDate,
    saved: saved,
    ledgerUpdated: ledgerUpdated,
    visibleUpdated: visibleUpdated,
    staleCleanupUpdated: staleCleanupUpdated,
    compactedRows: compactedRows,
    recoverySheetHidden: true
  };
}

function buildF5RecoveryReportForDate_(dateObj, sendSlack) {
  var ss = getDebugSpreadsheet_();
  var reportDate = Utilities.formatDate(dateObj, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  activateF5ManualRecoveryMode_(reportDate, 'F5 manual recovery in progress');
  var shipments = fetchShipStationShipmentsForDate_(reportDate);
  var f5Map = getF5RecoveryOrderStatusMap_(ss, reportDate);
  var ledgerState = typeof loadF5ShipmentLedgerState_ === 'function'
    ? loadF5ShipmentLedgerState_(ss, reportDate)
    : { byShipmentId: {}, byOrderDate: {} };
  var state = loadF5ManualRecoveryState_(ss);
  var cache = CacheService.getScriptCache();
  var orderShipmentCounts = getF5RecoveryShipmentCounts_(shipments);
  var uomCache = {};
  var aggregateMap = {};
  var aggregateRows = [];
  var pendingF5ReviewMap = {};
  var shipmentLookupCache = {};
  var reviewCount = 0;
  var reviewOrderRows = [];
  var detailRows = [];
  var reviewSeen = {};
  var manualSeen = {};
  var recoveredSeen = {};
  var actionOrderMap = {};
  var staleCleanupMap = {};
  var handledCount = 0;
  var activeShipmentCount = 0;
  var voidedCount = 0;
  var manualRecoveredCount = 0;

  var f5OrderKeys = Object.keys(f5Map);
  for (var f = 0; f < f5OrderKeys.length; f++) {
    var f5OrderKey = f5OrderKeys[f];
    var f5Info = f5Map[f5OrderKey];
    var f5Status = String(f5Info.status || '');
    var f5Recovery = findF5ManualRecoveryRecord_(state, f5Info.shipmentId, f5OrderKey, reportDate);
    var ledgerOrderEntries = ledgerState.byOrderDate[f5OrderKey + '|' + reportDate] || [];

    if (f5Recovery) {
      var f5RecoveredKey = (f5Recovery.shipmentId || ('ORDER|' + f5OrderKey + '|' + reportDate));
      if (!recoveredSeen[f5RecoveredKey]) {
        recoveredSeen[f5RecoveredKey] = true;
        manualRecoveredCount++;
      }
      handledCount++;
      continue;
    }
    if (f5Info.hasShippedEntry && f5Info.hasProblemEntry) {
      staleCleanupMap[f5OrderKey] = true;
    } else if (Number(f5Info.headerCount || 0) > 1 && (Number(orderShipmentCounts[f5OrderKey] || 0) <= 1)) {
      staleCleanupMap[f5OrderKey] = true;
    }
    if (ledgerOrderEntries.length > 0) continue;
    if (f5Status !== 'Partial' && f5Status !== 'Failed') continue;

    pendingF5ReviewMap[f5OrderKey] = {
      orderKey: f5OrderKey,
      shipmentId: f5Info.shipmentId || '',
      shipDate: f5Info.time || '',
      f5Status: f5Status,
      f5Info: f5Info,
      reason: f5Info.failedItems.length > 0
        ? 'F5 ' + f5Status + ' row awaiting shipment validation'
        : 'F5 ' + f5Status + ' row but no failed child lines were parsed'
    };
  }

  for (var i = 0; i < shipments.length; i++) {
    var shipment = shipments[i];
    var orderKey = normalizeF5RecoveryOrder_(shipment.orderNumber);
    if (!orderKey) continue;

    if (shipment.voidDate) {
      voidedCount++;
      continue;
    }

    activeShipmentCount++;

    var existingRecovery = findF5ManualRecoveryRecord_(
      state,
      String(shipment.shipmentId || ''),
      orderKey,
      reportDate
    );
    if (existingRecovery) {
      var recoveredKey = (existingRecovery.shipmentId || ('ORDER|' + orderKey + '|' + reportDate));
      if (!recoveredSeen[recoveredKey]) {
        recoveredSeen[recoveredKey] = true;
        manualRecoveredCount++;
      }
      handledCount++;
      continue;
    }

    var f5Info = f5Map[orderKey] || null;
    var existingStatus = f5Info ? String(f5Info.status || '') : '';
    var cacheKey = 'ss_shipped_' + shipment.shipmentId;
    var cacheHit = cache.get(cacheKey) ? 'Yes' : '';
    var multiShipment = (orderShipmentCounts[orderKey] || 0) > 1;
    var ledgerEntry = ledgerState.byShipmentId[String(shipment.shipmentId || '').trim()] || null;

    if (ledgerEntry && typeof classifyF5ShipmentLedgerEntry_ === 'function') {
      var ledgerClass = classifyF5ShipmentLedgerEntry_(ledgerEntry);
      if (ledgerClass.mode === 'handled') {
        handledCount++;
        continue;
      }
      if (ledgerClass.mode === 'manual') {
        actionOrderMap[orderKey] = true;
        for (var li = 0; li < ledgerClass.lines.length; li++) {
          var ledgerLine = ledgerClass.lines[li];
          addF5RecoveryManualItem_(
            aggregateMap,
            detailRows,
            manualSeen,
            {
              reportDate: reportDate,
              orderKey: orderKey,
              shipmentId: shipment.shipmentId || '',
              shipDate: ledgerLine.shipDate || formatF5RecoveryDateTime_(shipment.shipDate),
              shipTo: ledgerEntry.shipTo || getF5RecoveryShipToName_(shipment),
              sku: ledgerLine.sku,
              qty: ledgerLine.qty,
              bundleSku: ledgerLine.bundleSku || '',
              cacheHit: '',
              source: 'Ledger failed line',
              note: ledgerLine.evidence || ledgerLine.note || 'F5 ledger line failed'
            },
            uomCache
          );
        }
        continue;
      }
      if (ledgerClass.mode === 'review') {
        if (!reviewSeen[orderKey]) {
          reviewSeen[orderKey] = true;
          reviewCount++;
          addF5RecoveryReviewRow_(reviewOrderRows, {
            orderKey: orderKey,
            shipmentId: shipment.shipmentId || '',
            shipDate: formatF5RecoveryDateTime_(shipment.shipDate),
            reason: ledgerClass.reason || 'Ledger requires review'
          });
        }
        continue;
      }
    }

    if (existingStatus === 'Shipped' || existingStatus === 'Voided' || existingStatus === 'Returned') {
      handledCount++;
      continue;
    }

    if (existingStatus === 'Partial' || existingStatus === 'Failed') {
      if (actionOrderMap[orderKey]) continue;
      var inferredMissing = deriveF5RecoveryMissingItemsFromShipment_(shipment, f5Info);
      if (!inferredMissing.length) {
        inferredMissing = filterF5RecoveryFailedItemsToShipment_(f5Info.failedItems || [], shipment);
      }
      if (inferredMissing.length > 0) {
        actionOrderMap[orderKey] = true;
        delete pendingF5ReviewMap[orderKey];
        for (var im = 0; im < inferredMissing.length; im++) {
          addF5RecoveryManualItem_(
            aggregateMap,
            detailRows,
            manualSeen,
            {
              reportDate: reportDate,
              orderKey: orderKey,
              shipmentId: shipment.shipmentId || '',
              shipDate: f5Info.time || formatF5RecoveryDateTime_(shipment.shipDate),
              shipTo: getF5RecoveryShipToName_(shipment),
              sku: inferredMissing[im].sku,
              qty: inferredMissing[im].qty,
              bundleSku: inferredMissing[im].bundleSku || '',
              cacheHit: '',
              source: inferredMissing[im].source || 'F5 inferred missing line',
              note: inferredMissing[im].note || 'Derived from ShipStation shipment minus deducted F5 rows'
            },
            uomCache
          );
        }
        continue;
      }
      if (!reviewSeen[orderKey]) {
        reviewSeen[orderKey] = true;
        reviewCount++;
        addF5RecoveryReviewRow_(reviewOrderRows, {
          orderKey: orderKey,
          shipmentId: shipment.shipmentId || '',
          shipDate: formatF5RecoveryDateTime_(shipment.shipDate),
          reason: (pendingF5ReviewMap[orderKey] && pendingF5ReviewMap[orderKey].reason) || 'Existing F5 row needs review'
        });
        delete pendingF5ReviewMap[orderKey];
      }
      continue;
    }

    if (multiShipment) {
      if (!reviewSeen[orderKey]) {
        reviewSeen[orderKey] = true;
        reviewCount++;
        addF5RecoveryReviewRow_(reviewOrderRows, {
          orderKey: orderKey,
          shipmentId: shipment.shipmentId || '',
          shipDate: formatF5RecoveryDateTime_(shipment.shipDate),
          reason: 'Multiple ShipStation shipments for same order'
        });
      }
      continue;
    }

    if (cacheHit) {
      if (!reviewSeen[orderKey]) {
        reviewSeen[orderKey] = true;
        reviewCount++;
        addF5RecoveryReviewRow_(reviewOrderRows, {
          orderKey: orderKey,
          shipmentId: shipment.shipmentId || '',
          shipDate: formatF5RecoveryDateTime_(shipment.shipDate),
          reason: 'Cache hit but no visible F5 row'
        });
      }
      continue;
    }

    var expanded = expandF5RecoveryItems_(shipment.shipmentItems || shipment.items || shipment.orderItems || []);
    actionOrderMap[orderKey] = true;
    for (var j = 0; j < expanded.length; j++) {
      addF5RecoveryManualItem_(
        aggregateMap,
        detailRows,
        manualSeen,
        {
          reportDate: reportDate,
          orderKey: orderKey,
          shipmentId: shipment.shipmentId || '',
          shipDate: formatF5RecoveryDateTime_(shipment.shipDate),
          shipTo: getF5RecoveryShipToName_(shipment),
          sku: expanded[j].sku,
          qty: expanded[j].qty,
          bundleSku: expanded[j].bundleSku || '',
          cacheHit: cacheHit,
          source: 'Missing shipment',
          note: 'Not found in F5 log'
        },
        uomCache
      );
    }
  }

  aggregateRows = buildF5RecoveryAggregateRows_(detailRows);
  detailRows.sort(function(aRow, bRow) {
    var cmp = String(aRow[0]).localeCompare(String(bRow[0]));
    if (cmp !== 0) return cmp;
    cmp = Number(bRow[1] || 0) - Number(aRow[1] || 0);
    if (cmp !== 0) return cmp;
    return String(aRow[4]).localeCompare(String(bRow[4]));
  });
  var pendingReviewKeys = Object.keys(pendingF5ReviewMap);
  for (var pr = 0; pr < pendingReviewKeys.length; pr++) {
    var pendingInfo = pendingF5ReviewMap[pendingReviewKeys[pr]];
    if (!pendingInfo || reviewSeen[pendingInfo.orderKey]) continue;
    var pendingF5Info = f5Map[pendingInfo.orderKey] || null;
    var fallbackShipment = pendingF5Info
      ? findF5RecoveryShipmentByOrder_(pendingInfo.orderKey, reportDate, shipmentLookupCache)
      : null;
    if (fallbackShipment && pendingF5Info) {
      var fallbackMissing = deriveF5RecoveryMissingItemsFromShipment_(fallbackShipment, pendingF5Info);
      if (!fallbackMissing.length) {
        fallbackMissing = filterF5RecoveryFailedItemsToShipment_(pendingF5Info.failedItems || [], fallbackShipment);
      }
      if (fallbackMissing.length > 0) {
        actionOrderMap[pendingInfo.orderKey] = true;
        for (var fm = 0; fm < fallbackMissing.length; fm++) {
          addF5RecoveryManualItem_(
            aggregateMap,
            detailRows,
            manualSeen,
            {
              reportDate: reportDate,
              orderKey: pendingInfo.orderKey,
              shipmentId: fallbackShipment.shipmentId || '',
              shipDate: pendingInfo.shipDate || formatF5RecoveryDateTime_(fallbackShipment.shipDate),
              shipTo: getF5RecoveryShipToName_(fallbackShipment),
              sku: fallbackMissing[fm].sku,
              qty: fallbackMissing[fm].qty,
              bundleSku: fallbackMissing[fm].bundleSku || '',
              cacheHit: '',
              source: fallbackMissing[fm].source || 'F5 inferred missing line',
              note: fallbackMissing[fm].note || 'Derived from order lookup shipment minus deducted F5 rows'
            },
            uomCache
          );
        }
        continue;
      }
    }
    if (pendingF5Info && pendingF5Info.hasShippedEntry) {
      handledCount++;
      continue;
    }
    if (pendingF5Info) {
      var directFailed = buildF5RecoveryDirectFailedItems_(pendingF5Info);
      if (directFailed.length > 0) {
        actionOrderMap[pendingInfo.orderKey] = true;
        for (var df = 0; df < directFailed.length; df++) {
          addF5RecoveryManualItem_(
            aggregateMap,
            detailRows,
            manualSeen,
            {
              reportDate: reportDate,
              orderKey: pendingInfo.orderKey,
              shipmentId: pendingInfo.shipmentId || '',
              shipDate: pendingInfo.shipDate || (pendingF5Info.time || ''),
              shipTo: pendingF5Info.shipTo || '',
              sku: directFailed[df].sku,
              qty: directFailed[df].qty,
              bundleSku: directFailed[df].bundleSku || '',
              cacheHit: '',
              source: directFailed[df].source || 'F5 failed line',
              note: directFailed[df].note || 'Derived directly from unresolved F5 failed lines'
            },
            uomCache
          );
        }
        continue;
      }
    }
    reviewSeen[pendingInfo.orderKey] = true;
    reviewCount++;
    addF5RecoveryReviewRow_(reviewOrderRows, pendingInfo);
  }
  reviewOrderRows.sort(function(aRow, bRow) {
    return String(aRow[0] || '').localeCompare(String(bRow[0] || ''));
  });

  var summary = {
    reportDate: reportDate,
    shipStationShipments: shipments.length,
    activeShipments: activeShipmentCount,
    voidedShipments: voidedCount,
    handledOrders: handledCount,
    manualRecovered: manualRecoveredCount,
    actionOrders: Object.keys(actionOrderMap).length,
    reviewOrders: reviewCount,
    manualSkuLines: detailRows.length,
    manualSkuCount: aggregateRows.length,
    staleCleanupOrders: Object.keys(staleCleanupMap).length
  };
  saveF5RecoveryStaleCleanupOrders_(Object.keys(staleCleanupMap));
  var sheet = writeF5RecoveryReport_(ss, summary, detailRows, reviewOrderRows);
  if (summary.actionOrders === 0 && summary.reviewOrders === 0 && summary.staleCleanupOrders === 0) {
    clearF5ManualRecoveryMode_('No F5 manual recovery action needed for ' + reportDate);
  }

  Logger.log(
    'F5 recovery report built for ' + reportDate +
    ' | active=' + activeShipmentCount +
    ' | handled=' + handledCount +
    ' | manualRecovered=' + manualRecoveredCount +
    ' | actionOrders=' + summary.actionOrders +
    ' | review=' + reviewCount +
    ' | manual lines=' + detailRows.length
  );

  return summary;
}

function getF5RecoverySheet_(ss, createIfMissing) {
  ss = ss || getDebugSpreadsheet_();
  var sheet = ss.getSheetByName(F5_RECOVERY_CONFIG.SHEET_NAME);
  if (sheet) return sheet;

  for (var i = 0; i < F5_RECOVERY_CONFIG.LEGACY_SHEET_NAMES.length; i++) {
    var legacySheet = ss.getSheetByName(F5_RECOVERY_CONFIG.LEGACY_SHEET_NAMES[i]);
    if (legacySheet) {
      legacySheet.setName(F5_RECOVERY_CONFIG.SHEET_NAME);
      return legacySheet;
    }
  }

  if (!createIfMissing) return null;
  return ss.insertSheet(F5_RECOVERY_CONFIG.SHEET_NAME);
}

function fetchShipStationShipmentsForDate_(dateText) {
  var allShipments = [];
  var page = 1;
  var pageSize = F5_RECOVERY_CONFIG.PAGE_SIZE;

  while (true) {
    var endpoint = '/shipments?shipDateStart=' + encodeURIComponent(dateText) +
      '&includeShipmentItems=true&pageSize=' + pageSize + '&page=' + page +
      '&sortBy=ShipDate&sortDir=ASC';
    var result = callShipStationAPI(endpoint, 'GET');

    if (result.code !== 200) {
      throw new Error('ShipStation API error: HTTP ' + result.code + ' on page ' + page);
    }

    var data = result.data || {};
    var batch = data.shipments || [];
    for (var i = 0; i < batch.length; i++) {
      var shipDateText = formatF5RecoveryDate_(batch[i].shipDate);
      if (shipDateText === dateText) {
        allShipments.push(batch[i]);
      }
    }

    var pages = data.pages || 1;
    if (page >= pages || batch.length < pageSize) break;
    page++;
  }

  return allShipments;
}

function findF5RecoveryShipmentByOrder_(orderKey, reportDate, cache) {
  orderKey = normalizeF5RecoveryOrder_(orderKey);
  if (!orderKey) return null;

  cache = cache || {};
  if (cache.hasOwnProperty(orderKey)) {
    return cache[orderKey];
  }

  var endpoint = '/shipments?orderNumber=' + encodeURIComponent(orderKey) +
    '&includeShipmentItems=true&pageSize=100&sortBy=ShipDate&sortDir=ASC';
  var result = callShipStationAPI(endpoint, 'GET');
  if (result.code !== 200) {
    cache[orderKey] = null;
    return null;
  }

  var shipments = (result.data && result.data.shipments) || [];
  var active = [];
  var sameDate = [];

  for (var i = 0; i < shipments.length; i++) {
    if (shipments[i].voidDate) continue;
    active.push(shipments[i]);
    if (!reportDate || formatF5RecoveryDate_(shipments[i].shipDate) === reportDate) {
      sameDate.push(shipments[i]);
    }
  }

  var chosen = null;
  if (sameDate.length === 1) {
    chosen = sameDate[0];
  } else if (sameDate.length > 1) {
    chosen = sameDate[sameDate.length - 1];
  } else if (active.length === 1) {
    chosen = active[0];
  }

  cache[orderKey] = chosen;
  return chosen;
}

function getF5RecoveryOrderStatusMap_(ss, reportDate) {
  var sheet = ss.getSheetByName('Activity');
  var map = {};
  if (!sheet || sheet.getLastRow() < 4) return map;

  var data = sheet.getRange(4, 1, sheet.getLastRow() - 3, 6).getDisplayValues();
  var current = null;

  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var execId = row[0];
    var timeText = String(row[1] || '');
    var flowText = String(row[2] || '');
    var detailText = String(row[3] || '');
    var status = String(row[4] || '');
    var errorText = String(row[5] || '');

    if (execId) {
      if (flowText !== 'F5 Shipping') {
        current = null;
        continue;
      }
      if (reportDate && formatF5RecoveryDate_(timeText) !== reportDate) {
        current = null;
        continue;
      }

      var orderMatch = detailText.match(/#\s*([A-Za-z0-9-]+)/);
      var orderKey = normalizeF5RecoveryOrder_(orderMatch ? orderMatch[1] : '');
      if (!orderKey) {
        current = null;
        continue;
      }

      var previous = map[orderKey] || null;
      current = {
        status: status,
        execId: execId,
        time: formatF5RecoveryDateTime_(timeText),
        shipTo: '',
        shipmentId: '',
        allItems: [],
        deductedItems: [],
        failedItems: [],
        headerError: errorText || '',
        headerCount: previous ? Number(previous.headerCount || 0) : 0,
        hasShippedEntry: previous ? !!previous.hasShippedEntry : false,
        hasProblemEntry: previous ? !!previous.hasProblemEntry : false
      };
      current.headerCount++;
      if (status === 'Shipped') current.hasShippedEntry = true;
      if (status === 'Partial' || status === 'Failed' || errorText) current.hasProblemEntry = true;
      map[orderKey] = current;
      continue;
    }

    if (!current || !detailText) continue;

    var parsedItem = parseF5RecoverySubItem_(detailText);
    if (!parsedItem) continue;

    current.allItems.push(parsedItem);
    if (status === 'Failed') {
      current.hasProblemEntry = true;
      current.failedItems.push(parsedItem);
    } else if (status === 'Deducted') {
      current.deductedItems.push(parsedItem);
    }
    if (errorText) current.hasProblemEntry = true;
  }

  return map;
}

function getF5RecoveryShipmentCounts_(shipments) {
  var counts = {};
  for (var i = 0; i < shipments.length; i++) {
    if (shipments[i].voidDate) continue;
    var orderKey = normalizeF5RecoveryOrder_(shipments[i].orderNumber);
    if (!orderKey) continue;
    counts[orderKey] = (counts[orderKey] || 0) + 1;
  }
  return counts;
}

function expandF5RecoveryItems_(shipmentItems) {
  var expanded = [];
  for (var i = 0; i < shipmentItems.length; i++) {
    var item = shipmentItems[i];
    var sku = item.sku || item.SKU || '';
    var qty = Number(item.quantity || item.Quantity || 0) || 0;
    if (!sku || qty <= 0) continue;

    if (BUNDLE_MAP[sku]) {
      var components = BUNDLE_MAP[sku];
      for (var b = 0; b < components.length; b++) {
        expanded.push({
          sku: components[b].sku,
          qty: components[b].qty * qty,
          bundleSku: sku
        });
      }
    } else {
      expanded.push({
        sku: sku,
        qty: qty,
        bundleSku: ''
      });
    }
  }
  return expanded;
}

function summarizeF5RecoveryItems_(items, maxLen) {
  var uomCache = {};
  var parts = [];
  items = items || [];

  for (var i = 0; i < items.length; i++) {
    var uom = '';
    if (F5_RECOVERY_CONFIG.INCLUDE_UOM) {
      uom = resolveSkuUom(items[i].sku, uomCache) || '';
    }
    var part = items[i].sku + ' x' + items[i].qty + (uom ? ' ' + uom : '');
    if (items[i].bundleSku) part += ' (bundle: ' + items[i].bundleSku + ')';
    parts.push(part);
  }

  var text = parts.join(', ');
  maxLen = maxLen || 200;
  if (text.length > maxLen) return text.substring(0, maxLen - 3) + '...';
  return text;
}

function writeF5RecoveryReport_(ss, summary, detailRows, reviewOrderRows) {
  var sheet = getF5RecoverySheet_(ss, true);
  reviewOrderRows = reviewOrderRows || [];
  var targetLastRow = Math.max(
    10,
    9 + Math.max(detailRows.length, 1) + (reviewOrderRows.length ? (4 + reviewOrderRows.length) : 0)
  );
  var clearLastRow = Math.max(sheet.getLastRow(), targetLastRow);

  sheet.getRange('A1:J6').breakApart();
  sheet.getRange(1, 1, clearLastRow, 10).clearContent().clearFormat();
  if (sheet.getConditionalFormatRules().length) {
    sheet.setConditionalFormatRules([]);
  }
  if (sheet.getFilter()) sheet.getFilter().remove();
  if (sheet.isSheetHidden && sheet.isSheetHidden()) {
    sheet.showSheet();
  }

  var nowText = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');

  sheet.getRange('A1:J1').clearContent().merge();
  sheet.getRange('A1').setValue('F5 Manual Recovery');
  sheet.getRange('A2:J2').setValues([[
    'Report Date', "'" + summary.reportDate,
    'Generated', nowText,
    'WASP Target', F5_RECOVERY_CONFIG.WASP_SITE + ' / ' + F5_RECOVERY_CONFIG.WASP_LOCATION,
    'Hidden Review', summary.reviewOrders,
    '', ''
  ]]);
  sheet.getRange('A3:J3').setValues([[
    'Action Orders', summary.actionOrders,
    'Manual Lines', summary.manualSkuLines,
    'Manual-Recovered', summary.manualRecovered,
    'Handled', summary.handledOrders,
    '', ''
  ]]);
  sheet.getRange('A4:J4').clearContent().merge();
  if (summary.reviewOrders > 0) {
    sheet.getRange('A4').setValue(
      'REVIEW REQUIRED - ' + summary.reviewOrders + ' hidden review shipment(s) were excluded from the action list. Do not finalize until they are checked.'
    );
    sheet.getRange('A4').setBackground('#f8d7da').setFontColor('#721c24').setFontWeight('bold');
  } else {
    if ((Number(summary.staleCleanupOrders || 0) > 0) && Number(summary.actionOrders || 0) === 0) {
      sheet.getRange('A4').setValue(
        'SAFE TO FINALIZE - no manual deductions remain. Finalize now to clean stale duplicate F5 rows for ' +
        summary.staleCleanupOrders + ' order(s).'
      );
    } else {
      sheet.getRange('A4').setValue(
        'SAFE TO FINALIZE - no hidden review shipments for this report date.'
      );
    }
    sheet.getRange('A4').setBackground('#d4edda').setFontColor('#155724').setFontWeight('bold');
  }
  sheet.getRange('A5:J5').clearContent().merge();
  sheet.getRange('A5').setValue(
    'F5 live shipping and void polling are paused while this recovery is open. Use Manual Lines To Deduct for action and Review Required for blocked cases.'
  );
  sheet.getRange('A6:J6').clearContent().merge();
  if ((Number(summary.staleCleanupOrders || 0) > 0) && Number(summary.actionOrders || 0) === 0) {
    sheet.getRange('A6').setValue(
      'No manual WASP deductions remain. Run Katana-WASP > Finalize F5 Manual Recovery to clean stale duplicate F5 rows, resume polling, and hide this tab.'
    );
  } else {
    sheet.getRange('A6').setValue(
      'After manual WASP deductions, run Katana-WASP > Finalize F5 Manual Recovery. That resumes polling and hides this tab automatically.'
    );
  }

  sheet.getRange('A1:J1').setBackground('#d1ecf1').setFontWeight('bold').setFontSize(14);
  sheet.getRange('A1').setHorizontalAlignment('left');
  sheet.getRange('A2:J3').setBackground('#eef7fb');
  sheet.getRange('A2:J3').setFontWeight('bold');
  sheet.getRange('A2:J3').setWrap(false);
  sheet.getRange('A5:J6').setBackground('#eef7fb');
  sheet.getRange('A5:J6').setFontWeight('bold');

  var row = 8;

  sheet.getRange(row, 1).setValue('Manual Lines To Deduct');
  sheet.getRange(row, 1, 1, 8).setBackground('#d1ecf1').setFontWeight('bold');
  row++;
  sheet.getRange(row, 1, 1, 8).setValues([[
    'SKU', 'Qty', 'WASP Site', 'WASP Location', 'Order #', 'Shipment ID', 'Ship Date', 'Why'
  ]]);
  sheet.getRange(row, 1, 1, 8).setBackground('#fff3cd').setFontWeight('bold');
  row++;
  if (detailRows.length > 0) {
    sheet.getRange(row, 1, detailRows.length, 8).setValues(detailRows);
    row += detailRows.length;
  } else {
    sheet.getRange(row, 1).setValue('No manual lines.');
    row++;
  }

  if (reviewOrderRows.length > 0) {
    row += 2;
    sheet.getRange(row, 1).setValue('Review Required');
    sheet.getRange(row, 1, 1, 4).setBackground('#f8d7da').setFontWeight('bold');
    row++;
    sheet.getRange(row, 1, 1, 4).setValues([[
      'Order #', 'Shipment ID', 'Ship Date', 'Reason'
    ]]);
    sheet.getRange(row, 1, 1, 4).setBackground('#fdecef').setFontWeight('bold');
    row++;
    sheet.getRange(row, 1, reviewOrderRows.length, 4).setValues(reviewOrderRows);
  }

  sheet.setFrozenRows(9);
  sheet.setColumnWidth(1, 140);
  sheet.setColumnWidth(2, 90);
  sheet.setColumnWidth(3, 130);
  sheet.setColumnWidth(4, 130);
  sheet.setColumnWidth(5, 110);
  sheet.setColumnWidth(6, 110);
  sheet.setColumnWidth(7, 130);
  sheet.setColumnWidth(8, 320);
  sheet.setColumnWidth(9, 120);
  sheet.setColumnWidth(10, 110);
  sheet.getRange(1, 1, targetLastRow, 10).setVerticalAlignment('middle');
  sheet.getRange(8, 1, Math.max(targetLastRow - 7, 1), 8).setWrap(true);
  sheet.getRange('B:B').setHorizontalAlignment('center');
  sheet.getRange('B2').setNumberFormat('@');
  sheet.getRange(10, 2, Math.max(detailRows.length, 1), 1).setNumberFormat('0.###');
  sheet.getRange('A2:J3').setHorizontalAlignment('left');
  sheet.getRange('B2:B3').setHorizontalAlignment('left');
  sheet.getRange('D2:D3').setHorizontalAlignment('left');
  sheet.getRange('F2:F3').setHorizontalAlignment('center');
  sheet.getRange('H2:H3').setHorizontalAlignment('center');
  sheet.getRange('J2:J3').setHorizontalAlignment('center');
  ss.setActiveSheet(sheet);

  return sheet;
}

function sendF5RecoverySlackSummary_(summary, aggregateRows, sheet) {
  aggregateRows = aggregateRows || [];
  var topLines = [];
  for (var i = 0; i < aggregateRows.length && i < 5; i++) {
    var row = aggregateRows[i];
    topLines.push('- ' + row[0] + ' x' + row[1]);
  }

  var sheetUrl = 'https://docs.google.com/spreadsheets/d/' + CONFIG.DEBUG_SHEET_ID + '/edit#gid=' + sheet.getSheetId();
  var message =
    'F5 outage recovery report built for ' + summary.reportDate + '\n' +
    'Active ShipStation shipments: ' + summary.activeShipments + '\n' +
    'Action orders: ' + summary.actionOrders + '\n' +
    'Review orders: ' + summary.reviewOrders + '\n' +
    'Manual-recovered hidden: ' + summary.manualRecovered + '\n' +
    'Manual SKU lines: ' + summary.manualSkuLines + '\n' +
    (topLines.length ? ('Top manual removals:\n' + topLines.join('\n') + '\n') : '') +
    'Sheet: ' + sheetUrl;

  sendSlackNotification(message);
}

function ensureF5RecoveryStateSheet_(ss) {
  ss = ss || getDebugSpreadsheet_();
  var sheet = ss.getSheetByName(F5_RECOVERY_CONFIG.STATE_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(F5_RECOVERY_CONFIG.STATE_SHEET_NAME);
    sheet.getRange(1, 1, 1, 6).setValues([[
      'Shipment ID', 'Order #', 'Report Date', 'Marked At', 'Source', 'Note'
    ]]);
    sheet.hideSheet();
  }
  return sheet;
}

function loadF5ManualRecoveryState_(ss) {
  var sheet = ensureF5RecoveryStateSheet_(ss);
  var state = {
    byShipmentId: {},
    byOrderDate: {}
  };
  if (sheet.getLastRow() < 2) return state;

  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).getValues();
  for (var i = 0; i < data.length; i++) {
    var shipmentId = String(data[i][0] || '').trim();
    var orderKey = normalizeF5RecoveryOrder_(data[i][1]);
    var reportDate = String(data[i][2] || '').trim();
    var record = {
      shipmentId: shipmentId,
      orderKey: orderKey,
      reportDate: reportDate,
      markedAt: String(data[i][3] || '').trim(),
      source: String(data[i][4] || '').trim(),
      note: String(data[i][5] || '').trim()
    };
    if (shipmentId) state.byShipmentId[shipmentId] = record;
    if (orderKey && reportDate) state.byOrderDate[orderKey + '|' + reportDate] = record;
  }
  return state;
}

function findF5ManualRecoveryRecord_(state, shipmentId, orderKey, reportDate) {
  state = state || loadF5ManualRecoveryState_();
  shipmentId = String(shipmentId || '').trim();
  orderKey = normalizeF5RecoveryOrder_(orderKey);
  reportDate = String(reportDate || '').trim();

  if (shipmentId && state.byShipmentId[shipmentId]) return state.byShipmentId[shipmentId];
  if (orderKey && reportDate && state.byOrderDate[orderKey + '|' + reportDate]) {
    return state.byOrderDate[orderKey + '|' + reportDate];
  }
  return null;
}

function recordF5ManualRecovery_(info) {
  info = info || {};
  var shipmentId = String(info.shipmentId || '').trim();
  var orderKey = normalizeF5RecoveryOrder_(info.orderKey);
  var reportDate = String(info.reportDate || '').trim();
  if (!shipmentId && !orderKey) return false;

  var state = loadF5ManualRecoveryState_();
  if (findF5ManualRecoveryRecord_(state, shipmentId, orderKey, reportDate)) {
    if (shipmentId) {
      CacheService.getScriptCache().put('ss_shipped_' + shipmentId, 'manual_recovery', 86400);
    }
    return false;
  }

  var sheet = ensureF5RecoveryStateSheet_();
  var markedAt = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  sheet.appendRow([
    shipmentId,
    orderKey ? ('#' + orderKey) : '',
    reportDate,
    markedAt,
    String(info.source || 'F5 Recovery Finalize'),
    String(info.note || '')
  ]);
  if (shipmentId) {
    CacheService.getScriptCache().put('ss_shipped_' + shipmentId, 'manual_recovery', 86400);
  }
  return true;
}

function normalizeF5RecoveryOrder_(value) {
  if (value === null || value === undefined) return '';
  return String(value).replace('#', '').trim();
}

function addF5RecoveryManualItem_(aggregateMap, detailRows, manualSeen, itemInfo, uomCache) {
  if (!itemInfo || !itemInfo.sku || !(Number(itemInfo.qty) > 0)) return;

  var dedupKey = [
    itemInfo.orderKey || '',
    itemInfo.shipmentId || '',
    itemInfo.sku || '',
    Number(itemInfo.qty) || 0,
    itemInfo.bundleSku || '',
    itemInfo.source || '',
    itemInfo.note || ''
  ].join('|');
  if (manualSeen[dedupKey]) return;
  manualSeen[dedupKey] = true;

  var uom = '';
  if (F5_RECOVERY_CONFIG.INCLUDE_UOM) {
    uom = resolveSkuUom(itemInfo.sku, uomCache) || '';
  }
  var aggKey = itemInfo.sku + '|' + uom;
  if (!aggregateMap[aggKey]) {
    aggregateMap[aggKey] = {
      sku: itemInfo.sku,
      uom: uom,
      qty: 0,
      orderMap: {}
    };
  }
  aggregateMap[aggKey].qty += Number(itemInfo.qty) || 0;
  if (itemInfo.orderKey) {
    aggregateMap[aggKey].orderMap[itemInfo.orderKey] = true;
  }

  var note = String(itemInfo.note || '');
  if (itemInfo.bundleSku) {
    note += (note ? ' | ' : '') + 'bundle: ' + itemInfo.bundleSku;
  }
  if (itemInfo.cacheHit) {
    note += (note ? ' | ' : '') + 'cache:' + itemInfo.cacheHit;
  }

  detailRows.push([
    itemInfo.sku,
    Number(itemInfo.qty) || 0,
    F5_RECOVERY_CONFIG.WASP_SITE,
    F5_RECOVERY_CONFIG.WASP_LOCATION,
    itemInfo.orderKey ? ('#' + itemInfo.orderKey) : '',
    itemInfo.shipmentId || '',
    itemInfo.shipDate || '',
    (itemInfo.source || 'Manual remove') + (note ? ' | ' + note : '')
  ]);
}

function buildF5RecoveryAggregateRows_(detailRows) {
  var map = {};
  detailRows = detailRows || [];

  for (var i = 0; i < detailRows.length; i++) {
    var row = detailRows[i];
    var orderKey = normalizeF5RecoveryOrder_(row[4]);
    var sku = String(row[0] || '').trim();
    var qty = Number(row[1]) || 0;
    if (!sku || qty <= 0) continue;

    if (!map[sku]) {
      map[sku] = { sku: sku, qty: 0, orderMap: {} };
    }
    map[sku].qty += qty;
    if (orderKey) map[sku].orderMap[orderKey] = true;
  }

  var keys = Object.keys(map);
  var rows = [];
  for (var k = 0; k < keys.length; k++) {
    var item = map[keys[k]];
    rows.push([
      item.sku,
      item.qty,
      Object.keys(item.orderMap).length,
      Object.keys(item.orderMap).sort().join(', ')
    ]);
  }

  rows.sort(function(aRow, bRow) {
    if (bRow[1] !== aRow[1]) return bRow[1] - aRow[1];
    return String(aRow[0]).localeCompare(String(bRow[0]));
  });
  return rows;
}

function addF5RecoveryReviewRow_(reviewOrderRows, info) {
  reviewOrderRows.push([
    info && info.orderKey ? ('#' + normalizeF5RecoveryOrder_(info.orderKey)) : '',
    info && info.shipmentId ? String(info.shipmentId).trim() : '',
    info && info.shipDate ? String(info.shipDate).trim() : '',
    info && info.reason ? String(info.reason).trim() : 'Review required'
  ]);
}

function getF5RecoveryHiddenReviewCount_(sheet) {
  sheet = sheet || getF5RecoverySheet_(getDebugSpreadsheet_(), false);
  if (!sheet) return 0;

  var summary = sheet.getRange('A2:J3').getDisplayValues();
  for (var r = 0; r < summary.length; r++) {
    for (var c = 0; c < summary[r].length - 1; c++) {
      if (String(summary[r][c] || '').trim() === 'Hidden Review') {
        return Number(summary[r][c + 1] || 0) || 0;
      }
    }
  }

  var banner = String(sheet.getRange('A4').getDisplayValue() || '').trim();
  var match = banner.match(/REVIEW REQUIRED\s*-\s*(\d+)/i);
  return match ? (Number(match[1]) || 0) : 0;
}

function parseF5RecoverySubItem_(detailText) {
  var text = String(detailText || '');
  var match = text.match(/([A-Za-z0-9._\/-]+)\s*x\s*([0-9]+(?:\.[0-9]+)?)/i);
  if (!match) return null;

  var bundleMatch = text.match(/\(bundle:\s*([^)]+)\)/i);
  return {
    sku: match[1],
    qty: Number(match[2]) || 0,
    bundleSku: bundleMatch ? bundleMatch[1] : ''
  };
}

function deriveF5RecoveryMissingItemsFromShipment_(shipment, f5Info) {
  if (!shipment || !f5Info) return [];

  var expected = expandF5RecoveryItems_(shipment.shipmentItems || shipment.items || shipment.orderItems || []);
  var deducted = f5Info.deductedItems || [];
  if (!expected.length) return [];

  var expectedMap = {};
  var deductedMap = {};
  var keys = {};
  var results = [];

  for (var i = 0; i < expected.length; i++) {
    var expSku = String(expected[i].sku || '').trim();
    var expQty = Number(expected[i].qty || 0) || 0;
    if (!expSku || !(expQty > 0)) continue;
    expectedMap[expSku] = (expectedMap[expSku] || 0) + expQty;
    keys[expSku] = true;
  }

  for (var j = 0; j < deducted.length; j++) {
    var dedSku = String(deducted[j].sku || '').trim();
    var dedQty = Number(deducted[j].qty || 0) || 0;
    if (!dedSku || !(dedQty > 0)) continue;
    deductedMap[dedSku] = (deductedMap[dedSku] || 0) + dedQty;
    keys[dedSku] = true;
  }

  var skuKeys = Object.keys(keys);
  for (var k = 0; k < skuKeys.length; k++) {
    var sku = skuKeys[k];
    var remaining = (expectedMap[sku] || 0) - (deductedMap[sku] || 0);
    if (remaining > 0) {
      results.push({ sku: sku, qty: remaining, bundleSku: '' });
    }
  }

  return results;
}

function filterF5RecoveryFailedItemsToShipment_(failedItems, shipment) {
  failedItems = failedItems || [];
  shipment = shipment || null;
  if (!shipment || !failedItems.length) return [];

  var expected = expandF5RecoveryItems_(shipment.shipmentItems || shipment.items || shipment.orderItems || []);
  if (!expected.length) return [];

  var expectedMap = {};
  var results = [];

  for (var i = 0; i < expected.length; i++) {
    var expectedSku = String(expected[i].sku || '').trim();
    var expectedQty = Number(expected[i].qty || 0) || 0;
    if (!expectedSku || !(expectedQty > 0)) continue;
    expectedMap[expectedSku] = (expectedMap[expectedSku] || 0) + expectedQty;
  }

  for (var j = 0; j < failedItems.length; j++) {
    var failSku = String(failedItems[j].sku || '').trim();
    var failQty = Number(failedItems[j].qty || 0) || 0;
    if (!failSku || !(failQty > 0)) continue;
    if (!(expectedMap[failSku] > 0)) continue;

    var matchedQty = Math.min(expectedMap[failSku], failQty);
    if (!(matchedQty > 0)) continue;

    expectedMap[failSku] -= matchedQty;
    results.push({
      sku: failSku,
      qty: matchedQty,
      bundleSku: failedItems[j].bundleSku || '',
      source: 'F5 failed line',
      note: 'Validated against ShipStation shipment items'
    });
  }

  return results;
}

function buildF5RecoveryDirectFailedItems_(f5Info) {
  f5Info = f5Info || {};
  var failedItems = f5Info.failedItems || [];
  if (!failedItems.length) return [];

  var deductedMap = {};
  var deductedItems = f5Info.deductedItems || [];
  for (var i = 0; i < deductedItems.length; i++) {
    var dSku = String(deductedItems[i].sku || '').trim();
    var dQty = Number(deductedItems[i].qty || 0) || 0;
    if (!dSku || !(dQty > 0)) continue;
    deductedMap[dSku + '|' + dQty] = true;
  }

  var seen = {};
  var results = [];
  for (var j = 0; j < failedItems.length; j++) {
    var failSku = String(failedItems[j].sku || '').trim();
    var failQty = Number(failedItems[j].qty || 0) || 0;
    if (!failSku || !(failQty > 0)) continue;

    var key = failSku + '|' + failQty;
    if (deductedMap[key] || seen[key]) continue;
    seen[key] = true;

    results.push({
      sku: failSku,
      qty: failQty,
      bundleSku: failedItems[j].bundleSku || '',
      source: 'F5 failed line',
      note: 'Derived directly from unresolved F5 failed lines'
    });
  }

  return results;
}

function formatF5RecoveryDate_(value) {
  if (!value) return '';
  var d = (value instanceof Date) ? value : new Date(value);
  if (isNaN(d.getTime())) return String(value).substring(0, 10);
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function formatF5RecoveryDateTime_(value) {
  if (!value) return '';
  var d = (value instanceof Date) ? value : new Date(value);
  if (isNaN(d.getTime())) return String(value);
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
}

function getF5RecoveryShipToName_(shipment) {
  var shipTo = shipment && shipment.shipTo ? shipment.shipTo : {};
  return shipTo.name || shipTo.company || '';
}
