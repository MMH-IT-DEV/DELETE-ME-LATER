// ============================================
// 08_HealthCheck.gs — Daily Katana-WASP Item Health Check
// ============================================
// Compares items across both systems: lot tracking, quantity,
// UOM, category, and cost. Writes report tab + posts summary
// to debug script for Activity logging and Slack notifications.
// ============================================

var HEALTH_CHECK_TAB_NAME = 'Health Check';
var HEALTH_CHECK_QTY_TOLERANCE = 0.5;
var HEALTH_CHECK_COST_TOLERANCE_PCT = 0.05;

/**
 * Entry point — called daily by time-based trigger or from menu.
 */
function runDailyHealthCheck() {
  var startTime = new Date();
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  try {
    // Fetch all data from both systems
    var katanaData = fetchAllKatanaData();
    var waspData = fetchAllWaspData();
    var batchData = fetchAllBatchData(katanaData);
    var waspLotSettings = fetchWaspItemLotSettings(true);

    // Merge batch tracking info
    if (katanaData.batchTrackingSkus) {
      var btSkus = Object.keys(katanaData.batchTrackingSkus);
      for (var i = 0; i < btSkus.length; i++) {
        batchData.lotTrackedSkus[btSkus[i]] = true;
      }
    }

    // Build comparison maps
    var katanaMap = buildKatanaSkuMap_(katanaData, batchData);
    var waspMap = buildWaspSkuMap_(waspData, waspLotSettings);

    // Run all 5 comparisons
    var mismatches = compareHealthDimensions_(katanaMap, waspMap);

    // Count by dimension
    var counts = { lot: 0, qty: 0, uom: 0, category: 0, cost: 0 };
    var skuSet = {};
    for (var m = 0; m < mismatches.length; m++) {
      counts[mismatches[m].dimension] = (counts[mismatches[m].dimension] || 0) + 1;
      skuSet[mismatches[m].sku] = true;
    }
    var skusWithMismatches = Object.keys(skuSet);
    var totalMismatches = mismatches.length;

    // Write Health Check tab
    var timestamp = Utilities.formatDate(new Date(), 'America/Vancouver', 'yyyy-MM-dd HH:mm');
    writeHealthCheckTab_(ss, mismatches, timestamp, counts, Object.keys(katanaMap).length, Object.keys(waspMap).length);

    // Post summary to debug script
    var duration = ((new Date()) - startTime) / 1000;
    postHealthSummaryToDebug_(counts, totalMismatches, 0, Object.keys(katanaMap).length, Object.keys(waspMap).length, skusWithMismatches, timestamp, duration);

    // Heartbeat
    var details = 'Items: K=' + Object.keys(katanaMap).length + ' W=' + Object.keys(waspMap).length + ' | Mismatches: ' + totalMismatches;
    sendHeartbeatToMonitor_('2026_Katana-Wasp Health Check', totalMismatches === 0 ? 'success' : 'error', details);

    // Log to sync history
    logSyncHistory(ss, 'HEALTH_CHECK', details, 0, '', duration, totalMismatches === 0 ? 'SUCCESS' : 'MISMATCH');

    try { ss.toast('Health Check: ' + totalMismatches + ' mismatches found', 'Health Check', 5); } catch (e) {}

  } catch (e) {
    var duration = ((new Date()) - startTime) / 1000;
    logSyncHistory(ss, 'HEALTH_CHECK', 'ERROR: ' + e.message, 1, e.stack || e.message, duration, 'ERROR');
    sendHeartbeatToMonitor_('2026_Katana-Wasp Health Check', 'error', e.message);
    Logger.log('Health Check error: ' + e.message + '\n' + e.stack);
    try { ss.toast('Health Check error: ' + e.message, 'Health Check', 10); } catch (e2) {}
  }
}

// ============================================
// SKU MAP BUILDERS
// ============================================

/**
 * Build Katana SKU map: { sku -> { lotTracking, uom, purchaseUom, conversionRate,
 *   category, avgCost, locations: { locationName -> totalQty } } }
 */
function buildKatanaSkuMap_(katanaData, batchData) {
  var map = {};
  var items = katanaData.items || [];

  for (var i = 0; i < items.length; i++) {
    var item = items[i];
    var sku = item.sku;
    if (!sku || isSkippedSku(sku)) continue;

    if (!map[sku]) {
      map[sku] = {
        name: item.name || '',
        type: item.type || '',
        lotTracking: !!(batchData.lotTrackedSkus[sku] || katanaData.batchTrackingSkus[sku]),
        uom: item.uom || 'ea',
        purchaseUom: item.purchaseUom || '',
        conversionRate: item.conversionRate || null,
        category: item.category || '',
        avgCost: item.avgCost || 0,
        locations: {}
      };
    }

    var loc = item.locationName || '';
    if (loc) {
      map[sku].locations[loc] = (map[sku].locations[loc] || 0) + (item.qtyInStock || 0);
    }
  }

  return map;
}

/**
 * Build WASP SKU map: { sku -> { lotTracking, uom, category, cost,
 *   sites: { siteName -> totalQty } } }
 * Aggregates multi-lot rows into total qty per site.
 */
function buildWaspSkuMap_(waspData, waspLotSettings) {
  var map = {};
  var items = waspData.items || [];

  for (var i = 0; i < items.length; i++) {
    var item = items[i];
    var sku = item.itemNumber;
    if (!sku || isSkippedSku(sku)) continue;

    if (!map[sku]) {
      var lotSetting = waspLotSettings && waspLotSettings.hasOwnProperty(sku) ? waspLotSettings[sku] : undefined;
      if (lotSetting === undefined && item.lotTracking !== undefined) lotSetting = item.lotTracking;

      map[sku] = {
        description: item.description || '',
        lotTracking: !!lotSetting,
        uom: item.uom || 'Each',
        category: item.category || '',
        cost: item.cost || 0,
        sites: {}
      };
    }

    var site = item.siteName || '';
    if (site) {
      map[sku].sites[site] = (map[sku].sites[site] || 0) + (item.qtyAvailable || 0);
    }
  }

  return map;
}

// ============================================
// COMPARISON FUNCTIONS
// ============================================

/**
 * Run all 5 comparisons. Returns array of mismatch objects.
 */
function compareHealthDimensions_(katanaMap, waspMap) {
  var mismatches = [];

  compareLotTracking_(katanaMap, waspMap, mismatches);
  compareQuantities_(katanaMap, waspMap, mismatches);
  compareUoms_(katanaMap, waspMap, mismatches);
  compareCategories_(katanaMap, waspMap, mismatches);
  compareCosts_(katanaMap, waspMap, mismatches);

  return mismatches;
}

/**
 * Compare lot tracking settings between Katana and WASP.
 */
function compareLotTracking_(katanaMap, waspMap, mismatches) {
  for (var sku in katanaMap) {
    if (!katanaMap.hasOwnProperty(sku)) continue;
    if (!waspMap[sku]) continue;

    var kLot = katanaMap[sku].lotTracking;
    var wLot = waspMap[sku].lotTracking;

    if (kLot !== wLot) {
      mismatches.push({
        sku: sku,
        dimension: 'lot',
        site: '',
        katanaVal: kLot ? 'Enabled' : 'Disabled',
        waspVal: wLot ? 'Enabled' : 'Disabled',
        delta: '',
        notes: kLot && !wLot ? 'WASP needs lot enabled' : 'WASP has lot but Katana does not'
      });
    }
  }
}

/**
 * Compare quantities per location using SYNC_LOCATION_MAP.
 */
function compareQuantities_(katanaMap, waspMap, mismatches) {
  for (var sku in katanaMap) {
    if (!katanaMap.hasOwnProperty(sku)) continue;
    if (!waspMap[sku]) continue;

    var kItem = katanaMap[sku];
    var wItem = waspMap[sku];

    for (var locName in kItem.locations) {
      if (!kItem.locations.hasOwnProperty(locName)) continue;

      var kQty = kItem.locations[locName] || 0;

      // Map Katana location to WASP site
      var mapKey = locName + '|' + kItem.type;
      var waspDest = SYNC_LOCATION_MAP[mapKey] || SYNC_LOCATION_MAP[locName + '|'] || null;
      if (!waspDest) continue;

      var wSite = waspDest.site;
      var wQty = wItem.sites[wSite] || 0;

      // Apply UOM conversion if needed
      var factor = kItem.conversionRate || 1;
      if (factor !== 1) {
        kQty = kQty * factor;
      }

      var delta = Math.abs(kQty - wQty);
      if (delta > HEALTH_CHECK_QTY_TOLERANCE) {
        mismatches.push({
          sku: sku,
          dimension: 'qty',
          site: wSite,
          katanaVal: String(kQty),
          waspVal: String(wQty),
          delta: String(Math.round((kQty - wQty) * 100) / 100),
          notes: locName + ' → ' + wSite + (factor !== 1 ? ' (converted x' + factor + ')' : '')
        });
      }
    }
  }
}

/**
 * Compare UOM strings using normalized comparison.
 */
function compareUoms_(katanaMap, waspMap, mismatches) {
  var normalizeUom = function(u) {
    if (!u) return 'each';
    var s = String(u).trim().toLowerCase();
    if (s === 'ea' || s === 'pc' || s === 'pcs' || s === '1') return 'each';
    if (s === 'g' || s === 'grams' || s === 'gram') return 'grams';
    if (s === 'kg') return 'kg';
    if (s === 'lbs' || s === 'lb' || s === 'pound') return 'pound';
    return s;
  };

  for (var sku in katanaMap) {
    if (!katanaMap.hasOwnProperty(sku)) continue;
    if (!waspMap[sku]) continue;

    var kUom = normalizeUom(katanaMap[sku].uom);
    var wUom = normalizeUom(waspMap[sku].uom);

    // Skip items with UOM conversion — they intentionally differ
    if (katanaMap[sku].conversionRate && katanaMap[sku].conversionRate !== 1) continue;

    if (kUom !== wUom) {
      mismatches.push({
        sku: sku,
        dimension: 'uom',
        site: '',
        katanaVal: katanaMap[sku].uom || '(empty)',
        waspVal: waspMap[sku].uom || '(empty)',
        delta: '',
        notes: 'Stocking unit mismatch'
      });
    }
  }
}

/**
 * Compare category strings (case-insensitive).
 */
function compareCategories_(katanaMap, waspMap, mismatches) {
  for (var sku in katanaMap) {
    if (!katanaMap.hasOwnProperty(sku)) continue;
    if (!waspMap[sku]) continue;

    var kCat = String(katanaMap[sku].category || '').trim().toLowerCase();
    var wCat = String(waspMap[sku].category || '').trim().toLowerCase();

    // Skip if both empty
    if (!kCat && !wCat) continue;

    if (kCat !== wCat) {
      mismatches.push({
        sku: sku,
        dimension: 'category',
        site: '',
        katanaVal: katanaMap[sku].category || '(empty)',
        waspVal: waspMap[sku].category || '(empty)',
        delta: '',
        notes: 'Category mismatch'
      });
    }
  }
}

/**
 * Compare cost within tolerance. Skip items with zero cost on either side.
 */
function compareCosts_(katanaMap, waspMap, mismatches) {
  for (var sku in katanaMap) {
    if (!katanaMap.hasOwnProperty(sku)) continue;
    if (!waspMap[sku]) continue;

    var kCost = katanaMap[sku].avgCost || 0;
    var wCost = waspMap[sku].cost || 0;

    // Skip if either is zero (unset)
    if (kCost === 0 || wCost === 0) continue;

    var maxCost = Math.max(kCost, wCost);
    var pctDiff = Math.abs(kCost - wCost) / maxCost;

    if (pctDiff > HEALTH_CHECK_COST_TOLERANCE_PCT) {
      mismatches.push({
        sku: sku,
        dimension: 'cost',
        site: '',
        katanaVal: String(Math.round(kCost * 10000) / 10000),
        waspVal: String(Math.round(wCost * 10000) / 10000),
        delta: String(Math.round((kCost - wCost) * 10000) / 10000),
        notes: Math.round(pctDiff * 100) + '% difference'
      });
    }
  }
}

// ============================================
// TAB WRITER
// ============================================

/**
 * Write or refresh the Health Check tab.
 */
function writeHealthCheckTab_(ss, mismatches, timestamp, counts, katanaCount, waspCount) {
  var sheet = ss.getSheetByName(HEALTH_CHECK_TAB_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(HEALTH_CHECK_TAB_NAME);
  }

  // Clear existing data
  if (sheet.getLastRow() > 0) {
    sheet.clear();
  }

  // Header row
  var headers = ['Timestamp', 'SKU', 'Dimension', 'Site', 'Katana Value', 'WASP Value', 'Delta', 'Status', 'Notes'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setBackground('#263238').setFontColor('#ffffff').setFontWeight('bold');
  sheet.setFrozenRows(1);

  // Summary banner row
  var summaryParts = [
    'Items: K=' + katanaCount + ' W=' + waspCount,
    'Lot: ' + (counts.lot || 0),
    'Qty: ' + (counts.qty || 0),
    'UOM: ' + (counts.uom || 0),
    'Category: ' + (counts.category || 0),
    'Cost: ' + (counts.cost || 0),
    'Total: ' + mismatches.length
  ];
  var summaryRow = [timestamp, 'SUMMARY', summaryParts.join(' | '), '', '', '', '', mismatches.length === 0 ? 'ALL CLEAR' : 'MISMATCHES', ''];
  sheet.getRange(2, 1, 1, headers.length).setValues([summaryRow]);
  sheet.getRange(2, 1, 1, headers.length).setBackground(mismatches.length === 0 ? '#e8f5e9' : '#fff3e0').setFontWeight('bold');

  // Mismatch rows
  if (mismatches.length > 0) {
    var rows = [];
    for (var i = 0; i < mismatches.length; i++) {
      var m = mismatches[i];
      rows.push([timestamp, m.sku, m.dimension, m.site || '', m.katanaVal || '', m.waspVal || '', m.delta || '', 'MISMATCH', m.notes || '']);
    }
    sheet.getRange(3, 1, rows.length, headers.length).setValues(rows);

    // Color code by dimension
    for (var r = 0; r < rows.length; r++) {
      var dim = mismatches[r].dimension;
      var bg = '#ffebee'; // default red
      if (dim === 'qty') bg = '#fff3e0'; // amber
      if (dim === 'cost') bg = '#fff8e1'; // light yellow
      if (dim === 'lot') bg = '#fce4d6'; // peach
      sheet.getRange(3 + r, 1, 1, headers.length).setBackground(bg);
    }
  }

  // Auto-resize columns
  for (var c = 1; c <= headers.length; c++) {
    sheet.autoResizeColumn(c);
  }
}

// ============================================
// CROSS-SCRIPT POST TO DEBUG
// ============================================

/**
 * Post summary to debug script for Activity logging and Slack.
 * Reuses the enginPreMark_ pattern.
 */
function postHealthSummaryToDebug_(counts, totalMismatches, fixed, katanaCount, waspCount, skusWithMismatches, timestamp, duration) {
  var url = PropertiesService.getScriptProperties().getProperty('GAS_WEBHOOK_URL');
  if (!url) {
    Logger.log('Health Check: GAS_WEBHOOK_URL not set, skipping debug post');
    return;
  }

  var payload = {
    action: 'health_check_report',
    timestamp: timestamp,
    duration: duration,
    katanaCount: katanaCount,
    waspCount: waspCount,
    mismatches: counts,
    fixed: fixed,
    totalMismatches: totalMismatches,
    skusWithMismatches: skusWithMismatches.slice(0, 20)
  };

  try {
    var resp = UrlFetchApp.fetch(url, {
      method: 'POST',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
    Logger.log('Health Check debug post: HTTP ' + resp.getResponseCode());
  } catch (e) {
    Logger.log('Health Check debug post error: ' + e.message);
  }
}

// ============================================
// TRIGGER MANAGEMENT
// ============================================

/**
 * Create daily trigger at 6 AM Vancouver.
 */
function createDailyHealthCheckTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'runDailyHealthCheck') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  ScriptApp.newTrigger('runDailyHealthCheck')
    .timeBased()
    .everyDays(1)
    .atHour(6)
    .create();
  try {
    SpreadsheetApp.getActiveSpreadsheet().toast('Daily Health Check enabled: runs at 6 AM Vancouver', 'Health Check', 5);
  } catch (e) {}
}

/**
 * Remove daily health check trigger.
 */
function removeHealthCheckTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  var removed = 0;
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'runDailyHealthCheck') {
      ScriptApp.deleteTrigger(triggers[i]);
      removed++;
    }
  }
  try {
    SpreadsheetApp.getActiveSpreadsheet().toast(removed > 0 ? 'Daily Health Check disabled' : 'No Health Check trigger found', 'Health Check', 5);
  } catch (e) {}
}
