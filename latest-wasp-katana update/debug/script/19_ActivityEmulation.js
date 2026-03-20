// ============================================
// 19_ActivityEmulation.gs - BULK KATANA ACTIVITY EMULATION
// ============================================
// Safe preview tooling:
// - reads live Katana objects
// - rebuilds expected Activity blocks without mutating WASP
// - renders previews through the shared Activity renderer
// ============================================

var KATANA_ACTIVITY_EMU_CONFIG = {
  INPUT_SHEET_NAME: 'Activity QA Inputs',
  PREVIEW_SHEET_NAME: 'Activity QA Katana',
  REPORT_SHEET_NAME: 'Activity QA Katana Report',
  DEFAULT_RECENT_PER_FLOW: 3,
  DEEP_RECENT_PER_FLOW: 5,
  DATA_START_ROW: 4
};

function setupKatanaActivityEmulationInputs() {
  var ss = typeof getDebugSpreadsheet_ === 'function'
    ? getDebugSpreadsheet_()
    : SpreadsheetApp.openById(CONFIG.DEBUG_SHEET_ID);
  var sheet = ensureKatanaActivityEmuInputSheet_(ss);
  resetKatanaActivityEmuInputSheet_(sheet);
  writeKatanaActivityEmuInputRows_(sheet, buildDefaultKatanaActivityEmuRows_());
  return {
    status: 'ok',
    inputSheet: KATANA_ACTIVITY_EMU_CONFIG.INPUT_SHEET_NAME
  };
}

function populateRecentKatanaActivityEmulationInputs(limitPerFlow) {
  var ss = typeof getDebugSpreadsheet_ === 'function'
    ? getDebugSpreadsheet_()
    : SpreadsheetApp.openById(CONFIG.DEBUG_SHEET_ID);
  var sheet = ensureKatanaActivityEmuInputSheet_(ss);
  resetKatanaActivityEmuInputSheet_(sheet);
  writeKatanaActivityEmuInputRows_(sheet, buildRecentKatanaActivityEmuRows_(limitPerFlow || KATANA_ACTIVITY_EMU_CONFIG.DEFAULT_RECENT_PER_FLOW));
  return {
    status: 'ok',
    inputSheet: KATANA_ACTIVITY_EMU_CONFIG.INPUT_SHEET_NAME
  };
}

function runKatanaActivityEmulation() {
  var ss = typeof getDebugSpreadsheet_ === 'function'
    ? getDebugSpreadsheet_()
    : SpreadsheetApp.openById(CONFIG.DEBUG_SHEET_ID);
  var inputSheet = ensureKatanaActivityEmuInputSheet_(ss);
  var previewSheet = ensureKatanaActivityEmuPreviewSheet_(ss);
  var reportSheet = ensureKatanaActivityEmuReportSheet_(ss);

  resetKatanaActivityEmuPreviewSheet_(previewSheet);
  resetKatanaActivityEmuReportSheet_(reportSheet);

  var inputRows = readKatanaActivityEmuInputRows_(inputSheet);
  var reportRows = [];
  var rendered = 0;

  for (var i = 0; i < inputRows.length; i++) {
    var entry = inputRows[i];
    if (!entry.enabled) continue;

    try {
      var built = buildKatanaActivityFixtureFromInput_(entry, i);
      var findings = validateActivityFixture_(built.fixture);
      renderActivityFixtureToSheet_(previewSheet, built.fixture, 'KE-' + padActivityQANumber_(rendered + 1));
      rendered++;
      reportRows.push({
        scope: 'Katana',
        flow: built.fixture.flow,
        scenario: built.fixture.scenario,
        score: findings.length === 0 ? 100 : Math.max(0, 100 - findings.length * 15),
        findings: findings,
        ref: built.fixture.reference,
        rowNum: '',
        notes: built.notes.join(' | ')
      });
    } catch (err) {
      reportRows.push({
        scope: 'Katana',
        flow: entry.flow || '?',
        scenario: entry.scenario || (entry.type + ' ' + entry.katanaId),
        score: 0,
        findings: [err.message || String(err)],
        ref: entry.type + ':' + entry.katanaId,
        rowNum: entry.rowNum,
        notes: entry.notes || ''
      });
    }
  }

  writeKatanaActivityEmulationReport_(reportSheet, reportRows, rendered);
  return {
    status: 'ok',
    rendered: rendered,
    inputSheet: KATANA_ACTIVITY_EMU_CONFIG.INPUT_SHEET_NAME,
    previewSheet: KATANA_ACTIVITY_EMU_CONFIG.PREVIEW_SHEET_NAME,
    reportSheet: KATANA_ACTIVITY_EMU_CONFIG.REPORT_SHEET_NAME
  };
}

function runRecentKatanaActivityEmulation(limitPerFlow) {
  populateRecentKatanaActivityEmulationInputs(limitPerFlow || KATANA_ACTIVITY_EMU_CONFIG.DEFAULT_RECENT_PER_FLOW);
  return runKatanaActivityEmulation();
}

function populateDeepKatanaActivityEmulationInputs() {
  return populateRecentKatanaActivityEmulationInputs(KATANA_ACTIVITY_EMU_CONFIG.DEEP_RECENT_PER_FLOW);
}

function runDeepKatanaActivityEmulation() {
  return runRecentKatanaActivityEmulation(KATANA_ACTIVITY_EMU_CONFIG.DEEP_RECENT_PER_FLOW);
}

function ensureKatanaActivityEmuInputSheet_(ss) {
  var sheet = ss.getSheetByName(KATANA_ACTIVITY_EMU_CONFIG.INPUT_SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(KATANA_ACTIVITY_EMU_CONFIG.INPUT_SHEET_NAME);
  return sheet;
}

function ensureKatanaActivityEmuPreviewSheet_(ss) {
  var sheet = ss.getSheetByName(KATANA_ACTIVITY_EMU_CONFIG.PREVIEW_SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(KATANA_ACTIVITY_EMU_CONFIG.PREVIEW_SHEET_NAME);
  return sheet;
}

function ensureKatanaActivityEmuReportSheet_(ss) {
  var sheet = ss.getSheetByName(KATANA_ACTIVITY_EMU_CONFIG.REPORT_SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(KATANA_ACTIVITY_EMU_CONFIG.REPORT_SHEET_NAME);
  return sheet;
}

function clearKatanaActivityEmuSheet_(sheet, minCols) {
  var lastRow = sheet.getLastRow();
  var lastCol = Math.max(sheet.getLastColumn(), minCols || 1);
  if (lastRow > 0 && lastCol > 0) {
    sheet.getRange(1, 1, lastRow, lastCol).clear();
  }
}

function resetKatanaActivityEmuInputSheet_(sheet) {
  runSpreadsheetOpWithRetry_('resetKatanaActivityEmuInputSheet', function() {
    sheet.getRange(1, 1, 3, 10).breakApart();
    clearKatanaActivityEmuSheet_(sheet, 10);
    sheet.setFrozenRows(3);

    var widths = { 1: 80, 2: 80, 3: 90, 4: 100, 5: 130, 6: 80, 7: 180, 8: 280, 9: 120, 10: 120 };
    for (var col in widths) sheet.setColumnWidth(Number(col), widths[col]);

    sheet.getRange(1, 1, 1, 10).merge();
    sheet.getRange(2, 1, 1, 10).merge();

    sheet.getRange(1, 1).setValue('KATANA ACTIVITY EMULATION INPUTS')
      .setBackground(ACTIVITY_QA_CONFIG.TITLE_BG)
      .setFontColor(ACTIVITY_QA_CONFIG.TITLE_TXT)
      .setFontWeight('bold')
      .setFontSize(12)
      .setHorizontalAlignment('center');

    sheet.getRange(2, 1).setValue('Set Enabled=TRUE on the rows you want to preview. Flow chooses the Activity format. Source Override is optional and mostly useful for F2. Reverse=TRUE renders the inverse Activity shape.')
      .setBackground(ACTIVITY_QA_CONFIG.HEADER_BG)
      .setFontColor(ACTIVITY_QA_CONFIG.SUBTITLE_TXT)
      .setFontStyle('italic')
      .setWrap(true);

    sheet.getRange(3, 1, 1, 10).setValues([[
      'Enabled', 'Flow', 'Katana Type', 'Katana ID', 'Source Override',
      'Reverse', 'Scenario', 'Notes', 'Ref Hint', 'Status Hint'
    ]]);
    sheet.getRange(3, 1, 1, 10)
      .setBackground(ACTIVITY_QA_CONFIG.REPORT_HDR_BG)
      .setFontColor(ACTIVITY_QA_CONFIG.REPORT_HDR_TXT)
      .setFontWeight('bold')
      .setHorizontalAlignment('center');
  });
}

function resetKatanaActivityEmuPreviewSheet_(sheet) {
  runSpreadsheetOpWithRetry_('resetKatanaActivityEmuPreviewSheet', function() {
    sheet.getRange(1, 1, 3, 15).breakApart();
    clearKatanaActivityEmuSheet_(sheet, 15);

    sheet.setFrozenRows(3);
    sheet.setRowHeight(1, 26);
    sheet.setRowHeight(2, 24);
    sheet.setRowHeight(3, 26);

    sheet.setColumnWidth(1, 90);
    sheet.setColumnWidth(2, 135);
    sheet.setColumnWidth(3, 150);
    sheet.setColumnWidth(4, 640);
    sheet.setColumnWidth(5, 110);
    sheet.setColumnWidth(6, 220);
    sheet.setColumnWidth(7, 20);
    for (var c = 8; c <= 15; c++) sheet.setColumnWidth(c, 90);

    sheet.getRange(1, 1, 1, 6).merge();
    sheet.getRange(2, 1, 1, 6).merge();

    sheet.getRange(1, 1).setValue('KATANA ACTIVITY EMULATION PREVIEW')
      .setBackground(ACTIVITY_QA_CONFIG.TITLE_BG)
      .setFontColor(ACTIVITY_QA_CONFIG.TITLE_TXT)
      .setFontWeight('bold')
      .setFontSize(12)
      .setHorizontalAlignment('center');

    sheet.getRange(2, 1).setValue('Preview rows are rebuilt from live Katana objects and rendered through the shared Activity formatter. No WASP mutation happens here.')
      .setBackground(ACTIVITY_QA_CONFIG.HEADER_BG)
      .setFontColor(ACTIVITY_QA_CONFIG.SUBTITLE_TXT)
      .setFontStyle('italic')
      .setWrap(true);

    sheet.getRange(3, 1, 1, 6).setValues([['ID', 'Time', 'Flow', 'Details', 'Status', 'Error']]);
    sheet.getRange(3, 1, 1, 6)
      .setBackground('#2F4F6F')
      .setFontColor('#FFFFFF')
      .setFontWeight('bold')
      .setHorizontalAlignment('center');
  });
}

function resetKatanaActivityEmuReportSheet_(sheet) {
  runSpreadsheetOpWithRetry_('resetKatanaActivityEmuReportSheet', function() {
    sheet.getRange(1, 1, 2, 8).breakApart();
    clearKatanaActivityEmuSheet_(sheet, 8);

    sheet.setFrozenRows(4);
    sheet.setRowHeight(1, 26);
    sheet.setRowHeight(2, 24);
    sheet.setRowHeight(4, 24);

    var reportWidths = { 1: 70, 2: 80, 3: 190, 4: 80, 5: 360, 6: 140, 7: 100, 8: 280 };
    for (var col in reportWidths) sheet.setColumnWidth(Number(col), reportWidths[col]);

    sheet.getRange(1, 1, 1, 8).merge();
    sheet.getRange(2, 1, 1, 8).merge();

    sheet.getRange(1, 1).setValue('KATANA ACTIVITY EMULATION REPORT')
      .setBackground(ACTIVITY_QA_CONFIG.TITLE_BG)
      .setFontColor(ACTIVITY_QA_CONFIG.TITLE_TXT)
      .setFontWeight('bold')
      .setFontSize(12)
      .setHorizontalAlignment('center');

    sheet.getRange(2, 1).setValue('Built ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm'))
      .setBackground(ACTIVITY_QA_CONFIG.HEADER_BG)
      .setFontColor(ACTIVITY_QA_CONFIG.SUBTITLE_TXT)
      .setHorizontalAlignment('center');

    sheet.getRange(4, 1, 1, 8).setValues([['Scope', 'Flow', 'Scenario', 'Score', 'Findings', 'Ref', 'Row', 'Notes']]);
    sheet.getRange(4, 1, 1, 8)
      .setBackground(ACTIVITY_QA_CONFIG.REPORT_HDR_BG)
      .setFontColor(ACTIVITY_QA_CONFIG.REPORT_HDR_TXT)
      .setFontWeight('bold')
      .setHorizontalAlignment('center');
  });
}

function writeKatanaActivityEmuInputRows_(sheet, rows) {
  if (!rows || rows.length === 0) return;
  var values = [];
  for (var i = 0; i < rows.length; i++) {
    values.push([
      rows[i].enabled ? 'TRUE' : 'FALSE',
      rows[i].flow || '',
      rows[i].type || '',
      rows[i].katanaId || '',
      rows[i].sourceOverride || '',
      rows[i].reverse ? 'TRUE' : 'FALSE',
      rows[i].scenario || '',
      rows[i].notes || '',
      rows[i].referenceHint || '',
      rows[i].statusHint || ''
    ]);
  }
  sheet.getRange(4, 1, values.length, values[0].length).setValues(values);
}

function readKatanaActivityEmuInputRows_(sheet) {
  if (!sheet || sheet.getLastRow() < KATANA_ACTIVITY_EMU_CONFIG.DATA_START_ROW) return [];
  var values = sheet.getRange(KATANA_ACTIVITY_EMU_CONFIG.DATA_START_ROW, 1, sheet.getLastRow() - 3, 10).getDisplayValues();
  var rows = [];
  for (var i = 0; i < values.length; i++) {
    var row = values[i];
    if (!String(row[1] || row[2] || row[3] || '').trim()) continue;
    rows.push({
      rowNum: i + KATANA_ACTIVITY_EMU_CONFIG.DATA_START_ROW,
      enabled: normalizeKatanaEmuBoolean_(row[0]),
      flow: String(row[1] || '').trim().toUpperCase(),
      type: String(row[2] || '').trim().toLowerCase(),
      katanaId: String(row[3] || '').trim(),
      sourceOverride: String(row[4] || '').trim(),
      reverse: normalizeKatanaEmuBoolean_(row[5]),
      scenario: String(row[6] || '').trim(),
      notes: String(row[7] || '').trim(),
      referenceHint: String(row[8] || '').trim(),
      statusHint: String(row[9] || '').trim()
    });
  }
  return rows;
}

function normalizeKatanaEmuBoolean_(value) {
  var text = String(value || '').trim().toLowerCase();
  return text === 'true' || text === 'yes' || text === 'y' || text === '1';
}

function buildDefaultKatanaActivityEmuRows_() {
  return [
    { enabled: true, flow: 'F4', type: 'mo', katanaId: '15770320', scenario: 'Sample MO build', notes: 'User-provided manufacturing order' },
    { enabled: true, flow: 'F1', type: 'po', katanaId: '2622022', scenario: 'Sample PO receive', notes: 'User-provided purchase order' },
    { enabled: true, flow: 'F2', type: 'sa', katanaId: '2241787', scenario: 'Sample SA adjust', notes: 'User-provided stock adjustment' },
    { enabled: true, flow: 'F6', type: 'so', katanaId: '41588374', scenario: 'Sample Amazon US stage', notes: 'User-provided sales order' },
    { enabled: true, flow: 'F3', type: 'st', katanaId: '302781', scenario: 'Sample transfer', notes: 'User-provided stock transfer' },
    { enabled: false, flow: 'F5', type: 'so', katanaId: '', sourceOverride: 'ShipStation', scenario: 'Fill with a Shopify or normal shipment SO', notes: 'ShipStation-backed flow preview' }
  ];
}

function buildRecentKatanaActivityEmuRows_(limitPerFlow) {
  var limit = Math.max(1, Math.min(Number(limitPerFlow || KATANA_ACTIVITY_EMU_CONFIG.DEFAULT_RECENT_PER_FLOW), 5));
  var rows = [];
  var poRows = fetchKatanaRecentObjects_('purchase_orders', limit * 4);
  var saRows = fetchKatanaRecentObjects_('stock_adjustments', limit * 4);
  var stRows = fetchKatanaRecentObjects_('stock_transfers', limit * 4);
  var soRows = fetchKatanaRecentObjects_('sales_orders', limit * 6);
  var moRows = fetchKatanaRecentEligibleObjects_('manufacturing_orders', limit, function(mo) {
    var status = String(mo.status || '').toUpperCase();
    var locationName = resolveKatanaEmuLocationName_(mo.location_id || '', {}, mo.location_name || mo.manufacturing_location_name || '');
    return !!mo.id && (status === 'DONE' || status === 'COMPLETED') && !isIgnoredKatanaMOLocation_(locationName);
  }, 8);

  pushKatanaEmuRows_(rows, 'F1', 'po', poRows, limit, function(po) {
    var status = String(po.status || '').toUpperCase();
    return status.indexOf('RECEIVED') >= 0;
  }, function(po) {
    return {
      scenario: 'Recent PO ' + (po.order_no || ('PO-' + po.id)),
      notes: 'Autopopulated from recent purchase orders',
      referenceHint: po.order_no || ''
    };
  });

  pushKatanaEmuRows_(rows, 'F4', 'mo', moRows, limit, null, function(mo) {
    return {
      scenario: 'Recent MO ' + (mo.order_no || ('MO-' + mo.id)),
      notes: 'Autopopulated from recent manufacturing orders',
      referenceHint: mo.order_no || '',
      statusHint: mo.status || ''
    };
  });

  pushKatanaEmuRows_(rows, 'F2', 'sa', saRows, limit, function(sa) {
    return !!sa.id;
  }, function(sa) {
    return {
      scenario: 'Recent SA ' + (sa.stock_adjustment_number || ('SA-' + sa.id)),
      notes: 'Autopopulated from recent stock adjustments',
      referenceHint: sa.stock_adjustment_number || '',
      statusHint: sa.reason || ''
    };
  });

  pushKatanaEmuRows_(rows, 'F3', 'st', stRows, limit, function(st) {
    return !!st.id;
  }, function(st) {
    return {
      scenario: 'Recent ST ' + (st.stock_transfer_number || ('ST-' + st.id)),
      notes: 'Autopopulated from recent stock transfers',
      referenceHint: st.stock_transfer_number || '',
      statusHint: st.status || ''
    };
  });

  pushRecentSalesOrderEmuRows_(rows, soRows, limit, 'F5');
  pushRecentSalesOrderEmuRows_(rows, soRows, limit, 'F6');
  return rows;
}

function pushKatanaEmuRows_(target, flow, type, records, limit, filterFn, mapFn) {
  var added = 0;
  for (var i = 0; i < (records || []).length && added < limit; i++) {
    var record = unwrapKatanaObject_(records[i]);
    if (!record || !record.id) continue;
    if (filterFn && !filterFn(record)) continue;
    var extra = mapFn ? mapFn(record) : {};
    target.push({
      enabled: true,
      flow: flow,
      type: type,
      katanaId: String(record.id),
      sourceOverride: extra.sourceOverride || '',
      reverse: !!extra.reverse,
      scenario: extra.scenario || (flow + ' ' + record.id),
      notes: extra.notes || '',
      referenceHint: extra.referenceHint || '',
      statusHint: extra.statusHint || ''
    });
    added++;
  }
}

function pushRecentSalesOrderEmuRows_(target, soRows, limit, desiredFlow) {
  var added = 0;
  for (var i = 0; i < (soRows || []).length && added < limit; i++) {
    var so = unwrapKatanaObject_(soRows[i]);
    if (!so || !so.id) continue;
    if (inferSalesOrderEmuFlow_(so) !== desiredFlow) continue;
    target.push({
      enabled: true,
      flow: desiredFlow,
      type: 'so',
      katanaId: String(so.id),
      sourceOverride: desiredFlow === 'F5' ? 'ShipStation' : '',
      reverse: false,
      scenario: 'Recent SO ' + (so.order_no || ('SO-' + so.id)),
      notes: desiredFlow === 'F5'
        ? 'Autopopulated from recent sales orders; ShipStation lots are not on Katana SO rows'
        : 'Autopopulated from recent sales orders',
      referenceHint: so.order_no || '',
      statusHint: so.status || ''
    });
    added++;
  }
}

function fetchKatanaRecentObjects_(endpoint, limit) {
  var perPage = Math.max(1, Math.min(Number(limit || 10), 100));
  var result = katanaApiCall(endpoint + '?per_page=' + perPage + '&sort=-updated_at');
  return unwrapKatanaArray_(result);
}

function fetchKatanaRecentEligibleObjects_(endpoint, desiredCount, filterFn, maxPages) {
  var wanted = Math.max(1, Number(desiredCount || 1));
  var pages = Math.max(1, Number(maxPages || 5));
  var collected = [];
  for (var page = 1; page <= pages && collected.length < wanted; page++) {
    var result = katanaApiCall(endpoint + '?per_page=100&page=' + page + '&sort=-updated_at');
    var rows = unwrapKatanaArray_(result);
    if (!rows || rows.length === 0) break;
    for (var i = 0; i < rows.length && collected.length < wanted; i++) {
      var record = unwrapKatanaObject_(rows[i]);
      if (!record || !record.id) continue;
      if (filterFn && !filterFn(record)) continue;
      collected.push(record);
    }
  }
  return collected;
}

function buildKatanaActivityFixtureFromInput_(entry, index) {
  var flow = String(entry.flow || '').trim().toUpperCase();
  var type = String(entry.type || resolveKatanaEmuTypeFromFlow_(flow)).trim().toLowerCase();
  var katanaId = String(entry.katanaId || '').trim();
  if (!flow || !FLOW_LABELS[flow]) throw new Error('Unknown flow on input row ' + entry.rowNum);
  if (!type) throw new Error('Missing Katana type on input row ' + entry.rowNum);
  if (!katanaId) throw new Error('Missing Katana ID on input row ' + entry.rowNum);

  if (flow === 'F1' && type === 'po') return buildKatanaPOActivityFixture_(entry, index);
  if (flow === 'F2' && type === 'sa') return buildKatanaSAActivityFixture_(entry, index);
  if (flow === 'F3' && type === 'st') return buildKatanaSTActivityFixture_(entry, index);
  if (flow === 'F4' && type === 'mo') return buildKatanaMOActivityFixture_(entry, index);
  if ((flow === 'F5' || flow === 'F6') && type === 'so') return buildKatanaSOActivityFixture_(entry, index);

  throw new Error('Flow/type combination not supported: ' + flow + ' / ' + type);
}

function resolveKatanaEmuTypeFromFlow_(flow) {
  if (flow === 'F1') return 'po';
  if (flow === 'F2') return 'sa';
  if (flow === 'F3') return 'st';
  if (flow === 'F4') return 'mo';
  if (flow === 'F5' || flow === 'F6') return 'so';
  return '';
}

function unwrapKatanaObject_(data) {
  return data && data.data ? data.data : data;
}

function unwrapKatanaArray_(data) {
  if (!data) return [];
  return data.data || data || [];
}

function toKatanaEmuNumber_(value) {
  var num = parseFloat(value || 0);
  return isNaN(num) ? 0 : num;
}

function getKatanaEmuVariant_(variantId, cache) {
  if (!variantId) return null;
  var key = String(variantId);
  if (cache && Object.prototype.hasOwnProperty.call(cache, key)) return cache[key];
  var variant = unwrapKatanaObject_(fetchKatanaVariant(variantId));
  if (cache) cache[key] = variant || null;
  return variant || null;
}

function getKatanaEmuLocation_(locationId, cache) {
  if (!locationId) return null;
  var key = String(locationId);
  if (cache && Object.prototype.hasOwnProperty.call(cache, key)) return cache[key];
  var location = unwrapKatanaObject_(fetchKatanaLocation(locationId));
  if (cache) cache[key] = location || null;
  return location || null;
}

function resolveKatanaEmuDestination_(locationId, locationCache) {
  var location = getKatanaEmuLocation_(locationId, locationCache);
  var katanaName = location ? String(location.name || '').trim() : '';
  var mapped = katanaName && KATANA_LOCATION_TO_WASP[katanaName]
    ? KATANA_LOCATION_TO_WASP[katanaName]
    : { site: CONFIG.WASP_SITE, location: FLOWS.PO_RECEIVING_LOCATION };
  return {
    site: mapped.site || CONFIG.WASP_SITE,
    location: mapped.location || FLOWS.PO_RECEIVING_LOCATION,
    katanaLocationName: katanaName
  };
}

function resolveKatanaEmuLocationName_(locationId, locationCache, fallbackName) {
  var location = getKatanaEmuLocation_(locationId, locationCache);
  return String((location && location.name) || fallbackName || '').trim();
}

function isIgnoredKatanaMOLocation_(locationName) {
  return String(locationName || '').trim().toLowerCase() === 'shopify';
}

function buildKatanaPurchaseUomNote_(row, variant, stockQty) {
  var rate = toKatanaEmuNumber_(row.purchase_uom_conversion_rate || 1) || 1;
  if (rate === 1) return '';
  var rawQty = toKatanaEmuNumber_(row.received_quantity || row.quantity || 0);
  var purchaseUom = String(row.purchase_uom || row.purchasing_uom || 'unit').trim();
  var stockUom = normalizeUom(resolveVariantUom(variant) || row.uom || '');
  return 'Katana ' + rawQty + ' ' + purchaseUom + ' -> ' + stockQty + (stockUom ? ' ' + stockUom : '');
}

function inferKatanaAdjustmentSource_(sa) {
  var sourceText = (String(sa.reason || '') + ' ' + String(sa.additional_info || '') + ' ' + String(sa.stock_adjustment_number || '')).toLowerCase();
  if (sourceText.indexOf('sheet') >= 0) return 'Sheets';
  if (sourceText.indexOf('wasp') >= 0) return 'Wasp';
  return 'Katana';
}

function inferSalesOrderEmuFlow_(so) {
  var location = '';
  if (so.location_id) {
    var soLoc = unwrapKatanaObject_(fetchKatanaLocation(so.location_id));
    location = soLoc ? String(soLoc.name || '').trim() : '';
  }
  if (location === FLOWS.AMAZON_FBA_KATANA_LOCATION) return 'F6';
  if (isAmazonUSSalesOrder_(so, null)) return 'F6';
  return 'F5';
}

function buildEmuSourceContext_(sourceLabel, siteName) {
  var source = String(sourceLabel || '').trim();
  var site = getActivityDisplaySite_(siteName);
  if (source && site) return source + ' @ ' + site;
  return source || site || '';
}

function buildEmuTransferContext_(sourceLabel, fromSite, toSite) {
  var source = String(sourceLabel || '').trim();
  var fromLabel = getActivityDisplaySite_(fromSite) || 'mixed-source';
  var toLabel = getActivityDisplaySite_(toSite) || 'mixed-dest';
  return source ? (source + ' | ' + fromLabel + ' -> ' + toLabel) : (fromLabel + ' -> ' + toLabel);
}

function buildEmuDirectionalMeta_(direction, location, lot, expiry, extraSegments) {
  var loc = getActivityDisplayLocation_(location);
  var action = '';
  if (loc) action = String(direction || '').trim() + ' ' + loc;
  var filteredExtras = [];
  for (var i = 0; i < (extraSegments || []).length; i++) {
    var extra = String(extraSegments[i] || '').trim();
    if (!extra || /^batch_id:/i.test(extra)) continue;
    filteredExtras.push(extra);
  }
  return joinActivitySegments_([buildActivityBatchMeta_(lot, expiry), action].concat(filteredExtras));
}

function buildEmuVerbAction_(verb, location, lot, expiry, extraSegments) {
  var loc = getActivityDisplayLocation_(location);
  var action = String(verb || '').trim();
  if (loc) action = action ? (action + ' ' + loc) : loc;
  var filteredExtras = [];
  for (var i = 0; i < (extraSegments || []).length; i++) {
    var extra = String(extraSegments[i] || '').trim();
    if (!extra || /^batch_id:/i.test(extra)) continue;
    filteredExtras.push(extra);
  }
  return joinActivitySegments_([buildActivityBatchMeta_(lot, expiry), action].concat(filteredExtras));
}

function resolveKatanaMOIngredientQty_(ingredient) {
  return toKatanaEmuNumber_(
    ingredient.total_consumed_quantity ||
    ingredient.consumed_quantity ||
    ingredient.actual_quantity ||
    ingredient.quantity || 0
  );
}

function resolveKatanaEmuVariantId_(node) {
  if (!node) return '';
  return String(
    node.variant_id ||
    node.product_variant_id ||
    node.material_variant_id ||
    node.item_variant_id ||
    ''
  ).trim();
}

function collectKatanaIngredientBatchRows_(ingredient) {
  var raw = ingredient && (ingredient.batch_transactions || ingredient.batchTransactions || []);
  var variantId = resolveKatanaEmuVariantId_(ingredient);
  var variantBatchCache = null;
  var rows = [];
  for (var i = 0; i < raw.length; i++) {
    var batch = raw[i] || {};
    var batchId = String(batch.batch_id || batch.batchId || batch.batch_stock_id || batch.batchStockId || batch.id || '').trim();
    var batchInfo = batchId ? fetchKatanaBatchStock(batchId) : null;
    var qty = toKatanaEmuNumber_(
      batch.quantity ||
      batch.qty ||
      batch.actual_quantity ||
      batch.consumed_quantity ||
      0
    );
    var lot = extractKatanaBatchNumber_(batch) ||
      extractKatanaBatchNumber_(batch.batch || null) ||
      extractKatanaBatchNumber_(batch.batch_stock || batch.batchStock || null) ||
      extractKatanaBatchNumber_(batchInfo);
    var expiry = extractKatanaExpiryDate_(batch) ||
      extractKatanaExpiryDate_(batch.batch || null) ||
      extractKatanaExpiryDate_(batch.batch_stock || batch.batchStock || null) ||
      extractKatanaExpiryDate_(batchInfo);
    if (batchId && (!lot || !expiry) && variantId) {
      if (!variantBatchCache) {
        var variantBatchResult = katanaApiCall('batch_stocks?variant_id=' + variantId + '&include_deleted=true');
        variantBatchCache = unwrapKatanaArray_(variantBatchResult);
      }
      for (var vb = 0; vb < variantBatchCache.length; vb++) {
        var vbRow = variantBatchCache[vb] || {};
        var vbId = String(vbRow.batch_id || vbRow.batchId || vbRow.id || '').trim();
        if (!vbId || vbId !== batchId) continue;
        lot = lot || extractKatanaBatchNumber_(vbRow);
        expiry = expiry || extractKatanaExpiryDate_(vbRow);
        break;
      }
    }
    if (!qty && !lot && !expiry && !batchId) continue;
    rows.push({
      qty: qty,
      lot: lot || '',
      expiry: expiry || '',
      batchId: batchId || ''
    });
  }
  return rows;
}

function resolveKatanaEmuRowUom_(row, variant) {
  return normalizeUom(
    resolveVariantUom(variant) ||
    row.uom ||
    row.unit ||
    row.sales_uom ||
    row.selling_uom ||
    row.purchase_uom ||
    row.product_uom ||
    row.material_uom ||
    ''
  );
}

function buildKatanaPOActivityFixture_(entry, index) {
  var poId = entry.katanaId;
  var po = unwrapKatanaObject_(fetchKatanaPO(poId));
  if (!po) throw new Error('Could not fetch Katana PO ' + poId);
  var rows = unwrapKatanaArray_(fetchKatanaPORows(poId));
  var variantCache = {};
  var locationCache = {};
  var dest = resolveKatanaEmuDestination_(po.location_id || '', locationCache);
  var poStatus = String(po.status || '').toUpperCase();
  var isPartial = poStatus.indexOf('PARTIAL') >= 0;
  var isReverse = !!entry.reverse;
  var subItems = [];
  var topLevelCount = 0;

  for (var i = 0; i < rows.length; i++) {
    var row = rows[i];
    var variant = getKatanaEmuVariant_(row.variant_id, variantCache);
    var sku = variant ? String(variant.sku || '').trim() : '';
    if (!sku) continue;

    var qtyBase = toKatanaEmuNumber_(row.received_quantity || row.received_qty || row.quantity || 0);
    var puomRate = toKatanaEmuNumber_(row.purchase_uom_conversion_rate || 1) || 1;
    var qty = qtyBase * puomRate;
    if (qty <= 0) continue;

    var itemLocation = dest.location;
    if (dest.site === CONFIG.WASP_SITE) {
      var itemType = getKatanaItemType(variant);
      if (itemType === 'material' || itemType === 'intermediate') itemLocation = LOCATIONS.PRODUCTION;
    }

    var note = buildKatanaPurchaseUomNote_(row, variant, qty);
    var batches = row.batch_transactions || row.batchTransactions || [];
    if (batches.length > 1) {
      topLevelCount++;
      subItems.push({
        sku: sku,
        qty: qty,
        uom: normalizeUom(resolveVariantUom(variant)),
        action: buildEmuDirectionalMeta_(isReverse ? '<-' : '->', itemLocation, '', '', [note ? '(' + note + ')' : '']),
        status: '',
        qtyColor: 'grey',
        isParent: true,
        batchCount: batches.length
      });
      for (var b = 0; b < batches.length; b++) {
        var batch = batches[b];
        var batchQty = toKatanaEmuNumber_(batch.quantity || 0);
        if (batchQty <= 0) continue;
        var batchInfo = batch.batch_id ? fetchKatanaBatchStock(batch.batch_id) : null;
        var lot = extractKatanaBatchNumber_(batch) || extractKatanaBatchNumber_(batch.batch_stock || null) || extractKatanaBatchNumber_(batchInfo);
        var expiry = extractKatanaExpiryDate_(batch) || extractKatanaExpiryDate_(batch.batch_stock || null) || extractKatanaExpiryDate_(batchInfo);
        subItems.push({
          nested: true,
          qty: batchQty,
          uom: normalizeUom(resolveVariantUom(variant)),
          action: buildActivityBatchMeta_(lot, expiry),
          status: isReverse ? 'Removed' : 'Added',
          qtyColor: isReverse ? 'red' : 'green'
        });
      }
    } else {
      topLevelCount++;
      subItems.push({
        sku: sku,
        qty: qty,
        uom: normalizeUom(resolveVariantUom(variant)),
        action: buildEmuDirectionalMeta_(isReverse ? '<-' : '->', itemLocation, extractIngredientBatchNumber(row), extractIngredientExpiryDate(row), [note ? '(' + note + ')' : '']),
        status: isReverse ? 'Removed' : 'Added',
        qtyColor: isReverse ? 'red' : 'green'
      });
    }
  }

  if (subItems.length === 0) throw new Error('PO ' + poId + ' has no previewable received rows');

  var reference = extractCanonicalActivityRef_(po.order_no, 'PO-', poId);
  return {
    fixture: {
      flow: 'F1',
      scenario: entry.scenario || ('PO ' + poId),
      reference: reference,
      summary: buildActivityCountSummary_(topLevelCount, 'line', 'lines', isReverse ? 'reverted' : 'received'),
      context: buildEmuSourceContext_('Katana', dest.site),
      status: isReverse ? 'reverted' : (isPartial ? 'partial' : 'success'),
      linkInfo: buildActivityFixtureLink_('po', poId, reference),
      subItems: subItems
    },
    notes: [
      'Katana location maps to Wasp site ' + dest.site,
      'Wasp bin only appears when it differs from the site name'
    ]
  };
}

function buildKatanaSAActivityFixture_(entry, index) {
  var saId = entry.katanaId;
  var sa = unwrapKatanaObject_(fetchKatanaStockAdjustment(saId));
  if (!sa) throw new Error('Could not fetch Katana SA ' + saId);
  var rows = sa.stock_adjustment_rows || sa.rows || unwrapKatanaArray_(fetchKatanaStockAdjustmentRows(saId));
  var variantCache = {};
  var locationCache = {};
  var dest = resolveKatanaEmuDestination_(sa.location_id || '', locationCache);
  var source = entry.sourceOverride || inferKatanaAdjustmentSource_(sa);
  var isReverse = !!entry.reverse;
  var subItems = [];
  var topLevelCount = 0;

  for (var i = 0; i < rows.length; i++) {
    var row = rows[i];
    var variant = getKatanaEmuVariant_(row.variant_id, variantCache);
    var sku = variant ? String(variant.sku || '').trim() : '';
    var rawQty = toKatanaEmuNumber_(row.quantity || 0);
    if (!sku || rawQty === 0) continue;
    topLevelCount++;
    var rowAdds = rawQty > 0;
    var direction = rowAdds ? '->' : '<-';
    if (isReverse) direction = rowAdds ? '<-' : '->';
    subItems.push({
      sku: sku,
      qty: Math.abs(rawQty),
      uom: normalizeUom(resolveVariantUom(variant)),
      action: buildEmuDirectionalMeta_(direction, dest.location || dest.site, extractIngredientBatchNumber(row), extractIngredientExpiryDate(row)),
      status: isReverse ? (rowAdds ? 'Removed' : 'Restored') : 'Adjusted',
      qtyColor: isReverse ? (rowAdds ? 'red' : 'green') : (rowAdds ? 'green' : 'red')
    });
  }

  if (subItems.length === 0) throw new Error('SA ' + saId + ' has no previewable adjustment rows');

  var reference = extractCanonicalActivityRef_(sa.stock_adjustment_number, 'SA-', saId);
  var firstItem = findFirstTopLevelPreviewItem_(subItems);
  return {
    fixture: {
      flow: 'F2',
      scenario: entry.scenario || ('SA ' + saId),
      reference: reference,
      summary: topLevelCount === 1 && firstItem ? (firstItem.sku + ' x' + firstItem.qty) : buildActivityCountSummary_(topLevelCount, 'adjustment', 'adjustments', isReverse ? 'reversed' : 'posted'),
      context: buildEmuSourceContext_(isReverse ? 'Katana' : source, dest.site),
      status: isReverse ? 'reverted' : 'success',
      linkInfo: buildActivityFixtureLink_('sa', saId, reference),
      subItems: subItems
    },
    notes: [
      'Katana SA does not carry Wasp bin detail, so line actions keep site plus batch context only'
    ]
  };
}

function buildKatanaSTActivityFixture_(entry, index) {
  var stId = entry.katanaId;
  var transfer = unwrapKatanaObject_(fetchKatanaStockTransfer(stId));
  if (!transfer) throw new Error('Could not fetch Katana stock transfer ' + stId);
  var rows = transfer.stock_transfer_rows || transfer.rows || [];
  if (!rows || rows.length === 0) throw new Error('ST ' + stId + ' returned no stock transfer rows');

  var fromLoc = unwrapKatanaObject_(fetchKatanaLocation(transfer.source_location_id || ''));
  var toLoc = unwrapKatanaObject_(fetchKatanaLocation(transfer.target_location_id || ''));
  var fromWasp = mapKatanaToWaspSite(fromLoc ? fromLoc.name : '') || { site: fromLoc ? fromLoc.name : 'mixed-source', location: '' };
  var toWasp = mapKatanaToWaspSite(toLoc ? toLoc.name : '') || { site: toLoc ? toLoc.name : 'mixed-dest', location: '' };
  var variantCache = {};
  var isReverse = !!entry.reverse || /cancel|void|reject/i.test(String(transfer.status || ''));
  var isPartial = /partial/i.test(String(transfer.status || ''));
  var subItems = [];
  var topLevelCount = 0;

  for (var i = 0; i < rows.length; i++) {
    var row = rows[i];
    var variant = getKatanaEmuVariant_(row.variant_id, variantCache);
    var sku = variant ? String(variant.sku || '').trim() : '';
    var qty = toKatanaEmuNumber_(row.quantity || 0);
    if (!sku || qty <= 0) continue;
    var batches = row.batch_transactions || row.batchTransactions || [];
    if (batches.length > 1) {
      topLevelCount++;
      subItems.push({
        sku: sku,
        qty: qty,
        uom: normalizeUom(resolveVariantUom(variant)),
        action: '',
        status: '',
        qtyColor: 'grey',
        isParent: true,
        batchCount: batches.length
      });
      for (var b = 0; b < batches.length; b++) {
        var batch = batches[b];
        var batchInfo = batch.batch_id ? fetchKatanaBatchForF3(batch.batch_id) : null;
        subItems.push({
          nested: true,
          qty: toKatanaEmuNumber_(batch.quantity || 0),
          uom: normalizeUom(resolveVariantUom(variant)),
          action: buildActivityBatchMeta_(extractKatanaBatchNumber_(batch) || extractKatanaBatchNumber_(batchInfo), extractKatanaExpiryDate_(batch) || extractKatanaExpiryDate_(batchInfo)),
          status: isReverse ? 'Reversed' : 'Transferred',
          qtyColor: 'green'
        });
      }
    } else {
      topLevelCount++;
      subItems.push({
        sku: sku,
        qty: qty,
        uom: normalizeUom(resolveVariantUom(variant)),
        action: buildActivityBatchMeta_(extractIngredientBatchNumber(row), extractIngredientExpiryDate(row)),
        status: isReverse ? 'Reversed' : 'Transferred',
        qtyColor: 'green'
      });
    }
  }

  if (subItems.length === 0) throw new Error('ST ' + stId + ' has no previewable transfer rows');

  var reference = extractCanonicalActivityRef_(transfer.stock_transfer_number, 'ST-', stId);
  return {
    fixture: {
      flow: 'F3',
      scenario: entry.scenario || ('ST ' + stId),
      reference: reference,
      summary: buildActivityCountSummary_(topLevelCount, 'line', 'lines', isReverse ? 'reversed' : 'moved'),
      context: buildEmuTransferContext_('Katana poll', fromWasp.site, toWasp.site),
      status: isReverse ? 'reverted' : (isPartial ? 'partial' : 'success'),
      linkInfo: buildActivityFixtureLink_('st', stId, reference),
      subItems: subItems
    },
    notes: [
      'Katana transfer source/destination map to Wasp sites ' + fromWasp.site + ' -> ' + toWasp.site
    ]
  };
}

function buildKatanaMOActivityFixture_(entry, index) {
  var moId = entry.katanaId;
  var mo = unwrapKatanaObject_(fetchKatanaMO(moId));
  if (!mo) throw new Error('Could not fetch Katana MO ' + moId);
  var variantCache = {};
  var locationCache = {};
  var dest = resolveKatanaEmuDestination_(mo.location_id || '', locationCache);
  var moLocationName = resolveKatanaEmuLocationName_(mo.location_id || '', locationCache, mo.location_name || mo.manufacturing_location_name || '');
  if (isIgnoredKatanaMOLocation_(moLocationName)) {
    throw new Error('MO ' + moId + ' uses Shopify manufacturing location and is excluded from F4');
  }
  var ingredients = unwrapKatanaArray_(fetchKatanaMOIngredients(moId));
  var isReverse = !!entry.reverse;
  var moStatus = String(mo.status || '').toUpperCase();
  var subItems = [];
  var ingredientCount = 0;

  for (var i = 0; i < ingredients.length; i++) {
    var ing = ingredients[i];
    var variant = getKatanaEmuVariant_(resolveKatanaEmuVariantId_(ing), variantCache);
    var sku = variant ? String(variant.sku || '').trim() : String(ing.sku || '').trim();
    var batchRows = collectKatanaIngredientBatchRows_(ing);
    var qty = resolveKatanaMOIngredientQty_(ing);
    if (qty <= 0 && batchRows.length > 0) {
      for (var br = 0; br < batchRows.length; br++) {
        qty += toKatanaEmuNumber_(batchRows[br].qty || 0);
      }
    }
    if (!sku || qty <= 0) continue;

    var ingUom = normalizeUom(resolveVariantUom(variant));
    ingredientCount++;
    if (batchRows.length > 1) {
      subItems.push({
        sku: sku,
        qty: qty,
        uom: ingUom,
        action: buildEmuVerbAction_(isReverse ? 'restore to' : 'consume from', FLOWS.MO_INGREDIENT_LOCATION, '', '', []),
        status: '',
        qtyColor: isReverse ? 'green' : 'red',
        isParent: true,
        batchCount: batchRows.length
      });
      for (var b = 0; b < batchRows.length; b++) {
        var batch = batchRows[b];
        subItems.push({
          nested: true,
          qty: toKatanaEmuNumber_(batch.qty || 0),
          uom: ingUom,
          action: buildActivityBatchMeta_(batch.lot, batch.expiry),
          status: isReverse ? 'Restored' : 'Consumed',
          qtyColor: isReverse ? 'green' : 'red'
        });
      }
    } else {
      var ingLot = batchRows.length === 1 ? (batchRows[0].lot || '') : (extractIngredientBatchNumber(ing) || '');
      var ingExpiry = batchRows.length === 1 ? (batchRows[0].expiry || '') : (extractIngredientExpiryDate(ing) || '');
      var batchId = batchRows.length === 1 ? (batchRows[0].batchId || '') : '';
      subItems.push({
        sku: sku,
        qty: qty,
        uom: ingUom,
        action: buildEmuVerbAction_(isReverse ? 'restore to' : 'consume from', FLOWS.MO_INGREDIENT_LOCATION, ingLot, ingExpiry, [batchId && !ingLot ? 'batch_id:' + batchId : '']),
        status: isReverse ? 'Restored' : 'Consumed',
        qtyColor: isReverse ? 'green' : 'red'
      });
    }
  }

  var outputVariant = getKatanaEmuVariant_(mo.variant_id, variantCache);
  var outputSku = outputVariant ? String(outputVariant.sku || '').trim() : '';
  var outputQty = toKatanaEmuNumber_(mo.actual_quantity || mo.quantity || mo.planned_quantity || 0);
  var outputBatchData = typeof resolveMOOutputBatchData_ === 'function'
    ? resolveMOOutputBatchData_(mo, mo.variant_id, mo.order_no || '')
    : null;
  var outputLot = outputBatchData ? (outputBatchData.lot || '') : (extractKatanaBatchNumber_(mo) || '');
  var outputExpiry = outputBatchData ? (outputBatchData.expiry || '') : (extractKatanaExpiryDate_(mo) || '');
  if (outputSku && outputQty > 0) {
    subItems.push({
      sku: outputSku,
      qty: outputQty,
      uom: normalizeUom(resolveVariantUom(outputVariant)),
      action: buildEmuVerbAction_(isReverse ? 'remove from' : 'produce into', FLOWS.MO_OUTPUT_LOCATION, outputLot, outputExpiry),
      status: isReverse ? 'Removed' : 'Produced',
      qtyColor: isReverse ? 'red' : 'green'
    });
  }

  if (subItems.length === 0) throw new Error('MO ' + moId + ' has no previewable ingredient or output rows');

  var reference = extractCanonicalActivityRef_(mo.order_no, 'MO-', moId);
  return {
    fixture: {
      flow: 'F4',
      scenario: entry.scenario || ('MO ' + moId),
      reference: reference,
      summary: isReverse
        ? buildActivityCountSummary_(ingredientCount || 1, 'ingredient', 'ingredients', 'restored')
        : (outputSku ? (outputSku + ' x' + outputQty) : buildActivityCountSummary_(ingredientCount, 'line', 'lines', 'built')),
      context: isReverse
        ? '<- ' + getActivityDisplayLocation_(FLOWS.MO_INGREDIENT_LOCATION)
        : '-> ' + getActivityDisplayLocation_(FLOWS.MO_OUTPUT_LOCATION),
      status: isReverse ? 'reverted' : (moStatus === 'DONE' ? 'success' : 'partial'),
      linkInfo: buildActivityFixtureLink_('mo', moId, reference),
      subItems: subItems
    },
    notes: [
      'Katana location maps to Wasp site ' + dest.site
    ]
  };
}

function buildKatanaSOActivityFixture_(entry, index) {
  var soId = entry.katanaId;
  var so = unwrapKatanaObject_(fetchKatanaSalesOrder(soId));
  if (!so) throw new Error('Could not fetch Katana SO ' + soId);
  var flow = entry.flow || inferSalesOrderEmuFlow_(so);
  var rows = so.sales_order_rows || so.rows || unwrapKatanaArray_(fetchKatanaSalesOrderRows(soId));
  if (!rows || rows.length === 0) throw new Error('SO ' + soId + ' has no previewable sales rows');

  var variantCache = {};
  var locationCache = {};
  var soLoc = resolveKatanaEmuDestination_(so.location_id || '', locationCache);
  var isReverse = !!entry.reverse;
  var subItems = [];
  var topLevelCount = 0;

  for (var i = 0; i < rows.length; i++) {
    var row = rows[i];
    var variant = getKatanaEmuVariant_(row.variant_id, variantCache);
    var sku = variant ? String(variant.sku || '').trim() : '';
    var qty = toKatanaEmuNumber_(row.delivered_quantity || row.quantity || 0);
    if (!sku || qty <= 0) continue;
    topLevelCount++;
    if (flow === 'F6') {
      subItems.push({
        sku: sku,
        qty: qty,
        uom: resolveKatanaEmuRowUom_(row, variant),
        action: buildEmuDirectionalMeta_(isReverse ? '->' : '->', isReverse ? FLOWS.AMAZON_TRANSFER_LOCATION : FLOWS.AMAZON_FBA_WASP_LOCATION, '', ''),
        status: isReverse ? 'Restored' : 'Complete',
        qtyColor: 'green'
      });
    } else {
      subItems.push({
        sku: sku,
        qty: qty,
        uom: resolveKatanaEmuRowUom_(row, variant),
        action: buildEmuDirectionalMeta_(isReverse ? '->' : '<-', FLOWS.PICK_FROM_LOCATION, '', ''),
        status: isReverse ? 'Returned' : 'Deducted',
        qtyColor: isReverse ? 'green' : 'red'
      });
    }
  }

  if (subItems.length === 0) throw new Error('SO ' + soId + ' has no previewable shipped rows');

  var reference = extractCanonicalActivityRef_(so.order_no, 'SO-', soId);
  var notes = [];
  if (flow === 'F5') notes.push('ShipStation lot selection is not stored on Katana SO rows, so lot metadata is omitted in preview');
  if (flow === 'F6') notes.push('Exact lot resolution for FBA staging comes from WASP at runtime, so preview uses site/bin only');

  return {
    fixture: {
      flow: flow,
      scenario: entry.scenario || ('SO ' + soId),
      reference: reference,
      summary: buildActivityCountSummary_(topLevelCount, 'line', 'lines', isReverse ? (flow === 'F6' ? 'reversed' : 'returned') : (flow === 'F6' ? 'staged' : 'shipped')),
      context: flow === 'F6'
        ? buildEmuTransferContext_(isReverse ? 'Katana' : 'ShipStation', isReverse ? FLOWS.AMAZON_FBA_WASP_SITE : CONFIG.WASP_SITE, isReverse ? CONFIG.WASP_SITE : FLOWS.AMAZON_FBA_WASP_SITE)
        : buildEmuSourceContext_(isReverse ? 'Katana' : (entry.sourceOverride || 'ShipStation'), soLoc.site),
      status: isReverse ? 'reverted' : 'success',
      linkInfo: buildActivityFixtureLink_('so', soId, reference),
      subItems: subItems
    },
    notes: notes
  };
}

function writeKatanaActivityEmulationReport_(sheet, rows, renderedCount) {
  var okCount = 0;
  for (var i = 0; i < rows.length; i++) {
    if (!rows[i].findings || rows[i].findings.length === 0) okCount++;
  }

  sheet.getRange(3, 1, 1, 8).setValues([[
    'Summary',
    '',
    renderedCount + ' previews',
    rows.length ? Math.round((okCount / rows.length) * 100) : 0,
    '',
    '',
    '',
    'Preview sheet: ' + KATANA_ACTIVITY_EMU_CONFIG.PREVIEW_SHEET_NAME
  ]]);
  sheet.getRange(3, 1, 1, 8)
    .setBackground('#DDE7F3')
    .setFontWeight('bold');

  if (!rows || rows.length === 0) return;

  var reportRows = [];
  for (var r = 0; r < rows.length; r++) reportRows.push(formatActivityQAReportRow_(rows[r]));
  sheet.getRange(5, 1, reportRows.length, 8).setValues(reportRows);

  for (var x = 0; x < reportRows.length; x++) {
    var bg = reportRows[x][3] >= 90 ? ACTIVITY_QA_CONFIG.OK_BG : reportRows[x][3] >= 70 ? ACTIVITY_QA_CONFIG.WARN_BG : ACTIVITY_QA_CONFIG.FAIL_BG;
    sheet.getRange(x + 5, 4).setBackground(bg);
    if (reportRows[x][4]) sheet.getRange(x + 5, 5).setWrap(true);
    if (reportRows[x][7]) sheet.getRange(x + 5, 8).setWrap(true);
  }
}
