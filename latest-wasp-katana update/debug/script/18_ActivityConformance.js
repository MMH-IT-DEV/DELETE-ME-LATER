// ============================================
// 18_ActivityConformance.gs - ACTIVITY QA HARNESS
// ============================================
// Safe conformance tooling for the Activity tab.
// - Renders canonical reference scenarios to a scratch sheet
// - Audits the latest live Activity block for each flow
// - Produces a pass-rate report without mutating inventory
// ============================================

var ACTIVITY_QA_CONFIG = {
  PREVIEW_SHEET_NAME: 'Activity QA',
  REPORT_SHEET_NAME: 'Activity QA Report',
  REAL_PREVIEW_SHEET_NAME: 'Activity QA Real',
  REAL_REPORT_SHEET_NAME: 'Activity QA Real Report',
  HEADER_BG: '#1C2333',
  TITLE_BG: '#111E2D',
  TITLE_TXT: '#E7EEF8',
  SUBTITLE_TXT: '#AABBCC',
  REPORT_HDR_BG: '#2F4F6F',
  REPORT_HDR_TXT: '#FFFFFF',
  OK_BG: '#E8F5E9',
  WARN_BG: '#FFF8E1',
  FAIL_BG: '#FDECEC',
  DATA_START_ROW: 4,
  LIVE_AUDIT_ROW_WINDOW: 1200,
  REAL_SAMPLE_LIMIT_PER_FLOW: 2,
  WEBHOOK_AUDIT_ROW_WINDOW: 250
};

function runActivityConformanceSuite() {
  var ss = typeof getDebugSpreadsheet_ === 'function'
    ? getDebugSpreadsheet_()
    : SpreadsheetApp.openById(CONFIG.DEBUG_SHEET_ID);

  var previewSheet = ensureActivityQAPreviewSheet_(ss);
  var reportSheet = ensureActivityQAReportSheet_(ss);

  resetActivityQAPreviewSheet_(previewSheet);
  resetActivityQAReportSheet_(reportSheet);

  var fixtures = buildActivityConformanceFixtures_();
  var fixtureResults = [];

  for (var i = 0; i < fixtures.length; i++) {
    var fixture = fixtures[i];
    var execId = 'QA-' + padActivityQANumber_(i + 1);
    var startRow = previewSheet.getLastRow() + 1;
    var findings = validateActivityFixture_(fixture);

    renderActivityFixtureToSheet_(previewSheet, fixture, execId);
    findings = findings.concat(validateLiveActivityBlock_(previewSheet, readActivityBlockAtRow_(previewSheet, startRow, fixture.subItems || [])));
    fixtureResults.push({
      scope: 'Fixture',
      flow: fixture.flow,
      scenario: fixture.scenario,
      ref: fixture.reference,
      score: findings.length === 0 ? 100 : Math.max(0, 100 - findings.length * 20),
      findings: findings
    });
  }

  var liveAuditResults = auditLatestActivityBlocks_();
  writeActivityQAReport_(reportSheet, fixtureResults, liveAuditResults);

  var fixturePassRate = calculateActivitySuitePassRate_(fixtureResults);
  var livePassRate = calculateActivitySuitePassRate_(liveAuditResults);
  Logger.log(
    'Activity conformance suite complete | fixtures=' + fixturePassRate + '% | live=' + livePassRate + '%'
  );

  return {
    status: 'ok',
    fixturePassRate: fixturePassRate,
    livePassRate: livePassRate,
    previewSheet: ACTIVITY_QA_CONFIG.PREVIEW_SHEET_NAME,
    reportSheet: ACTIVITY_QA_CONFIG.REPORT_SHEET_NAME
  };
}

function runRealActivitySamples() {
  var ss = typeof getDebugSpreadsheet_ === 'function'
    ? getDebugSpreadsheet_()
    : SpreadsheetApp.openById(CONFIG.DEBUG_SHEET_ID);
  var activitySheet = ss.getSheetByName('Activity');
  if (!activitySheet || activitySheet.getLastRow() < ACTIVITY_QA_CONFIG.DATA_START_ROW) {
    throw new Error('Activity tab not found or has no data rows.');
  }

  var previewSheet = ensureActivityQARealPreviewSheet_(ss);
  var reportSheet = ensureActivityQARealReportSheet_(ss);
  resetActivityQARealPreviewSheet_(previewSheet);
  resetActivityQARealReportSheet_(reportSheet);

  var blocks = readRecentActivityBlocksByFlow_(ACTIVITY_QA_CONFIG.REAL_SAMPLE_LIMIT_PER_FLOW);
  var reportRows = [];
  var flowCounts = {};

  for (var i = 0; i < blocks.length; i++) {
    var block = blocks[i];
    flowCounts[block.flow] = (flowCounts[block.flow] || 0) + 1;
    var fixture = buildRealActivityFixtureFromBlock_(activitySheet, block, i);
    if (!fixture) {
      reportRows.push({
        scope: 'Real',
        flow: block.flow,
        scenario: 'sample ' + flowCounts[block.flow],
        score: 0,
        findings: ['Could not rebuild this Activity block into a QA preview'],
        ref: getActivityReferenceToken_(block.header),
        rowNum: block.headerRow,
        notes: describeRealSampleSource_(block) + ' | source Activity row ' + block.headerRow
      });
      continue;
    }

    var startRow = previewSheet.getLastRow() + 1;
    var findings = validateActivityFixture_(fixture);
    renderActivityFixtureToSheet_(previewSheet, fixture, 'QR-' + padActivityQANumber_(i + 1));
    findings = findings.concat(validateLiveActivityBlock_(previewSheet, readActivityBlockAtRow_(previewSheet, startRow, fixture.subItems || [])));
    reportRows.push({
      scope: 'Real',
      flow: block.flow,
      scenario: 'sample ' + flowCounts[block.flow],
      score: findings.length === 0 ? 100 : Math.max(0, 100 - findings.length * 15),
      findings: findings,
      ref: fixture.reference,
      rowNum: block.headerRow,
      notes: describeRealSampleSource_(block) + ' | rebuilt from Activity row ' + block.headerRow
    });
  }

  var flowOrder = ['F1', 'F2', 'F3', 'F4', 'F5', 'F6'];
  for (var f = 0; f < flowOrder.length; f++) {
    var flow = flowOrder[f];
    var seen = flowCounts[flow] || 0;
    while (seen < ACTIVITY_QA_CONFIG.REAL_SAMPLE_LIMIT_PER_FLOW) {
      seen++;
      reportRows.push({
        scope: 'Real',
        flow: flow,
        scenario: 'sample ' + seen,
        score: 0,
        findings: ['No recent real Activity block found for this sample slot'],
        ref: '',
        rowNum: '',
        notes: describeRealSampleSource_({ flow: flow })
      });
    }
  }

  writeActivityRealSampleReport_(reportSheet, reportRows);

  return {
    status: 'ok',
    sampleCount: reportRows.length,
    previewSheet: ACTIVITY_QA_CONFIG.REAL_PREVIEW_SHEET_NAME,
    reportSheet: ACTIVITY_QA_CONFIG.REAL_REPORT_SHEET_NAME
  };
}

function ensureActivityQAPreviewSheet_(ss) {
  var sheet = ss.getSheetByName(ACTIVITY_QA_CONFIG.PREVIEW_SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(ACTIVITY_QA_CONFIG.PREVIEW_SHEET_NAME);
  return sheet;
}

function ensureActivityQAReportSheet_(ss) {
  var sheet = ss.getSheetByName(ACTIVITY_QA_CONFIG.REPORT_SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(ACTIVITY_QA_CONFIG.REPORT_SHEET_NAME);
  return sheet;
}

function ensureActivityQARealPreviewSheet_(ss) {
  var sheet = ss.getSheetByName(ACTIVITY_QA_CONFIG.REAL_PREVIEW_SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(ACTIVITY_QA_CONFIG.REAL_PREVIEW_SHEET_NAME);
  return sheet;
}

function ensureActivityQARealReportSheet_(ss) {
  var sheet = ss.getSheetByName(ACTIVITY_QA_CONFIG.REAL_REPORT_SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(ACTIVITY_QA_CONFIG.REAL_REPORT_SHEET_NAME);
  return sheet;
}

function runSpreadsheetOpWithRetry_(label, fn) {
  var lastErr = null;
  for (var attempt = 0; attempt < 4; attempt++) {
    try {
      return fn();
    } catch (err) {
      lastErr = err;
      var message = String(err && err.message || err || '');
      var isTimeout = /Service Spreadsheets timed out|timed out while accessing document/i.test(message);
      if (!isTimeout || attempt === 3) throw err;
      Utilities.sleep((attempt + 1) * 1500);
    }
  }
  if (lastErr) throw lastErr;
}

function ensureSheetGridSize_(sheet, minRows, minCols) {
  runSpreadsheetOpWithRetry_('ensureSheetGridSize', function() {
    if (sheet.getMaxRows() < minRows) {
      sheet.insertRowsAfter(sheet.getMaxRows(), minRows - sheet.getMaxRows());
    }
    if (sheet.getMaxColumns() < minCols) {
      sheet.insertColumnsAfter(sheet.getMaxColumns(), minCols - sheet.getMaxColumns());
    }
  });
}

function clearActivityQASheet_(sheet, minCols) {
  var lastRow = sheet.getLastRow();
  var lastCol = Math.max(sheet.getLastColumn(), minCols || 1);
  if (lastRow > 0 && lastCol > 0) {
    sheet.getRange(1, 1, lastRow, lastCol).clear();
  }
}

function resetActivityQAPreviewSheet_(sheet) {
  runSpreadsheetOpWithRetry_('resetActivityQAPreviewSheet', function() {
    sheet.getRange(1, 1, 3, 15).breakApart();
    clearActivityQASheet_(sheet, 15);

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

    sheet.getRange(1, 1).setValue('ACTIVITY QA PREVIEW')
      .setBackground(ACTIVITY_QA_CONFIG.TITLE_BG)
      .setFontColor(ACTIVITY_QA_CONFIG.TITLE_TXT)
      .setFontWeight('bold')
      .setFontSize(12)
      .setHorizontalAlignment('center');

    sheet.getRange(2, 1).setValue('Canonical flow scenarios rendered through the shared Activity renderer. Rows 4+ are safe preview output only.')
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

    sheet.getRange(1, 8, 2, 8).setBackground(ACTIVITY_QA_CONFIG.TITLE_BG).setFontColor('#63758A');
    sheet.getRange(1, 8).setValue('Reserved');
    sheet.getRange(2, 8).setValue('H:O kept clear to mirror the Activity standard reservation.');

    sheet.getRange(1, 1, 3, 15).setVerticalAlignment('middle');
  });
}

function resetActivityQAReportSheet_(sheet) {
  runSpreadsheetOpWithRetry_('resetActivityQAReportSheet', function() {
    sheet.getRange(1, 1, 2, 8).breakApart();
    clearActivityQASheet_(sheet, 8);

    sheet.setFrozenRows(4);
    sheet.setRowHeight(1, 26);
    sheet.setRowHeight(2, 24);
    sheet.setRowHeight(4, 24);

    var reportWidths = { 1: 70, 2: 80, 3: 190, 4: 80, 5: 360, 6: 110, 7: 120, 8: 180 };
    for (var col in reportWidths) {
      sheet.setColumnWidth(Number(col), reportWidths[col]);
    }

    sheet.getRange(1, 1, 1, 8).merge();
    sheet.getRange(2, 1, 1, 8).merge();

    sheet.getRange(1, 1).setValue('ACTIVITY CONFORMANCE REPORT')
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

function resetActivityQARealPreviewSheet_(sheet) {
  runSpreadsheetOpWithRetry_('resetActivityQARealPreviewSheet', function() {
    sheet.getRange(1, 1, 3, 15).breakApart();
    clearActivityQASheet_(sheet, 15);

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

    sheet.getRange(1, 1).setValue('ACTIVITY QA REAL SAMPLES')
      .setBackground(ACTIVITY_QA_CONFIG.TITLE_BG)
      .setFontColor(ACTIVITY_QA_CONFIG.TITLE_TXT)
      .setFontWeight('bold')
      .setFontSize(12)
      .setHorizontalAlignment('center');

    sheet.getRange(2, 1).setValue('Last ' + ACTIVITY_QA_CONFIG.REAL_SAMPLE_LIMIT_PER_FLOW + ' real processed Activity blocks per flow. Rebuilt through the current Activity renderer for side-effect-free review.')
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

    sheet.getRange(1, 8, 2, 8).setBackground(ACTIVITY_QA_CONFIG.TITLE_BG).setFontColor('#63758A');
    sheet.getRange(1, 8).setValue('Replay mode');
    sheet.getRange(2, 8).setValue('Preview rows are regenerated from live Activity samples, not copied directly.');
    sheet.getRange(1, 1, 3, 15).setVerticalAlignment('middle');
  });
}

function resetActivityQARealReportSheet_(sheet) {
  runSpreadsheetOpWithRetry_('resetActivityQARealReportSheet', function() {
    sheet.getRange(1, 1, 2, 8).breakApart();
    clearActivityQASheet_(sheet, 8);

    sheet.setFrozenRows(4);
    sheet.setRowHeight(1, 26);
    sheet.setRowHeight(2, 24);
    sheet.setRowHeight(4, 24);

    var reportWidths = { 1: 70, 2: 80, 3: 170, 4: 80, 5: 360, 6: 120, 7: 120, 8: 260 };
    for (var col in reportWidths) {
      sheet.setColumnWidth(Number(col), reportWidths[col]);
    }

    sheet.getRange(1, 1, 1, 8).merge();
    sheet.getRange(2, 1, 1, 8).merge();

    sheet.getRange(1, 1).setValue('ACTIVITY REAL SAMPLE REPORT')
      .setBackground(ACTIVITY_QA_CONFIG.TITLE_BG)
      .setFontColor(ACTIVITY_QA_CONFIG.TITLE_TXT)
      .setFontWeight('bold')
      .setFontSize(12)
      .setHorizontalAlignment('center');

    sheet.getRange(2, 1).setValue('Built ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm'))
      .setBackground(ACTIVITY_QA_CONFIG.HEADER_BG)
      .setFontColor(ACTIVITY_QA_CONFIG.SUBTITLE_TXT)
      .setHorizontalAlignment('center');

    sheet.getRange(4, 1, 1, 8).setValues([['Scope', 'Flow', 'Sample', 'Score', 'Findings', 'Ref', 'Row', 'Notes']]);
    sheet.getRange(4, 1, 1, 8)
      .setBackground(ACTIVITY_QA_CONFIG.REPORT_HDR_BG)
      .setFontColor(ACTIVITY_QA_CONFIG.REPORT_HDR_TXT)
      .setFontWeight('bold')
      .setHorizontalAlignment('center');
  });
}

function padActivityQANumber_(num) {
  if (num < 10) return '00' + num;
  if (num < 100) return '0' + num;
  return String(num);
}

function buildActivityFixtureLink_(type, id, refText) {
  var url = '';
  if (typeof getKatanaWebUrl === 'function') {
    url = getKatanaWebUrl(type, id);
  }
  return { text: refText, url: url || CONFIG.KATANA_WEB_URL };
}

function buildActivityConformanceFixtures_() {
  return [
    {
      flow: 'F1',
      scenario: 'single-line receive',
      reference: 'PO-608',
      summary: '1 line received',
      context: '-> RECEIVING-DOCK @ MMH Kelowna',
      status: 'success',
      linkInfo: buildActivityFixtureLink_('po', 608, 'PO-608'),
      subItems: [
        { sku: 'UFC-1OZ', qty: 600, uom: 'each', action: 'receive into RECEIVING-DOCK', status: 'Added', qtyColor: 'green' }
      ]
    },
    {
      flow: 'F1',
      scenario: 'split-batch receive',
      reference: 'PO-609',
      summary: '1 line received',
      context: '-> RECEIVING-DOCK @ MMH Kelowna',
      status: 'success',
      linkInfo: buildActivityFixtureLink_('po', 609, 'PO-609'),
      subItems: [
        { sku: 'UFC-1OZ', qty: 600, uom: 'each', action: 'receive into RECEIVING-DOCK', isParent: true, batchCount: 2, qtyColor: 'green' },
        { nested: true, qty: 497, uom: 'each', action: 'lot:UFCA-120  exp:2029-03-31', status: 'Added', qtyColor: 'green' },
        { nested: true, qty: 103, uom: 'each', action: 'lot:UFCA-121  exp:2029-04-15', status: 'Added', qtyColor: 'green' }
      ]
    },
    {
      flow: 'F1',
      scenario: 'revert receive',
      reference: 'PO-610',
      summary: '1 line reverted',
      context: '<- RECEIVING-DOCK @ MMH Kelowna',
      status: 'reverted',
      linkInfo: buildActivityFixtureLink_('po', 610, 'PO-610'),
      subItems: [
        { sku: 'AGJAR-2', qty: 1, uom: 'each', action: 'remove from RECEIVING-DOCK', status: 'Removed', qtyColor: 'red' }
      ]
    },
    {
      flow: 'F1',
      scenario: 're-receive after revert',
      reference: 'PO-610',
      summary: '1 line received',
      context: '-> RECEIVING-DOCK @ MMH Kelowna',
      status: 'success',
      linkInfo: buildActivityFixtureLink_('po', 610, 'PO-610'),
      subItems: [
        { sku: 'AGJAR-2', qty: 1, uom: 'each', action: 'receive into RECEIVING-DOCK', status: 'Added', qtyColor: 'green' }
      ]
    },
    {
      flow: 'F1',
      scenario: 'partial receive',
      reference: 'PO-611',
      summary: '2 lines received',
      context: '-> RECEIVING-DOCK @ MMH Kelowna',
      status: 'partial',
      linkInfo: buildActivityFixtureLink_('po', 611, 'PO-611'),
      subItems: [
        { sku: 'LUP-2', qty: 240, uom: 'each', action: 'receive into RECEIVING-DOCK', status: 'Added', qtyColor: 'green' },
        { sku: 'TTFC-1OZ', qty: 120, uom: 'each', action: 'lot:TTFC611A  exp:2029-03-31  receive into RECEIVING-DOCK', status: 'Skipped', qtyColor: 'grey' }
      ]
    },
    {
      flow: 'F1',
      scenario: 'final receive after partial',
      reference: 'PO-612',
      summary: '1 line received',
      context: '-> RECEIVING-DOCK @ MMH Kelowna',
      status: 'success',
      linkInfo: buildActivityFixtureLink_('po', 612, 'PO-612'),
      subItems: [
        { sku: 'TTFC-1OZ', qty: 120, uom: 'each', action: 'lot:TTFC611A  exp:2029-03-31  receive into RECEIVING-DOCK', status: 'Added', qtyColor: 'green' }
      ]
    },
    {
      flow: 'F1',
      scenario: 'purchase uom conversion',
      reference: 'PO-613',
      summary: '1 line received',
      context: '-> RECEIVING-DOCK @ MMH Kelowna',
      status: 'success',
      linkInfo: buildActivityFixtureLink_('po', 613, 'PO-613'),
      subItems: [
        { sku: 'IP-EGG', qty: 144, uom: 'each', action: 'Katana 12 dozen  receive into RECEIVING-DOCK', status: 'Added', qtyColor: 'green' }
      ]
    },
    {
      flow: 'F1',
      scenario: 'exact lot not in wasp',
      reference: 'PO-614',
      summary: '1 line received',
      context: '-> RECEIVING-DOCK @ MMH Kelowna',
      status: 'partial',
      linkInfo: buildActivityFixtureLink_('po', 614, 'PO-614'),
      subItems: [
        { sku: 'EO-TT', qty: 20, uom: 'kg', action: 'lot:5006286976  exp:2028-03-31  receive into RECEIVING-DOCK', status: 'Skipped', qtyColor: 'grey' }
      ]
    },
    {
      flow: 'F2',
      scenario: 'synced add',
      reference: 'SA-22080',
      summary: 'UFC-1OZ x1050',
      context: 'Wasp | adjust @ MMH Kelowna',
      status: 'success',
      linkInfo: buildActivityFixtureLink_('sa', 22080, 'SA-22080'),
      subItems: [
        { sku: 'UFC-1OZ', qty: 1050, uom: 'each', action: 'bin:PROD-RECEIVING', status: 'Synced', qtyColor: 'green' }
      ]
    },
    {
      flow: 'F2',
      scenario: 'skipped lot not in Katana',
      reference: 'SA-22081',
      summary: 'UFC-1OZ x1050',
      context: 'Wasp | adjust @ MMH Kelowna',
      status: 'skipped',
      linkInfo: buildActivityFixtureLink_('sa', 22081, 'SA-22081'),
      subItems: [
        { sku: 'UFC-1OZ', qty: 1050, uom: 'each', action: 'bin:PROD-RECEIVING  lot not in Katana', status: 'Skipped', qtyColor: 'grey' }
      ]
    },
    {
      flow: 'F2',
      scenario: 'revert adjustment',
      reference: 'SA-22082',
      summary: '1 adjustment reversed',
      context: 'Katana | adjust revert @ MMH Kelowna',
      status: 'reverted',
      linkInfo: buildActivityFixtureLink_('sa', 22082, 'SA-22082'),
      subItems: [
        { sku: 'UFC-1OZ', qty: 1050, uom: 'each', action: 'bin:PROD-RECEIVING', status: 'Restored', qtyColor: 'green' }
      ]
    },
    {
      flow: 'F2',
      scenario: 'split-batch add',
      reference: 'SA-22083',
      summary: 'UFC-1OZ x600',
      context: 'Wasp | adjust @ MMH Kelowna',
      status: 'partial',
      linkInfo: buildActivityFixtureLink_('sa', 22083, 'SA-22083'),
      subItems: [
        { sku: 'UFC-1OZ', qty: 600, uom: 'each', action: 'bin:PROD-RECEIVING', isParent: true, batchCount: 2, qtyColor: 'grey' },
        { nested: true, qty: 497, uom: 'each', action: 'lot:ADJ-LOT-A  exp:2029-03-31', status: 'Synced', qtyColor: 'green' },
        { nested: true, qty: 103, uom: 'each', action: 'lot:ADJ-LOT-B  exp:2029-04-15', status: 'Skipped', qtyColor: 'grey' }
      ]
    },
    {
      flow: 'F3',
      scenario: 'synced transfer',
      reference: 'ST-253',
      summary: '3 lines moved',
      context: 'Katana poll | transfer Storage Warehouse -> MMH Kelowna',
      status: 'success',
      linkInfo: buildActivityFixtureLink_('st', 253, 'ST-253'),
      subItems: [
        { sku: 'AGJAR-2', qty: 7600, uom: 'each', action: '', status: 'Synced', qtyColor: 'green' },
        { sku: 'LABSSEAL-2', qty: 1360, uom: 'each', action: '', status: 'Synced', qtyColor: 'green' },
        { sku: 'B-YELLOW-2', qty: 550, uom: 'each', action: '', status: 'Synced', qtyColor: 'green' }
      ]
    },
    {
      flow: 'F3',
      scenario: 'mixed-source partial',
      reference: 'ST-258',
      summary: '4 lines moved',
      context: 'Katana poll | transfer MMH Kelowna -> MMH Mayfair',
      status: 'partial',
      linkInfo: buildActivityFixtureLink_('st', 258, 'ST-258'),
      subItems: [
        { sku: 'UFC-EOF-1OZ', qty: 1350, uom: 'each', action: 'move PROD-RECEIVING -> QA-Hold-1', status: 'Synced', qtyColor: 'green' },
        { sku: 'UFC-1OZ', qty: 2850, uom: 'each', action: 'move PROD-RECEIVING -> QA-Hold-1', status: 'Synced', qtyColor: 'green' },
        { sku: 'AGJAR-2', qty: 5760, uom: 'each', action: 'move RECEIVING-DOCK -> QA-Hold-1', status: 'Synced', qtyColor: 'green' },
        { sku: 'UFC-EOF-2OZ', qty: 228, uom: 'each', action: 'move RECEIVING-DOCK -> QA-Hold-1', status: 'Failed', error: 'No source stock found', qtyColor: 'red' }
      ]
    },
    {
      flow: 'F3',
      scenario: 'split-batch transfer',
      reference: 'ST-259',
      summary: '1 line moved',
      context: 'Katana poll | transfer MMH Kelowna -> MMH Mayfair',
      status: 'partial',
      linkInfo: buildActivityFixtureLink_('st', 259, 'ST-259'),
      subItems: [
        { sku: 'UFC-1OZ', qty: 600, uom: 'each', action: '', isParent: true, batchCount: 2, qtyColor: 'grey' },
        { nested: true, qty: 497, uom: 'each', action: 'lot:MOVE-LOT-A  exp:2029-03-31', status: 'Synced', qtyColor: 'green' },
        { nested: true, qty: 103, uom: 'each', action: 'lot:MOVE-LOT-B  exp:2029-04-15', status: 'Failed', error: 'Move failed', qtyColor: 'red' }
      ]
    },
    {
      flow: 'F3',
      scenario: 'revert transfer',
      reference: 'ST-260',
      summary: '1 line reversed',
      context: 'Katana poll | transfer revert MMH Mayfair -> MMH Kelowna',
      status: 'reverted',
      linkInfo: buildActivityFixtureLink_('st', 260, 'ST-260'),
      subItems: [
        { sku: 'UFC-1OZ', qty: 600, uom: 'each', action: '', status: 'Synced', qtyColor: 'green' }
      ]
    },
    {
      flow: 'F4',
      scenario: 'split-batch partial',
      reference: 'MO-7379',
      summary: 'UFC-1OZ x600',
      context: '-> PROD-RECEIVING',
      status: 'partial',
      linkInfo: buildActivityFixtureLink_('mo', 7379, 'MO-7379'),
      subItems: [
        { sku: 'LUP-1', qty: 600, uom: 'each', action: 'consume from PRODUCTION', isParent: true, batchCount: 2, qtyColor: 'red' },
        { nested: true, qty: 497, uom: 'each', action: 'lot:UFC416A  exp:2029-03-31', status: 'Skipped', qtyColor: 'red' },
        { nested: true, qty: 103, uom: 'each', action: 'lot:UFC416B  exp:2029-03-31', status: 'Consumed', qtyColor: 'red' },
        { sku: 'UFC-1OZ', qty: 600, uom: 'each', action: 'produce into PROD-RECEIVING', status: 'Produced', qtyColor: 'green' }
      ]
    },
    {
      flow: 'F4',
      scenario: 'failed manufacture',
      reference: 'MO-7391',
      summary: 'UFC-4OZ x240',
      context: '-> PROD-RECEIVING',
      status: 'failed',
      linkInfo: buildActivityFixtureLink_('mo', 7391, 'MO-7391'),
      headerError: 'Manufacturing failed',
      subItems: [
        { sku: 'B-WAX', qty: 7200, uom: 'grams', action: 'consume from PRODUCTION', status: 'Failed', error: 'Exact Katana lot unresolved', qtyColor: 'red' }
      ]
    },
    {
      flow: 'F4',
      scenario: 'revert',
      reference: 'MO-7359',
      summary: '1 ingredient restored',
      context: '<- PRODUCTION',
      status: 'reverted',
      linkInfo: buildActivityFixtureLink_('mo', 7359, 'MO-7359'),
      subItems: [
        { sku: 'PL-PURPLE-1', qty: 1724, uom: 'each', action: 'restore to PRODUCTION', status: 'Restored', qtyColor: 'green' },
        { sku: 'LUP-1', qty: 1714, uom: 'each', action: 'remove from PROD-RECEIVING', status: 'Removed', qtyColor: 'red' }
      ]
    },
    {
      flow: 'F4',
      scenario: 'clean manufacture',
      reference: 'MO-7360',
      summary: 'UFC-2OZ x456',
      context: '-> PROD-RECEIVING',
      status: 'success',
      linkInfo: buildActivityFixtureLink_('mo', 7360, 'MO-7360'),
      subItems: [
        { sku: 'FG-US-2', qty: 456, uom: 'each', action: 'lot:UEOF051C  exp:2029-03-31  consume from PRODUCTION', status: 'Consumed', qtyColor: 'red' },
        { sku: 'PL-YELLOW-2', qty: 521, uom: 'each', action: 'consume from PRODUCTION', status: 'Consumed', qtyColor: 'red' },
        { sku: 'LUY-2', qty: 456, uom: 'each', action: 'lot:UEOF051C  exp:2029-03-31  produce into PROD-RECEIVING', status: 'Produced', qtyColor: 'green' }
      ]
    },
    {
      flow: 'F4',
      scenario: 'partial manufacture exact lot not in wasp',
      reference: 'MO-7361',
      summary: 'UFC-4OZ x240',
      context: '-> PROD-RECEIVING',
      status: 'partial',
      linkInfo: buildActivityFixtureLink_('mo', 7361, 'MO-7361'),
      subItems: [
        { sku: 'EO-LAV', qty: 300, uom: 'grams', action: 'lot:5005828385  exp:2029-07-01  consume from PRODUCTION', status: 'Consumed', qtyColor: 'red' },
        { sku: 'B-WAX', qty: 7200, uom: 'grams', action: 'lot:MHF121525  exp:2026-12-15  consume from PRODUCTION', status: 'Skipped', qtyColor: 'grey' },
        { sku: 'UFC-4OZ', qty: 240, uom: 'each', action: 'lot:UFC418B  exp:2029-03-31  produce into PROD-RECEIVING', status: 'Produced', qtyColor: 'green' }
      ]
    },
    {
      flow: 'F4',
      scenario: 'output lot from order fallback',
      reference: 'MO-7362',
      summary: 'UFC-1OZ x1714',
      context: '-> PROD-RECEIVING',
      status: 'success',
      linkInfo: buildActivityFixtureLink_('mo', 7362, 'MO-7362'),
      subItems: [
        { sku: 'LABSSEAL-1', qty: 1714, uom: 'each', action: 'consume from PRODUCTION', status: 'Consumed', qtyColor: 'red' },
        { sku: 'IPL-J-1', qty: 1714, uom: 'each', action: 'consume from PRODUCTION', status: 'Consumed', qtyColor: 'red' },
        { sku: 'FGJ-MS-1', qty: 1714, uom: 'each', action: 'lot:UFC418B  exp:2029-03-09  produce into PROD-RECEIVING', status: 'Produced', qtyColor: 'green' }
      ]
    },
    {
      flow: 'F4',
      scenario: 'repeat build partial after revert',
      reference: 'MO-7432',
      summary: 'CAR-1OZ x1',
      context: '-> PROD-RECEIVING',
      status: 'partial',
      linkInfo: buildActivityFixtureLink_('mo', 7432, 'MO-7432'),
      subItems: [
        { sku: 'NI-B2221210', qty: 0.00666, uom: 'box', action: 'bin:PRODUCTION', status: 'Consumed', qtyColor: 'red' },
        { sku: 'FBA-FRAGILE', qty: 0.01333, uom: 'each', action: 'bin:PRODUCTION', status: 'Consumed', qtyColor: 'red' },
        { sku: 'B-PINK-1', qty: 1, uom: 'each', action: 'bin:PRODUCTION', status: 'Consumed', qtyColor: 'red' },
        { sku: 'LCP-1', qty: 1, uom: 'each', action: 'lot:CAR024  exp:2029-02-28  bin:PRODUCTION', status: 'Skipped', qtyColor: 'grey' },
        { sku: 'L-P', qty: 1, uom: 'each', action: 'bin:PRODUCTION', status: 'Consumed', qtyColor: 'red' },
        { sku: 'CAR-1OZ', qty: 1, uom: 'each', action: 'bin:PROD-RECEIVING', status: 'Produced', qtyColor: 'green' }
      ]
    },
    {
      flow: 'F4',
      scenario: 'repeat build revert after edit state',
      reference: 'MO-7432',
      summary: '4 ingredients restored',
      context: '<- PRODUCTION',
      status: 'reverted',
      linkInfo: buildActivityFixtureLink_('mo', 7432, 'MO-7432'),
      subItems: [
        { sku: 'NI-B2221210', qty: 0.00666, uom: 'box', action: 'bin:PRODUCTION', status: 'Restored', qtyColor: 'green' },
        { sku: 'FBA-FRAGILE', qty: 0.01333, uom: 'each', action: 'bin:PRODUCTION', status: 'Restored', qtyColor: 'green' },
        { sku: 'B-PINK-1', qty: 1, uom: 'each', action: 'bin:PRODUCTION', status: 'Restored', qtyColor: 'green' },
        { sku: 'L-P', qty: 1, uom: 'each', action: 'bin:PRODUCTION', status: 'Restored', qtyColor: 'green' },
        { sku: 'CAR-1OZ', qty: 1, uom: 'each', action: 'bin:PROD-RECEIVING', status: 'Removed', qtyColor: 'red' }
      ]
    },
    {
      flow: 'F4',
      scenario: 'edited rebuild after repeat revert',
      reference: 'MO-7432',
      summary: 'CAR-1OZ x7',
      context: '-> PROD-RECEIVING',
      status: 'partial',
      linkInfo: buildActivityFixtureLink_('mo', 7432, 'MO-7432'),
      subItems: [
        { sku: 'NI-PUMP-0', qty: 3, uom: 'each', action: 'bin:PRODUCTION', status: 'Consumed', qtyColor: 'red' },
        { sku: 'NI-B2221210', qty: 0.00666, uom: 'box', action: 'bin:PRODUCTION', status: 'Consumed', qtyColor: 'red' },
        { sku: 'FBA-FRAGILE', qty: 0.01333, uom: 'each', action: 'bin:PRODUCTION', status: 'Consumed', qtyColor: 'red' },
        { sku: 'B-PINK-1', qty: 1, uom: 'each', action: 'bin:PRODUCTION', status: 'Consumed', qtyColor: 'red' },
        { sku: 'LCP-1', qty: 1, uom: 'each', action: 'lot:CAR024  exp:2029-02-28  bin:PRODUCTION', status: 'Skipped', qtyColor: 'grey' },
        { sku: 'L-P', qty: 1, uom: 'each', action: 'bin:PRODUCTION', status: 'Consumed', qtyColor: 'red' },
        { sku: 'CAR-1OZ', qty: 7, uom: 'each', action: 'bin:PROD-RECEIVING', status: 'Produced', qtyColor: 'green' }
      ]
    },
    {
      flow: 'F5',
      scenario: 'full shipment',
      reference: '#94042',
      summary: '3 lines shipped',
      context: '-> SHOPIFY @ MMH Kelowna',
      status: 'success',
      linkInfo: buildActivityFixtureLink_('so', 94042, '#94042'),
      subItems: [
        { sku: 'UFC-4OZ', qty: 3, uom: 'each', action: 'deduct for shipment', status: 'Deducted', qtyColor: 'red' },
        { sku: 'WH-8OZ-W', qty: 1, uom: 'each', action: 'deduct for shipment  lot:AS081525-1  exp:2027-08-30', status: 'Deducted', qtyColor: 'red' },
        { sku: 'TTFC-1OZ', qty: 1, uom: 'each', action: 'deduct for shipment', status: 'Deducted', qtyColor: 'red' }
      ]
    },
    {
      flow: 'F5',
      scenario: 'split-batch partial shipment',
      reference: '#94188',
      summary: '1 line shipped',
      context: '-> SHOPIFY @ MMH Kelowna',
      status: 'partial',
      linkInfo: buildActivityFixtureLink_('so', 94188, '#94188'),
      subItems: [
        { sku: 'UFC-1OZ', qty: 600, uom: 'each', action: 'deduct for shipment', isParent: true, batchCount: 2, qtyColor: 'red' },
        { nested: true, qty: 497, uom: 'each', action: 'lot:SHIP-LOT-A  exp:2029-03-31', status: 'Deducted', qtyColor: 'red' },
        { nested: true, qty: 103, uom: 'each', action: 'lot:SHIP-LOT-B  exp:2029-04-15', status: 'Failed', error: 'Remove failed', qtyColor: 'red' }
      ]
    },
    {
      flow: 'F5',
      scenario: 'shipment reversal',
      reference: '#94189',
      summary: '1 line returned',
      context: '<- SHOPIFY @ MMH Kelowna',
      status: 'reverted',
      linkInfo: buildActivityFixtureLink_('so', 94189, '#94189'),
      subItems: [
        { sku: 'UFC-1OZ', qty: 600, uom: 'each', action: 'return to stock', status: 'Returned', qtyColor: 'green' }
      ]
    },
    {
      flow: 'F6',
      scenario: 'full staging',
      reference: 'SO-AMZ-22081',
      summary: '1 line staged',
      context: '-> AMAZON-FBA',
      status: 'success',
      linkInfo: buildActivityFixtureLink_('so', 22081, 'SO-AMZ-22081'),
      subItems: [
        { sku: 'VB-UFC-KIT', qty: 24, uom: 'each', action: 'stage into AMAZON-FBA', status: 'Complete', qtyColor: 'green' }
      ]
    },
    {
      flow: 'F6',
      scenario: 'split-batch partial staging',
      reference: 'SO-AMZ-22082',
      summary: '1 line staged',
      context: '-> AMAZON-FBA',
      status: 'partial',
      linkInfo: buildActivityFixtureLink_('so', 22082, 'SO-AMZ-22082'),
      subItems: [
        { sku: 'VB-UFC-KIT', qty: 600, uom: 'each', action: 'stage into AMAZON-FBA', isParent: true, batchCount: 2, qtyColor: 'green' },
        { nested: true, qty: 497, uom: 'each', action: 'lot:FBA-A1  exp:2029-03-31', status: 'Complete', qtyColor: 'green' },
        { nested: true, qty: 103, uom: 'each', action: 'lot:FBA-A2  exp:2029-04-15', status: 'Skipped', qtyColor: 'grey' }
      ]
    },
    {
      flow: 'F6',
      scenario: 'staging reversal',
      reference: 'SO-AMZ-22083',
      summary: '1 line reversed',
      context: '<- SHIPPING-DOCK',
      status: 'reverted',
      linkInfo: buildActivityFixtureLink_('so', 22083, 'SO-AMZ-22083'),
      subItems: [
        { sku: 'VB-UFC-KIT', qty: 24, uom: 'each', action: 'restore to SHIPPING-DOCK', status: 'Restored', qtyColor: 'green' }
      ]
    }
  ];
}

function validateActivityFixture_(fixture) {
  var findings = [];
  if (!fixture.flow || !FLOW_LABELS[fixture.flow]) findings.push('Unknown flow key');
  if (!fixture.reference) findings.push('Missing header reference');
  if (!fixture.summary) findings.push('Missing header summary');
  if (!fixture.context) findings.push('Missing header context');
  if (!fixture.linkInfo || fixture.linkInfo.text !== fixture.reference || !fixture.linkInfo.url) {
    findings.push('Header link must target only the reference token');
  }
  if (!fixture.subItems || fixture.subItems.length === 0) findings.push('Entry has no sub-items');

  for (var i = 0; fixture.subItems && i < fixture.subItems.length; i++) {
    var item = fixture.subItems[i];
    if (item.nested && item.sku) findings.push('Nested batch row should not repeat SKU');
    if (item.isParent && (!item.batchCount || item.batchCount < 2)) findings.push('Split-batch parent missing batch count');
    if (item.status === 'Skipped' && item.error) findings.push('Skipped row should keep Error blank');
    if (item.error && !item.status) findings.push('Error row missing status');
  }

  return findings;
}

function renderActivityFixtureToSheet_(sheet, fixture, execId) {
  var details = fixture.reference + '  ' + fixture.summary;
  return appendActivityBlockToSheet_(
    sheet,
    fixture.flow,
    details,
    fixture.status,
    fixture.context,
    fixture.subItems,
    fixture.linkInfo,
    execId,
    fixture.headerError || ''
  );
}

function buildRealActivityFixtureFromBlock_(sheet, block, index) {
  var header = block.header || [];
  var headerText = String(header[3] || '').trim();
  var statusText = String(header[4] || '').trim();
  var headerError = String(header[5] || '').trim();
  var statusKey = mapActivityHeaderStatusToKey_(statusText);
  var sourceLink = getActivityHeaderLinkData_(sheet, block.headerRow);
  var reference = resolveRealActivityReference_(block.flow, headerText, sourceLink);
  var linkInfo = buildRealActivityLinkInfo_(block.flow, reference, sourceLink);
  var subItems = buildRealActivitySubItems_(block.flow, block, statusKey);
  var summary = buildRealActivitySummary_(block.flow, statusKey, subItems, headerText);
  var context = buildRealActivityContext_(block.flow, statusKey, subItems, headerText);

  return {
    flow: block.flow,
    scenario: 'real sample ' + (index + 1),
    reference: reference,
    summary: summary,
    context: context,
    status: statusKey,
    headerError: normalizeRealActivityHeaderError_(block.flow, statusKey, headerError),
    linkInfo: linkInfo,
    subItems: subItems
  };
}

function mapActivityHeaderStatusToKey_(statusText) {
  var text = String(statusText || '').trim().toLowerCase();
  if (text === 'failed') return 'failed';
  if (text === 'partial') return 'partial';
  if (text === 'skipped') return 'skipped';
  if (text === 'reverted') return 'reverted';
  return 'success';
}

function normalizeRealActivityHeaderError_(flow, statusKey, headerError) {
  var text = String(headerError || '').trim();
  if (/^\d+\s+(failed|skipped|warning|warnings)(,\s*\d+\s+\w+)*$/i.test(text)) text = '';
  if (text) return text;
  if (statusKey !== 'failed') return '';
  if (flow === 'F1') return 'Receive failed';
  if (flow === 'F2') return 'Adjustment failed';
  if (flow === 'F3') return 'Transfer failed';
  if (flow === 'F4') return 'Manufacturing failed';
  if (flow === 'F5') return 'Shipment failed';
  if (flow === 'F6') return 'FBA staging failed';
  return '';
}

function getActivityHeaderLinkData_(sheet, rowNum) {
  try {
    var richText = sheet.getRange(rowNum, 4).getRichTextValue();
    if (!richText || typeof richText.getRuns !== 'function') return null;
    var runs = richText.getRuns();
    for (var i = 0; i < runs.length; i++) {
      var url = runs[i].getLinkUrl && runs[i].getLinkUrl();
      if (url) {
        return {
          text: String(runs[i].getText() || '').trim(),
          url: url
        };
      }
    }
  } catch (e) {}
  return null;
}

function resolveRealActivityReference_(flow, headerText, sourceLink) {
  var token = getActivityReferenceToken_([null, null, null, headerText]);
  var extracted = extractCanonicalActivityRef_(token, '', '');

  if (flow === 'F2' && sourceLink && sourceLink.url) {
    var saMatch = String(sourceLink.url).match(/stockadjustment\/([^\/?#]+)/i);
    if (saMatch) return 'SA-' + saMatch[1];
  }

  if (flow === 'F2' && (!extracted || /^WASP$/i.test(extracted))) {
    return 'WASP Adjustment';
  }
  return extracted || token || (FLOW_LABELS[flow] || flow);
}

function buildRealActivityLinkInfo_(flow, reference, sourceLink) {
  if (sourceLink && sourceLink.url) {
    return { text: reference, url: sourceLink.url };
  }

  var refText = String(reference || '');
  var numMatch = refText.match(/(\d+)/);
  if (!numMatch) return null;

  if (flow === 'F1') return { text: refText, url: getKatanaWebUrl('po', numMatch[1]) };
  if (flow === 'F2') return { text: refText, url: getKatanaWebUrl('sa', numMatch[1]) };
  if (flow === 'F3') return { text: refText, url: getKatanaWebUrl('st', numMatch[1]) };
  if (flow === 'F4') return { text: refText, url: getKatanaWebUrl('mo', numMatch[1]) };
  if (flow === 'F6') return { text: refText, url: getKatanaWebUrl('so', numMatch[1]) };
  return null;
}

function buildRealActivitySubItems_(flow, block, statusKey) {
  var items = [];
  var headerText = String(block.header && block.header[3] || '').trim();
  var headerSite = extractPreviewSite_(headerText);
  var f3ShowRoute = /mixed-source|mixed-dest/i.test(headerText);

  for (var i = 0; i < block.subRows.length; i++) {
    var row = block.subRows[i];
    var parsed = parsePreviewDetailRow_(row.values[3]);
    var rowStatus = String(row.values[4] || '').trim();
    var rowError = String(row.values[5] || '').trim();
    var locations = extractPreviewLocations_(parsed.tail);
    var lot = extractPreviewLot_(parsed.tail);
    var expiry = extractPreviewExp_(parsed.tail);
    var batchId = extractPreviewBatchId_(parsed.tail);
    var item = {
      sku: parsed.sku,
      qty: parsed.qty,
      uom: parsed.uom,
      nested: parsed.nested,
      status: rowStatus,
      error: rowStatus === 'Skipped' ? '' : rowError,
      success: rowStatus !== 'Failed'
    };

    if (parsed.batchCount) {
      item.isParent = true;
      item.batchCount = parsed.batchCount;
      item.status = '';
      item.error = '';
      item.qtyColor = 'grey';
    }

    if (flow === 'F1') {
      var f1Loc = getActivityDisplayLocation_(locations[0] || 'RECEIVING-DOCK');
      if (item.nested) {
        item.action = buildActivityBatchMeta_(lot, expiry);
        item.qtyColor = (statusKey === 'reverted' || rowStatus === 'Removed' || rowStatus === 'Failed') ? 'red' : 'green';
      } else if (item.isParent) {
        item.action = f1Loc;
      } else {
        item.action = buildActivityActionText_(
          (statusKey === 'reverted' || rowStatus === 'Removed') ? ('remove from ' + f1Loc) : ('receive into ' + f1Loc),
          lot,
          expiry
        );
        item.qtyColor = (statusKey === 'reverted' || rowStatus === 'Removed' || rowStatus === 'Failed') ? 'red' : 'green';
      }
    } else if (flow === 'F2') {
      var f2Loc = getActivityDisplayLocation_(locations[0] || '');
      var f2Extras = [];
      if (/lot not in katana/i.test(parsed.tail)) f2Extras.push('lot not in Katana');
      else if (/not in katana/i.test(parsed.tail)) f2Extras.push('not in Katana');
      item.action = buildActivityCompactMeta_(headerSite, f2Loc, lot, expiry, f2Extras);
      item.qtyColor = (rowStatus === 'Removed' || rowStatus === 'Failed') ? 'red' : 'green';
    } else if (flow === 'F3') {
      var f3From = getActivityDisplayLocation_(locations[0] || 'mixed-source');
      var f3To = getActivityDisplayLocation_(locations[1] || '');
      if (item.nested) {
        item.action = buildActivityBatchMeta_(lot, expiry);
        item.qtyColor = rowStatus === 'Failed' ? 'red' : 'green';
      } else if (item.isParent) {
        item.action = f3ShowRoute ? (f3From + (f3To ? ' -> ' + f3To : '')) : '';
      } else {
        item.action = f3ShowRoute
          ? buildActivityActionText_('move ' + f3From + (f3To ? ' -> ' + f3To : ''), lot, expiry)
          : buildActivityCompactMeta_('', '', lot, expiry);
        item.qtyColor = rowStatus === 'Failed' ? 'red' : 'green';
      }
    } else if (flow === 'F4') {
      var f4Loc = getActivityDisplayLocation_(locations[0] || ((rowStatus === 'Produced' || rowStatus === 'Removed') ? 'PROD-RECEIVING' : 'PRODUCTION'));
      var f4BaseAction = 'consume from ' + f4Loc;
      if (rowStatus === 'Produced') f4BaseAction = 'produce into ' + f4Loc;
      else if (rowStatus === 'Restored') f4BaseAction = 'restore to ' + f4Loc;
      else if (rowStatus === 'Removed') f4BaseAction = 'remove from ' + f4Loc;

      if (item.nested) {
        item.action = buildActivityBatchMeta_(lot, expiry);
      } else if (item.isParent) {
        item.action = f4Loc;
      } else {
        item.action = joinActivitySegments_([
          f4BaseAction,
          batchId && !lot ? 'batch_id:' + batchId : '',
          buildActivityBatchMeta_(lot, expiry)
        ]);
      }
      item.qtyColor = (rowStatus === 'Produced' || rowStatus === 'Restored') ? 'green' : (item.isParent ? 'grey' : 'red');
    } else if (flow === 'F5') {
      if (item.nested) {
        item.action = buildActivityBatchMeta_(lot, expiry);
        item.qtyColor = rowStatus === 'Returned' ? 'green' : 'red';
      } else if (item.isParent) {
        item.action = 'SHOPIFY';
      } else {
        item.action = buildActivityActionText_(rowStatus === 'Returned' ? 'return to stock' : 'deduct for shipment', lot, expiry);
        item.qtyColor = rowStatus === 'Returned' ? 'green' : 'red';
      }
    } else if (flow === 'F6') {
      if (item.nested) {
        item.action = buildActivityBatchMeta_(lot, expiry);
      } else if (item.isParent) {
        item.action = 'AMAZON-FBA';
      } else {
        item.action = buildActivityActionText_(
          statusKey === 'reverted' || rowStatus === 'Reverted'
            ? 'restore to SHIPPING-DOCK'
            : 'stage into AMAZON-FBA',
          lot,
          expiry
        );
      }
      item.qtyColor = item.isParent ? 'grey' : 'green';
    }

    items.push(item);
  }

  return items;
}

function buildRealActivitySummary_(flow, statusKey, subItems, headerText) {
  var topLevelCount = countTopLevelPreviewItems_(subItems);
  if (flow === 'F1') return buildActivityCountSummary_(topLevelCount, 'line', 'lines', statusKey === 'reverted' ? 'reverted' : 'received');
  if (flow === 'F2') {
    var firstF2 = findFirstTopLevelPreviewItem_(subItems);
    return topLevelCount === 1 && firstF2 ? (firstF2.sku + ' x' + firstF2.qty) : buildActivityCountSummary_(topLevelCount, 'adjustment', 'adjustments', '');
  }
  if (flow === 'F3') return buildActivityCountSummary_(topLevelCount, 'line', 'lines', /reversed/i.test(headerText) ? 'reversed' : 'moved');
  if (flow === 'F4') {
    if (statusKey === 'reverted' || hasPreviewStatus_(subItems, 'Restored')) {
      return buildActivityCountSummary_(countPreviewStatus_(subItems, 'Restored') || topLevelCount, 'ingredient', 'ingredients', 'restored');
    }
    var produced = findTopLevelPreviewItemByStatus_(subItems, 'Produced');
    return produced ? (produced.sku + ' x' + produced.qty) : buildActivityCountSummary_(topLevelCount, 'line', 'lines', 'manufactured');
  }
  if (flow === 'F5') return buildActivityCountSummary_(topLevelCount, 'line', 'lines', hasPreviewStatus_(subItems, 'Returned') ? 'returned' : 'shipped');
  if (flow === 'F6') return buildActivityCountSummary_(topLevelCount, 'line', 'lines', statusKey === 'reverted' ? 'reversed' : 'staged');
  return buildActivityCountSummary_(topLevelCount, 'line', 'lines', 'processed');
}

function buildRealActivityContext_(flow, statusKey, subItems, headerText) {
  var site = extractPreviewSite_(headerText);
  var sites = extractPreviewSites_(headerText);

  if (flow === 'F1') {
    var f1Loc = findPreviewLocationByStatus_(subItems, statusKey === 'reverted' ? 'Removed' : 'Added') || 'RECEIVING-DOCK';
    return (statusKey === 'reverted' ? '<- ' : '-> ') + f1Loc + (site ? ' @ ' + site : '');
  }

  if (flow === 'F2') {
    var f2Source = /sheets/i.test(headerText) ? 'Sheets' : (/activity/i.test(headerText) ? 'Activity' : (statusKey === 'reverted' ? 'Katana' : 'Wasp'));
    var f2Action = /retry/i.test(headerText) ? 'retry' : (statusKey === 'reverted' ? 'adjust revert' : 'adjust');
    return buildActivitySourceActionContext_(f2Source, f2Action, site);
  }

  if (flow === 'F3') {
    if (sites.length >= 2) {
      return buildActivityTransferContext_(
        'Katana poll',
        /reversed/i.test(headerText) ? 'transfer revert' : 'transfer',
        sites[0],
        sites[1]
      );
    }
    var pairs = collectPreviewMovePairs_(subItems);
    var fromParts = {};
    var toParts = {};
    for (var i = 0; i < pairs.length; i++) {
      fromParts[pairs[i].from || 'mixed-source'] = true;
      toParts[pairs[i].to || 'mixed-dest'] = true;
    }
    var fromKeys = Object.keys(fromParts);
    var toKeys = Object.keys(toParts);
    var fromLabel = fromKeys.length > 1 ? 'mixed-source' : (fromKeys[0] || 'mixed-source');
    var toLabel = toKeys.length > 1 ? 'mixed-dest' : (toKeys[0] || 'mixed-dest');
    return buildActivityTransferContext_('Katana poll', /reversed/i.test(headerText) ? 'transfer revert' : 'transfer', fromLabel, toLabel);
  }

  if (flow === 'F4') {
    if (statusKey === 'reverted' || hasPreviewStatus_(subItems, 'Restored')) return '<- PRODUCTION';
    var produced = findTopLevelPreviewItemByStatus_(subItems, 'Produced');
    var producedLoc = extractPreviewLocations_(produced && produced.action)[0] || extractPreviewLocations_(headerText)[0] || 'PRODUCTION';
    return '-> ' + getActivityDisplayLocation_(producedLoc);
  }

  if (flow === 'F5') {
    return (hasPreviewStatus_(subItems, 'Returned') ? '<- ' : '-> ') + 'SHOPIFY' + (site ? ' @ ' + site : '');
  }

  if (flow === 'F6') {
    return statusKey === 'reverted' ? '<- SHIPPING-DOCK' : '-> AMAZON-FBA';
  }

  return '';
}

function parsePreviewDetailRow_(detailText) {
  var text = String(detailText || '');
  var cleaned = text.replace(/^[\s\|\-`~\u2500-\u257F]+/, '').trim();
  var batchCountMatch = cleaned.match(/\((\d+)\s+batches\)\s*$/i);
  var batchCount = batchCountMatch ? Number(batchCountMatch[1]) : 0;
  if (batchCount) cleaned = cleaned.replace(/\((\d+)\s+batches\)\s*$/i, '').trim();

  var nested = /^x[^\s]+/i.test(cleaned);
  var match = nested
    ? cleaned.match(/^x([^\s]+)(?:\s+([A-Za-z]+))?(?:\s+(.*))?$/)
    : cleaned.match(/^(\S+)\s+x([^\s]+)(?:\s+([A-Za-z]+))?(?:\s+(.*))?$/);

  var qty = nested ? (match ? match[1] : '') : (match ? match[2] : '');
  qty = String(qty || '').replace(/^-/, '');

  return {
    text: cleaned,
    nested: nested,
    sku: nested ? '' : (match ? match[1] : ''),
    qty: qty,
    uom: normalizeUom(match ? (nested ? match[2] : match[3]) || '' : ''),
    tail: String(match ? (nested ? match[3] : match[4]) || '' : '').trim(),
    batchCount: batchCount
  };
}

function extractPreviewLocations_(text) {
  var matches = String(text || '').match(/QA-Hold-\d+|RECEIVING-DOCK|PROD-RECEIVING|PRODUCTION|SW-STORAGE|SHOPIFY|AMAZON-FBA-USA|AMAZON-FBA|SHIPPING-DOCK|BULK|UNSORTED/gi) || [];
  var deduped = [];
  for (var i = 0; i < matches.length; i++) {
    var loc = getActivityDisplayLocation_(matches[i]);
    if (deduped.indexOf(loc) < 0) deduped.push(loc);
  }
  return deduped;
}

function extractPreviewLot_(text) {
  var match = String(text || '').match(/lot:([^\s]+)/i);
  return match ? match[1] : '';
}

function extractPreviewExp_(text) {
  var match = String(text || '').match(/exp:([0-9-]+)/i);
  return match ? match[1] : '';
}

function extractPreviewBatchId_(text) {
  var match = String(text || '').match(/batch_id:([^\s]+)/i);
  return match ? match[1] : '';
}

function extractPreviewSite_(headerText) {
  var text = String(headerText || '').trim();
  var atMatch = text.match(/@\s*([^@]+)$/);
  if (atMatch) return atMatch[1].trim();
  var siteMatch = text.match(/(MMH\s+[A-Za-z]+(?:\s+[A-Za-z]+)*|Storage Warehouse|Amazon USA)/);
  return siteMatch ? siteMatch[1].trim() : '';
}

function extractPreviewSites_(headerText) {
  var matches = String(headerText || '').match(/MMH\s+[A-Za-z]+(?:\s+[A-Za-z]+)*|Storage Warehouse|Amazon USA/g) || [];
  var deduped = [];
  for (var i = 0; i < matches.length; i++) {
    var site = String(matches[i] || '').trim();
    if (site && deduped.indexOf(site) < 0) deduped.push(site);
  }
  return deduped;
}

function countTopLevelPreviewItems_(subItems) {
  var count = 0;
  for (var i = 0; i < subItems.length; i++) {
    if (!subItems[i].nested) count++;
  }
  return count;
}

function hasPreviewStatus_(subItems, status) {
  for (var i = 0; i < subItems.length; i++) {
    if (String(subItems[i].status || '').trim() === status) return true;
  }
  return false;
}

function countPreviewStatus_(subItems, status) {
  var count = 0;
  for (var i = 0; i < subItems.length; i++) {
    if (String(subItems[i].status || '').trim() === status && !subItems[i].nested) count++;
  }
  return count;
}

function findFirstTopLevelPreviewItem_(subItems) {
  for (var i = 0; i < subItems.length; i++) {
    if (!subItems[i].nested) return subItems[i];
  }
  return null;
}

function findTopLevelPreviewItemByStatus_(subItems, status) {
  for (var i = 0; i < subItems.length; i++) {
    if (!subItems[i].nested && String(subItems[i].status || '').trim() === status) return subItems[i];
  }
  return null;
}

function findPreviewLocationByStatus_(subItems, status) {
  for (var i = 0; i < subItems.length; i++) {
    var item = subItems[i];
    if (item.nested) continue;
    if (status && String(item.status || '').trim() !== status) continue;
    var locations = extractPreviewLocations_(item.action);
    if (locations.length) return locations[locations.length - 1];
  }
  return '';
}

function collectPreviewMovePairs_(subItems) {
  var pairs = [];
  for (var i = 0; i < subItems.length; i++) {
    var item = subItems[i];
    if (item.nested) continue;
    var locations = extractPreviewLocations_(item.action);
    pairs.push({
      from: locations[0] || '',
      to: locations[1] || ''
    });
  }
  return pairs;
}

function readActivityBlockAtRow_(sheet, headerRow, subItems) {
  var subCount = subItems ? subItems.length : 0;
  var values = sheet.getRange(headerRow, 1, subCount + 1, 6).getDisplayValues();
  var block = {
    flow: getFlowKeyByLabel_(String(values[0][2] || '').trim()),
    flowLabel: String(values[0][2] || '').trim(),
    headerRow: headerRow,
    header: values[0],
    subRows: []
  };

  for (var i = 1; i < values.length; i++) {
    block.subRows.push({
      rowNum: headerRow + i,
      values: values[i]
    });
  }

  return block;
}

function getFlowKeyByLabel_(label) {
  for (var flow in FLOW_LABELS) {
    if (FLOW_LABELS[flow] === label) return flow;
  }
  return '';
}

function readActivityBlocksWithinWindow_(sheet, rowWindow) {
  if (!sheet || sheet.getLastRow() < ACTIVITY_QA_CONFIG.DATA_START_ROW) return [];

  var lastRow = sheet.getLastRow();
  var totalDataRows = lastRow - 3;
  var rowsToRead = Math.min(totalDataRows, rowWindow || ACTIVITY_QA_CONFIG.LIVE_AUDIT_ROW_WINDOW);
  var startRow = Math.max(ACTIVITY_QA_CONFIG.DATA_START_ROW, lastRow - rowsToRead + 1);
  var values = sheet.getRange(startRow, 1, rowsToRead, 6).getDisplayValues();
  var allBlocks = [];
  var currentBlock = null;

  for (var i = 0; i < values.length; i++) {
    var rowNum = startRow + i;
    var row = values[i];
    var execId = String(row[0] || '').trim();
    var flowLabel = String(row[2] || '').trim();
    var hasData = String(row[3] || row[4] || row[5] || '').trim() !== '';

    if (execId && flowLabel) {
      if (currentBlock) allBlocks.push(currentBlock);
      currentBlock = {
        flow: getFlowKeyByLabel_(flowLabel),
        flowLabel: flowLabel,
        headerRow: rowNum,
        header: row,
        subRows: []
      };
      continue;
    }

    if (currentBlock && hasData) {
      currentBlock.subRows.push({
        rowNum: rowNum,
        values: row
      });
    }
  }

  if (currentBlock) allBlocks.push(currentBlock);
  return allBlocks;
}

function readRecentActivityBlocksByFlow_(limitPerFlow, rowWindow) {
  var ss = typeof getDebugSpreadsheet_ === 'function'
    ? getDebugSpreadsheet_()
    : SpreadsheetApp.openById(CONFIG.DEBUG_SHEET_ID);
  var sheet = ss.getSheetByName('Activity');
  var allBlocks = readActivityBlocksWithinWindow_(sheet, rowWindow || ACTIVITY_QA_CONFIG.LIVE_AUDIT_ROW_WINDOW);
  var selectedByFlow = {};

  for (var i = allBlocks.length - 1; i >= 0; i--) {
    var block = allBlocks[i];
    if (!block.flow) continue;
    if (!selectedByFlow[block.flow]) selectedByFlow[block.flow] = [];
    if (selectedByFlow[block.flow].length < limitPerFlow) {
      selectedByFlow[block.flow].push(block);
    }
  }

  var results = [];
  var displayOrder = ['F1', 'F2', 'F3', 'F4', 'F5', 'F6'];
  for (var d = 0; d < displayOrder.length; d++) {
    var flow = displayOrder[d];
    var flowBlocks = selectedByFlow[flow] || [];
    flowBlocks.reverse();
    for (var fb = 0; fb < flowBlocks.length; fb++) {
      results.push(flowBlocks[fb]);
    }
  }
  return results;
}

function readLatestActivityBlocksByFlow_() {
  return readRecentActivityBlocksByFlow_(1, ACTIVITY_QA_CONFIG.LIVE_AUDIT_ROW_WINDOW);
}

function inspectActivityHeaderLink_(sheet, rowNum, referenceToken) {
  var findings = [];
  try {
    var richText = sheet.getRange(rowNum, 4).getRichTextValue();
    if (!richText || typeof richText.getRuns !== 'function') return findings;

    var runs = richText.getRuns();
    var linkedRuns = [];
    for (var i = 0; i < runs.length; i++) {
      if (runs[i].getLinkUrl && runs[i].getLinkUrl()) linkedRuns.push(runs[i]);
    }

    if (linkedRuns.length === 0) {
      findings.push('Header reference is not linked');
    } else if (linkedRuns.length > 1) {
      findings.push('Header has multiple linked text runs');
    } else if (String(linkedRuns[0].getText() || '').trim() !== referenceToken) {
      findings.push('Linked text is not limited to the reference token');
    }
  } catch (e) {
    findings.push('Could not inspect header link');
  }
  return findings;
}

function validateLiveActivityBlock_(sheet, block) {
  var findings = [];
  var header = block.header || [];
  var details = String(header[3] || '').trim();
  var status = String(header[4] || '').trim();
  var headerError = String(header[5] || '').trim();
  var refMatch = details.match(/^(\S+)/);
  var referenceToken = refMatch ? refMatch[1] : '';

  if (!String(header[0] || '').trim()) findings.push('Header ID is blank');
  if (!String(header[1] || '').trim()) findings.push('Header time is blank');
  if (!block.flow) findings.push('Flow label does not map to a known flow');
  if (!details) findings.push('Header details are blank');
  if (!status) findings.push('Header status is blank');
  if (details && details.indexOf('  ') < 0) findings.push('Header details do not use canonical double-space separators');
  if (/\b\d+\s+(failed|skipped|warning|warnings)\b/i.test(details)) findings.push('Header details still include status or error counts');
  if (/^\d+\s+(failed|skipped|warning|warnings)(,\s*\d+\s+\w+)*$/i.test(headerError)) findings.push('Header error uses deprecated count-only summary');
  if (referenceToken) findings = findings.concat(inspectActivityHeaderLink_(sheet, block.headerRow, referenceToken));

  var sawParent = false;
  var sawNestedAfterParent = false;

  for (var i = 0; i < block.subRows.length; i++) {
    var row = block.subRows[i];
    var vals = row.values;
    var detail = String(vals[3] || '');
    var subStatus = String(vals[4] || '').trim();
    var subError = String(vals[5] || '').trim();

    if (String(vals[0] || '').trim() || String(vals[1] || '').trim() || String(vals[2] || '').trim()) {
      findings.push('Sub-item row has unexpected data in columns A:C');
      break;
    }
    if (!/^\s+/.test(detail)) findings.push('Sub-item row is missing tree indentation');
    if (subError && !subStatus) findings.push('Error sub-item row is missing status');
    if (subStatus === 'Skipped' && subError) findings.push('Skipped row should not carry error text');

    if (/\(\d+\s+batches\)/.test(detail)) {
      sawParent = true;
      sawNestedAfterParent = false;
    } else if (sawParent && detail.indexOf('│') >= 0) {
      sawNestedAfterParent = true;
    } else if (sawParent && !sawNestedAfterParent) {
      findings.push('Split-batch parent is not followed by nested batch rows');
      sawParent = false;
    }
  }

  if (sawParent && !sawNestedAfterParent) findings.push('Split-batch parent is not followed by nested batch rows');

  return findings;
}

function auditLatestActivityBlocks_() {
  var ss = typeof getDebugSpreadsheet_ === 'function'
    ? getDebugSpreadsheet_()
    : SpreadsheetApp.openById(CONFIG.DEBUG_SHEET_ID);
  var sheet = ss.getSheetByName('Activity');
  var blocks = readLatestActivityBlocksByFlow_();
  var results = [];

  for (var i = 0; i < blocks.length; i++) {
    var block = blocks[i];
    var findings = validateLiveActivityBlock_(sheet, block);
    results.push({
      scope: 'Live',
      flow: block.flow || '?',
      scenario: 'latest live block',
      ref: String(block.header[3] || '').split(/\s{2,}/)[0] || '',
      score: findings.length === 0 ? 100 : Math.max(0, 100 - findings.length * 15),
      findings: findings,
      rowNum: block.headerRow,
      notes: String(block.header[0] || '')
    });
  }

  for (var flow in FLOW_LABELS) {
    var found = false;
    for (var r = 0; r < results.length; r++) {
      if (results[r].flow === flow) {
        found = true;
        break;
      }
    }
    if (!found) {
      results.push({
        scope: 'Live',
        flow: flow,
        scenario: 'latest live block',
        ref: '',
        score: 0,
        findings: ['No live Activity block found for this flow in the last ' + ACTIVITY_QA_CONFIG.LIVE_AUDIT_ROW_WINDOW + ' rows'],
        rowNum: '',
        notes: ''
      });
    }
  }

  return results;
}

function copyActivityBlockToSheet_(sourceSheet, destSheet, block, destStartRow) {
  var rowCount = block.subRows.length + 1;
  var sourceRange = sourceSheet.getRange(block.headerRow, 1, rowCount, 6);
  var destRange = destSheet.getRange(destStartRow, 1, rowCount, 6);
  var richTextValues = sourceRange.getRichTextValues();

  sourceRange.copyFormatToRange(destSheet, 1, 6, destStartRow, destStartRow + rowCount - 1);
  destRange.setRichTextValues(richTextValues);
}

function getActivityReferenceToken_(header) {
  var details = String((header && header[3]) || '').trim();
  var match = details.match(/^(\S+)/);
  return match ? match[1] : '';
}

function describeRealSampleSource_(block) {
  if (block.flow === 'F1') return 'Source: Katana PO webhook + Katana API';
  if (block.flow === 'F2') return 'Source: WASP callout via Webhook Queue';
  if (block.flow === 'F3') return 'Source: Katana stock transfer polling';
  if (block.flow === 'F4') return 'Source: Katana MO webhook + Katana API';
  if (block.flow === 'F5') return 'Source: ShipStation shipment processing';
  if (block.flow === 'F6') return 'Source: Katana SO webhook or F6 staging path';
  return 'Source: live Activity';
}

function readRecentWebhookQueueEntries_(rowWindow) {
  var ss = typeof getDebugSpreadsheet_ === 'function'
    ? getDebugSpreadsheet_()
    : SpreadsheetApp.openById(CONFIG.DEBUG_SHEET_ID);
  var sheet = ss.getSheetByName('Webhook Queue');
  if (!sheet || sheet.getLastRow() < 2) return [];

  var totalRows = sheet.getLastRow() - 1;
  var rowsToRead = Math.min(totalRows, rowWindow || ACTIVITY_QA_CONFIG.WEBHOOK_AUDIT_ROW_WINDOW);
  var startRow = Math.max(2, sheet.getLastRow() - rowsToRead + 1);
  var values = sheet.getRange(startRow, 1, rowsToRead, 5).getDisplayValues();
  var entries = [];

  for (var i = values.length - 1; i >= 0; i--) {
    var row = values[i];
    var payload = null;
    try { payload = row[4] ? JSON.parse(row[4]) : null; } catch (e) { payload = null; }
    entries.push({
      rowNum: startRow + i,
      timestamp: row[0],
      action: row[1],
      status: row[2],
      result: row[3],
      payload: payload
    });
  }

  return entries;
}

function normalizeReferenceBase_(ref) {
  var text = String(ref || '').trim();
  if (!text) return '';
  text = text.replace(/\/\d+$/, '');
  return text;
}

function extractResultMessage_(resultText) {
  var text = String(resultText || '').trim();
  var idx = text.indexOf(':');
  return idx >= 0 ? text.substring(idx + 1).trim() : text;
}

function verifyQueueKatanaRef_(queueEntries, actions, fetchFn, refToken, cache, refBuilder) {
  var refBase = normalizeReferenceBase_(refToken);
  var checked = 0;

  for (var i = 0; i < queueEntries.length; i++) {
    var entry = queueEntries[i];
    if (actions.indexOf(entry.action) < 0) continue;
    var id = entry.payload && entry.payload.object ? entry.payload.object.id : null;
    if (!id) continue;
    checked++;
    if (checked > 30) break;

    var key = String(id);
    if (!Object.prototype.hasOwnProperty.call(cache, key)) {
      var fetched = null;
      try { fetched = fetchFn(id); } catch (e) { fetched = null; }
      cache[key] = fetched ? (fetched.data || fetched) : null;
    }

    var obj = cache[key];
    if (!obj) continue;
    var candidateRef = normalizeReferenceBase_(refBuilder(obj, id));
    if (candidateRef === refBase) {
      return 'Verified via queue row ' + entry.rowNum + ' and Katana id ' + id;
    }
  }

  return 'No recent queue match found';
}

function verifyF3RealSampleSource_(refToken, caches) {
  var refBase = normalizeReferenceBase_(refToken);
  if (!caches.transferList) {
    var sinceDate = new Date();
    sinceDate.setDate(sinceDate.getDate() - 30);
    caches.transferList = fetchKatanaStockTransfers(sinceDate.toISOString());
  }

  for (var i = 0; i < caches.transferList.length; i++) {
    var transfer = caches.transferList[i];
    var stRef = normalizeReferenceBase_(transfer.stock_transfer_number || ('ST-' + transfer.id));
    if (stRef === refBase) {
      return 'Verified in Katana stock transfers (id ' + transfer.id + ')';
    }
  }

  return 'No recent Katana stock transfer match found';
}

function verifyF5RealSampleSource_(refToken, caches) {
  var orderNumber = String(refToken || '').replace(/^#/, '').trim();
  if (!orderNumber) return 'No order number parsed';
  if (Object.prototype.hasOwnProperty.call(caches.shipstation, orderNumber)) {
    return caches.shipstation[orderNumber];
  }

  try {
    var response = callShipStationAPI('/shipments?orderNumber=' + encodeURIComponent(orderNumber) + '&includeShipmentItems=true', 'GET');
    var shipments = response && response.data ? (response.data.shipments || []) : [];
    caches.shipstation[orderNumber] = (response && response.code === 200 && shipments.length > 0)
      ? 'Verified in ShipStation (' + shipments.length + ' shipment' + (shipments.length !== 1 ? 's' : '') + ')'
      : 'ShipStation shipment not found';
  } catch (e) {
    caches.shipstation[orderNumber] = 'ShipStation lookup failed: ' + e.message;
  }

  return caches.shipstation[orderNumber];
}

function verifyRealSampleSource_(block, caches) {
  var refToken = getActivityReferenceToken_(block.header);

  if (block.flow === 'F1') {
    return verifyQueueKatanaRef_(
      caches.queueEntries,
      ['purchase_order.received', 'purchase_order.partially_received'],
      fetchKatanaPO,
      refToken,
      caches.katanaPO,
      function(po, id) { return po.order_no || ('PO-' + id); }
    );
  }

  if (block.flow === 'F2') {
    return 'Webhook Queue / WASP sample source; direct SA ref verification not implemented';
  }

  if (block.flow === 'F3') {
    return verifyF3RealSampleSource_(refToken, caches);
  }

  if (block.flow === 'F4') {
    return verifyQueueKatanaRef_(
      caches.queueEntries,
      ['manufacturing_order.done', 'manufacturing_order.completed', 'manufacturing_order.updated', 'manufacturing_order.deleted'],
      fetchKatanaMO,
      refToken,
      caches.katanaMO,
      function(mo, id) { return mo.order_no || ('MO-' + id); }
    );
  }

  if (block.flow === 'F5') {
    return verifyF5RealSampleSource_(refToken, caches);
  }

  if (block.flow === 'F6') {
    if (refToken.indexOf('MO-') === 0) {
      return verifyQueueKatanaRef_(
        caches.queueEntries,
        ['manufacturing_order.done', 'manufacturing_order.completed'],
        fetchKatanaMO,
        refToken,
        caches.katanaMO,
        function(mo, id) { return mo.order_no || ('MO-' + id); }
      );
    }

    return verifyQueueKatanaRef_(
      caches.queueEntries,
      ['sales_order.delivered', 'sales_order.cancelled', 'sales_order.updated'],
      fetchKatanaSalesOrder,
      refToken,
      caches.katanaSO,
      function(so, id) { return so.order_no || so.name || ('SO-' + id); }
    );
  }

  return 'No source verification available';
}

function calculateActivitySuitePassRate_(rows) {
  if (!rows || rows.length === 0) return 0;
  var total = 0;
  for (var i = 0; i < rows.length; i++) total += Number(rows[i].score || 0);
  return Math.round(total / rows.length);
}

function writeActivityQAReport_(sheet, fixtureResults, liveAuditResults) {
  var fixtureRate = calculateActivitySuitePassRate_(fixtureResults);
  var liveRate = calculateActivitySuitePassRate_(liveAuditResults);

  sheet.getRange(3, 1, 1, 8).setValues([[
    'Summary',
    '',
    'Fixtures ' + fixtureRate + '%',
    'Live ' + liveRate + '%',
    '',
    '',
    '',
    'Preview: ' + ACTIVITY_QA_CONFIG.PREVIEW_SHEET_NAME
  ]]);
  sheet.getRange(3, 1, 1, 8)
    .setBackground('#DDE7F3')
    .setFontWeight('bold');

  var rows = [];
  for (var i = 0; i < fixtureResults.length; i++) {
    rows.push(formatActivityQAReportRow_(fixtureResults[i]));
  }
  for (var j = 0; j < liveAuditResults.length; j++) {
    rows.push(formatActivityQAReportRow_(liveAuditResults[j]));
  }

  if (rows.length > 0) {
    sheet.getRange(5, 1, rows.length, 8).setValues(rows);
    for (var r = 0; r < rows.length; r++) {
      var bg = rows[r][3] >= 90 ? ACTIVITY_QA_CONFIG.OK_BG : rows[r][3] >= 70 ? ACTIVITY_QA_CONFIG.WARN_BG : ACTIVITY_QA_CONFIG.FAIL_BG;
      sheet.getRange(r + 5, 4).setBackground(bg);
      if (rows[r][4]) sheet.getRange(r + 5, 5).setWrap(true);
    }
  }
}

function writeActivityRealSampleReport_(sheet, rows) {
  var okCount = 0;
  for (var i = 0; i < rows.length; i++) {
    if (!rows[i].findings || rows[i].findings.length === 0) okCount++;
  }

  sheet.getRange(3, 1, 1, 8).setValues([[
    'Summary',
    '',
    rows.length + ' samples',
    rows.length ? Math.round((okCount / rows.length) * 100) : 0,
    '',
    '',
    '',
    'Last ' + ACTIVITY_QA_CONFIG.REAL_SAMPLE_LIMIT_PER_FLOW + ' rebuilt real samples per flow'
  ]]);
  sheet.getRange(3, 1, 1, 8)
    .setBackground('#DDE7F3')
    .setFontWeight('bold');

  var reportRows = [];
  for (var r = 0; r < rows.length; r++) {
    reportRows.push(formatActivityQAReportRow_(rows[r]));
  }

  if (reportRows.length > 0) {
    sheet.getRange(5, 1, reportRows.length, 8).setValues(reportRows);
    for (var x = 0; x < reportRows.length; x++) {
      var bg = reportRows[x][3] >= 90 ? ACTIVITY_QA_CONFIG.OK_BG : reportRows[x][3] >= 70 ? ACTIVITY_QA_CONFIG.WARN_BG : ACTIVITY_QA_CONFIG.FAIL_BG;
      sheet.getRange(x + 5, 4).setBackground(bg);
      if (reportRows[x][4]) sheet.getRange(x + 5, 5).setWrap(true);
      if (reportRows[x][7]) sheet.getRange(x + 5, 8).setWrap(true);
    }
  }
}

function formatActivityQAReportRow_(result) {
  return [
    result.scope || '',
    result.flow || '',
    result.scenario || '',
    Number(result.score || 0),
    result.findings && result.findings.length ? result.findings.join(' | ') : 'OK',
    result.ref || '',
    result.rowNum || '',
    result.notes || ''
  ];
}
