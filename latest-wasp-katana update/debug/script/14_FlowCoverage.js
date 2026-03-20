// ============================================
// 14_FlowCoverage.gs — FLOW COVERAGE PANEL
// ============================================
// Counts events received vs logged for each flow and
// writes a live summary panel to the Activity sheet header.
//
// Panel location: Activity tab, rows 1-2, columns H-N (right of existing 6 cols)
// Runs daily at 08:00 via time-based trigger.
// Also runnable manually: IT Menu → Update Flow Coverage Panel
//
// Accuracy approach per flow:
//   F1 (webhook)  — Webhook Queue received  vs  Activity logged
//   F2 (callout)  — Webhook Queue received / processed / skipped (3-way)
//   F3 (polling)  — Activity logged only (polling = inherently complete)
//   F4 (webhook)  — Webhook Queue received  vs  Activity logged
//   F5 (polling)  — Activity logged only
//   F6 (polling)  — Activity logged only
//
// Receiving count uses the Webhook Queue tab (already captures every POST).
// Logged count uses the Activity tab (confirms the full F-flow completed).
// Any difference = event received but not fully processed = GAP.
// ============================================

// ============================================
// TRIGGER MANAGEMENT
// ============================================

/**
 * Create a daily 08:00 trigger for updateFlowCoveragePanel().
 * Run once from the Apps Script editor.
 */
function setupFlowCoverageTrigger() {
  // Remove existing coverage triggers first
  removeFlowCoverageTrigger();
  ScriptApp.newTrigger('updateFlowCoveragePanel')
    .timeBased()
    .atHour(8)
    .everyDays(1)
    .create();
  Logger.log('Flow coverage trigger created: daily at 08:00');
}

function removeFlowCoverageTrigger() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'updateFlowCoveragePanel') {
      ScriptApp.deleteTrigger(t);
    }
  });
}

// ============================================
// MAIN ENTRY POINT
// ============================================

/**
 * Gather today's flow counts and write the coverage panel to
 * Activity tab rows 1-2, columns H-N.
 * Safe to call at any time — only touches cols H-N rows 1-2.
 */
function updateFlowCoveragePanel() {
  var tz      = Session.getScriptTimeZone();
  var today   = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  var timeStr = Utilities.formatDate(new Date(), tz, 'HH:mm');

  var ss            = SpreadsheetApp.openById(CONFIG.DEBUG_SHEET_ID);
  var activitySheet = ss.getSheetByName('Activity');
  var queueSheet    = ss.getSheetByName('Webhook Queue');

  if (!activitySheet) {
    Logger.log('updateFlowCoveragePanel: Activity tab not found');
    return;
  }

  var data = getCoverageData_(activitySheet, queueSheet, today);
  writeFlowPanel_(activitySheet, data, today, timeStr);

  Logger.log('Flow coverage panel updated for ' + today);
}

// ============================================
// DATA GATHERING
// ============================================

/**
 * Build coverage data object for each flow.
 * Returns an array of flow-stat objects in display order.
 */
function getCoverageData_(activitySheet, queueSheet, today) {
  // ── Read Activity tab once ────────────────────────────────────────────
  var actLastRow  = activitySheet.getLastRow();
  var actData     = actLastRow > 3
    ? activitySheet.getRange(4, 1, actLastRow - 3, 3).getValues()
    : [];

  // Activity counts: header rows only (col A has WK-xxx, col C has flow label)
  var actCounts = {};
  for (var i = 0; i < actData.length; i++) {
    var id   = String(actData[i][0]).trim();
    var time = String(actData[i][1]).trim();
    var flow = String(actData[i][2]).trim();
    if (!id || !flow) continue;          // sub-item rows have empty id
    if (time.substring(0, 10) !== today) continue;
    actCounts[flow] = (actCounts[flow] || 0) + 1;
  }

  // ── Read Webhook Queue tab once ───────────────────────────────────────
  var wqData = [];
  if (queueSheet) {
    var wqLast = queueSheet.getLastRow();
    if (wqLast > 1) {
      wqData = queueSheet.getRange(2, 1, wqLast - 1, 4).getValues();
      // cols: Timestamp(0) | Action(1) | Status(2) | Result(3)
    }
  }

  // Helper: count Webhook Queue rows matching action list + optional status filter
  function wqCount(actions, statusFilter) {
    var n = 0;
    for (var j = 0; j < wqData.length; j++) {
      var ts     = String(wqData[j][0]).trim();
      var action = String(wqData[j][1]).trim();
      var status = String(wqData[j][2]).trim().toLowerCase();
      if (ts.substring(0, 10) !== today) continue;
      if (actions.indexOf(action) < 0) continue;
      if (statusFilter && statusFilter.indexOf(status) < 0) continue;
      n++;
    }
    return n;
  }

  // ── Per-flow stats ────────────────────────────────────────────────────

  // F1 — PO Receiving (webhook)
  var f1Received = wqCount(
    (typeof HOTFIX_FLAGS !== 'undefined' && HOTFIX_FLAGS.FLOW_COVERAGE_COUNT_PARTIAL_PO)
      ? ['purchase_order.received', 'purchase_order.partially_received']
      : ['purchase_order.received']
  );
  var f1Logged   = actCounts['F1 Receiving'] || 0;

  // F2 — WASP Adjustments (callout) — 3-way breakdown
  var f2Actions   = ['quantity_added', 'quantity_removed'];
  var f2Received  = wqCount(f2Actions);
  var f2Skipped   = wqCount(f2Actions, ['skipped']);
  var f2Logged    = actCounts['F2 Adjustments'] || 0;
  var f2Errors    = wqCount(f2Actions, ['error']);
  // Pending = arrived but processing crashed before writing a final status
  var f2Pending   = wqCount(f2Actions, ['pending']);
  // Processed = received minus all known non-processed outcomes
  var f2Processed = f2Received - f2Skipped - f2Errors - f2Pending;

  // F3 — Stock Transfers (polling)
  var f3Logged = actCounts['F3 Transfers'] || 0;

  // F4 — Manufacturing (webhook)
  // Count any MO webhook that was NOT ignored (i.e. a done/completed event)
  var f4MoActions  = ['manufacturing_order.done', 'manufacturing_order.completed',
                      'manufacturing_order.updated'];
  var f4Received   = wqCount(f4MoActions) - wqCount(f4MoActions, ['ignored', 'skipped']);
  var f4Logged     = actCounts['F4 Manufacturing'] || 0;

  // F5 — Shipping (polling)
  var f5Logged = actCounts['F5 Shipping'] || 0;

  // F6 — Amazon FBA (polling)
  var f6Logged = actCounts['F6 Amazon FBA'] || 0;

  return [
    {
      label: 'F1', name: 'Receiving', type: 'webhook',
      received: f1Received, logged: f1Logged
    },
    {
      label: 'F2', name: 'Adjustments', type: 'callout',
      received: f2Received, processed: f2Processed,
      skipped: f2Skipped, errors: f2Errors, pending: f2Pending, logged: f2Logged
    },
    {
      label: 'F3', name: 'Transfers', type: 'poll',
      logged: f3Logged
    },
    {
      label: 'F4', name: 'Mfg', type: 'webhook',
      received: f4Received, logged: f4Logged
    },
    {
      label: 'F5', name: 'Shipping', type: 'poll',
      logged: f5Logged
    },
    {
      label: 'F6', name: 'Amazon', type: 'poll',
      logged: f6Logged
    }
  ];
}

// ============================================
// PANEL WRITING
// ============================================

/**
 * Write the coverage panel to Activity sheet cols H-N, rows 1-2.
 * Rows 3+ are untouched. Cols A-G are untouched.
 *
 * Layout:
 *   Row 1  H1 → spacer   I1:N1 merged → title bar   O1 → last-synced
 *   Row 2  H2 → spacer   I2-N2 → one box per flow   O2 → dark bg
 */
function writeFlowPanel_(sheet, flows, today, timeStr) {
  var HDR_BG = '#1C2333';   // matches existing header rows

  // Panel colour palette (dark theme to match sheet)
  var MATCH_BG  = '#1B3A2E'; var MATCH_TXT = '#6FCF97';  // green
  var WARN_BG   = '#3A2A00'; var WARN_TXT  = '#F2994A';  // amber
  var MISS_BG   = '#3A0000'; var MISS_TXT  = '#EB5757';  // red
  var NA_BG     = '#151E2B'; var NA_TXT    = '#4A6074';  // grey / no-ops

  // ── Row heights ───────────────────────────────────────────────────────
  sheet.setRowHeight(1, 22);
  sheet.setRowHeight(2, 38);

  // ── Column widths for panel cols ──────────────────────────────────────
  var panelWidths = { 8: 3, 9: 11, 10: 14, 11: 11, 12: 11, 13: 11, 14: 11, 15: 11 };
  for (var c in panelWidths) {
    sheet.setColumnWidth(Number(c), panelWidths[c]);
  }
  // col 8 = H = narrow spacer; cols 9-14 = I-N = flow boxes; col 15 = O = synced

  // ── Clear any leftover panel content from cols D-J (previous attempt) ─
  sheet.getRange(1, 4, 2, 12).breakApart();
  sheet.getRange(1, 4, 2, 4).setValue('').setBackground(HDR_BG)
    .setFontColor(HDR_BG).setBorder(false, false, false, false, false, false);

  // ── Row 1: title bar ─────────────────────────────────────────────────
  // Spacer H1 — dark bg, no content
  var g1 = sheet.getRange(1, 8);
  g1.setBackground(HDR_BG);
  g1.setValue('');

  // Merge I1:N1, set title
  var titleRange = sheet.getRange(1, 9, 1, 6); // I1:N1
  titleRange.merge();
  titleRange.getCell(1, 1)
    .setValue('FLOW COVERAGE  ·  ' + today)
    .setBackground('#111E2D')
    .setFontColor('#AABBCC')
    .setFontSize(8)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  // O1: last synced
  var syncCell = sheet.getRange(1, 15);
  syncCell.setValue('Synced\n' + timeStr)
    .setBackground('#111E2D')
    .setFontColor('#556677')
    .setFontSize(7)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setWrap(true);

  // ── Row 2: flow boxes ─────────────────────────────────────────────────
  // H2 spacer
  sheet.getRange(2, 8).setBackground(HDR_BG).setValue('');

  for (var fi = 0; fi < flows.length; fi++) {
    var f    = flows[fi];
    var col  = 9 + fi;   // I=9(F1), J=10(F2), K=11(F3), L=12(F4), M=13(F5), N=14(F6)
    var cell = sheet.getRange(2, col);

    var bg, txt, line1, line2, border;

    if (f.type === 'poll') {
      // Polling flows: just show count
      if (f.logged === 0) {
        bg = NA_BG; txt = NA_TXT;
        line1 = f.label + ' ' + f.name;
        line2 = '– no events';
        border = null;
      } else {
        bg = MATCH_BG; txt = MATCH_TXT;
        line1 = f.label + ' ' + f.name;
        line2 = f.logged + ' events ✓';
        border = '#1DB954';
      }

    } else if (f.type === 'callout') {
      // F2: 3-way breakdown
      if (f.received === 0) {
        bg = NA_BG; txt = NA_TXT;
        line1 = f.label + ' ' + f.name;
        line2 = '– no callouts';
        border = null;
      } else {
        var gap2 = f.pending || 0;  // pending = arrived but crashed before final status
        if (gap2 > 0) {
          bg = MISS_BG; txt = MISS_TXT; border = '#EB5757';
        } else if (f.errors > 0) {
          bg = WARN_BG; txt = WARN_TXT; border = '#F2994A';
        } else {
          bg = MATCH_BG; txt = MATCH_TXT; border = '#1DB954';
        }
        line1 = f.label + ' ' + f.name;
        line2 = f.received + ' in  ' + f.processed + ' ok';
        if (f.skipped > 0) line2 += '  ' + f.skipped + ' skip';
        if (f.errors  > 0) line2 += '  ' + f.errors  + ' err';
        if (gap2       > 0) line2 += '  ' + gap2      + ' stuck';
      }

    } else {
      // Webhook flows (F1, F4): received vs logged
      if (f.received === 0 && f.logged === 0) {
        bg = NA_BG; txt = NA_TXT;
        line1 = f.label + ' ' + f.name;
        line2 = '– no events';
        border = null;
      } else if (f.logged >= f.received) {
        bg = MATCH_BG; txt = MATCH_TXT; border = '#1DB954';
        line1 = f.label + ' ' + f.name;
        line2 = f.logged + '/' + f.received + ' ✓';
      } else if (f.logged > 0) {
        bg = WARN_BG; txt = WARN_TXT; border = '#F2994A';
        line1 = f.label + ' ' + f.name;
        line2 = f.logged + '/' + f.received + ' !';
      } else {
        bg = MISS_BG; txt = MISS_TXT; border = '#EB5757';
        line1 = f.label + ' ' + f.name;
        line2 = '0/' + f.received + ' ✗ MISS';
      }
    }

    cell.setValue(line1 + '\n' + line2)
      .setBackground(bg)
      .setFontColor(txt)
      .setFontSize(8)
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle')
      .setWrap(true);

    if (border) {
      var side = SpreadsheetApp.newTextStyle().build(); // reset
      cell.setBorder(true, true, true, true, false, false,
        border, SpreadsheetApp.BorderStyle.MEDIUM);
    } else {
      cell.setBorder(false, false, false, false, false, false);
    }
  }

  // O2 spacer
  sheet.getRange(2, 15).setBackground(HDR_BG).setValue('');

  // Row 3 in panel cols — keep dark header bg
  for (var pc = 8; pc <= 15; pc++) {
    sheet.getRange(3, pc).setBackground(HDR_BG);
  }

  SpreadsheetApp.flush();
}

// ============================================
// DIAGNOSTIC
// ============================================

/**
 * Run from Apps Script editor to check today's counts without
 * writing to the sheet. Output visible in View → Logs.
 */
function diagFlowCoverage() {
  var tz      = Session.getScriptTimeZone();
  var today   = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  var ss      = SpreadsheetApp.openById(CONFIG.DEBUG_SHEET_ID);
  var data    = getCoverageData_(
    ss.getSheetByName('Activity'),
    ss.getSheetByName('Webhook Queue'),
    today
  );

  Logger.log('=== Flow Coverage Diagnostic: ' + today + ' ===');
  for (var i = 0; i < data.length; i++) {
    var f = data[i];
    if (f.type === 'callout') {
      Logger.log(f.label + ' ' + f.name + ': received=' + f.received +
        ' processed=' + f.processed + ' skipped=' + f.skipped +
        ' errors=' + f.errors + ' pending=' + (f.pending || 0));
    } else if (f.type === 'webhook') {
      Logger.log(f.label + ' ' + f.name + ': received=' + f.received + ' logged=' + f.logged +
        (f.received > f.logged ? ' ← GAP: ' + (f.received - f.logged) + ' missed' : ' ✓'));
    } else {
      Logger.log(f.label + ' ' + f.name + ': ' + f.logged + ' events (polling)');
    }
  }
  Logger.log('=== END ===');
}
