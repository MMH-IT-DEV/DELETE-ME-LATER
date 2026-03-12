// ============================================
// 12_SyncSheetFormat.gs - SYNC SHEET DARK THEME
// ============================================
// Applies the command-center dark theme to all tabs
// in the sync Google Sheet. Matches the design from
// the Queue tab in the system-architecture sheet.
//
// Run: formatSyncSheet() from the GAS editor.
//
// Dependencies:
//   SYNC_CONFIG.SYNC_SHEET_ID — 00_Config.gs
// ============================================

var THEME = {
  BG_DARK:    '#1c2333',
  BG_HEADER:  '#232c3d',
  TEXT_WHITE:  '#ffffff',
  TEXT_MUTED:  '#8b949e',
  TEXT_TITLE:  '#ffffff',
  ACCENT_BLUE: '#1f6feb',
  STAT_GREEN:  '#3fb950',
  STAT_AMBER:  '#d29922',
  STAT_RED:    '#f85149',
  STAT_GRAY:   '#8b949e',
  STATUS_OK:      '#1a7f37',
  STATUS_ERROR:   '#cf222e',
  STATUS_PENDING: '#9a6700',
  STATUS_SKIP:    '#57606a'
};

/**
 * Apply the dark command-center theme to ALL tabs in the sync sheet.
 * Run this from the GAS editor after the tabs are populated.
 */
function formatSyncSheet() {
  var ss = SpreadsheetApp.openById(SYNC_CONFIG.SYNC_SHEET_ID);
  var tabs = ss.getSheets();

  for (var t = 0; t < tabs.length; t++) {
    var tab = tabs[t];
    var name = tab.getName();
    Logger.log('Formatting: ' + name);

    if (name === 'WASP Export') {
      formatWaspExport_(tab);
    } else if (name === 'Item Map') {
      formatItemMap_(tab);
    } else if (name === 'Zero Plan') {
      formatZeroPlan_(tab);
    } else if (name === 'Re-Add Plan') {
      formatReAddPlan_(tab);
    } else if (name === 'Results') {
      formatResults_(tab);
    }
  }

  SpreadsheetApp.flush();
  Logger.log('===== FORMAT COMPLETE =====');
}

// ============================================
// SHARED FORMATTING HELPERS
// ============================================

/**
 * Apply dark background + white text to the entire sheet area.
 * Covers used range + padding so empty cells match too.
 */
function applyDarkBase_(tab) {
  var lastRow = Math.max(tab.getLastRow(), 100);
  var lastCol = Math.max(tab.getLastColumn(), 12);
  var fullRange = tab.getRange(1, 1, lastRow, lastCol);

  fullRange.setBackground(THEME.BG_DARK);
  fullRange.setFontColor(THEME.TEXT_WHITE);
  fullRange.setFontFamily('Roboto');
  fullRange.setFontSize(10);
  fullRange.setVerticalAlignment('middle');

  // Set tab color
  tab.setTabColor(THEME.ACCENT_BLUE);

  // Clear ALL existing conditional format rules (remove light-theme artifacts)
  tab.setConditionalFormatRules([]);
}

/**
 * Format the title area (rows 1-5) for tabs with title rows.
 */
function formatTitleArea_(tab) {
  // Row 1 — title: bold, white, 14pt
  var titleCell = tab.getRange(1, 2);
  titleCell.setFontSize(14);
  titleCell.setFontWeight('bold');
  titleCell.setFontColor(THEME.TEXT_TITLE);

  // Row 2 — subtitle: muted gray, normal weight, 10pt (NOT italic)
  var subtitleCell = tab.getRange(2, 2);
  subtitleCell.setFontSize(10);
  subtitleCell.setFontWeight('normal');
  subtitleCell.setFontColor(THEME.TEXT_MUTED);
  subtitleCell.setFontStyle('normal');

  // Row 4 — stats: bold, 11pt, with per-cell colors
  var statsRange = tab.getRange(4, 2, 1, 5);
  statsRange.setFontWeight('bold');
  statsRange.setFontSize(11);
}

/**
 * Color individual stat cells in row 4.
 * Pass an array of {col, color} pairs.
 */
function colorStats_(tab, statColors) {
  for (var i = 0; i < statColors.length; i++) {
    var sc = statColors[i];
    tab.getRange(4, sc.col).setFontColor(sc.color);
  }
}

/**
 * Format the header row (row 6) — consistent dark header across all tabs.
 */
function formatHeaderRow_(tab) {
  var lastCol = Math.max(tab.getLastColumn(), 10);
  var headerRange = tab.getRange(6, 1, 1, lastCol);
  headerRange.setBackground(THEME.BG_HEADER);
  headerRange.setFontColor(THEME.TEXT_TITLE);
  headerRange.setFontWeight('bold');
  headerRange.setFontSize(10);

  // Freeze title + header rows, and first column
  tab.setFrozenRows(6);
  tab.setFrozenColumns(1);
}

/**
 * Apply status-column conditional formatting with dark-theme colors.
 */
function applyStatusColors_(tab, statusCol, startRow) {
  var lastRow = Math.max(tab.getLastRow(), 500);
  var range = tab.getRange(startRow, statusCol, lastRow - startRow + 1, 1);

  var rules = [];

  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('OK')
      .setBackground(THEME.STATUS_OK)
      .setFontColor('#ffffff')
      .setRanges([range])
      .build()
  );
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('SYNC')
      .setBackground(THEME.STATUS_OK)
      .setFontColor('#ffffff')
      .setRanges([range])
      .build()
  );
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('PENDING')
      .setBackground(THEME.STATUS_PENDING)
      .setFontColor('#ffffff')
      .setRanges([range])
      .build()
  );
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('ERROR')
      .setBackground(THEME.STATUS_ERROR)
      .setFontColor('#ffffff')
      .setRanges([range])
      .build()
  );
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('SKIP')
      .setBackground(THEME.STATUS_SKIP)
      .setFontColor('#ffffff')
      .setRanges([range])
      .build()
  );

  tab.setConditionalFormatRules(rules);
}

// ============================================
// PER-TAB FORMATTERS
// ============================================

function formatWaspExport_(tab) {
  applyDarkBase_(tab);

  // WASP Export has headers at row 1 (no title rows)
  var lastCol = Math.max(tab.getLastColumn(), 12);
  var headerRange = tab.getRange(1, 1, 1, lastCol);
  headerRange.setBackground(THEME.BG_HEADER);
  headerRange.setFontColor(THEME.TEXT_TITLE);
  headerRange.setFontWeight('bold');
  tab.setFrozenRows(1);

  tab.setColumnWidth(1, 140);
  tab.setColumnWidth(2, 280);
  tab.setColumnWidth(3, 100);
  tab.setColumnWidth(4, 110);
  tab.setColumnWidth(5, 100);
  tab.setColumnWidth(6, 100);
  tab.setColumnWidth(7, 150);
  tab.setColumnWidth(8, 140);
}

function formatItemMap_(tab) {
  applyDarkBase_(tab);
  formatTitleArea_(tab);
  formatHeaderRow_(tab);

  // Stats colors: "137 SYNC" green, "162 SKIP" gray, "299 Total" white
  colorStats_(tab, [
    { col: 2, color: THEME.STAT_GREEN },
    { col: 3, color: THEME.STAT_GRAY },
    { col: 4, color: THEME.TEXT_WHITE }
  ]);

  tab.setColumnWidth(1, 140);
  tab.setColumnWidth(2, 280);
  tab.setColumnWidth(3, 170);
  tab.setColumnWidth(4, 140);
  tab.setColumnWidth(5, 150);
  tab.setColumnWidth(6, 70);
  tab.setColumnWidth(7, 250);

  // Action column (F = col 6)
  applyStatusColors_(tab, 6, 7);
}

function formatZeroPlan_(tab) {
  applyDarkBase_(tab);
  formatTitleArea_(tab);
  formatHeaderRow_(tab);

  // Stats colors: pending=amber, ok=green, error=red, skip=gray
  colorStats_(tab, [
    { col: 2, color: THEME.STAT_AMBER },
    { col: 3, color: THEME.STAT_GREEN },
    { col: 4, color: THEME.STAT_RED },
    { col: 5, color: THEME.STAT_GRAY }
  ]);

  tab.setColumnWidth(1, 140);
  tab.setColumnWidth(2, 250);
  tab.setColumnWidth(3, 140);
  tab.setColumnWidth(4, 150);
  tab.setColumnWidth(5, 100);
  tab.setColumnWidth(6, 100);
  tab.setColumnWidth(7, 130);
  tab.setColumnWidth(8, 80);

  // Status column (H = col 8)
  applyStatusColors_(tab, 8, 7);
}

function formatReAddPlan_(tab) {
  applyDarkBase_(tab);
  formatTitleArea_(tab);
  formatHeaderRow_(tab);

  // Stats colors: pending=amber, ok=green, error=red
  colorStats_(tab, [
    { col: 2, color: THEME.STAT_AMBER },
    { col: 3, color: THEME.STAT_GREEN },
    { col: 4, color: THEME.STAT_RED }
  ]);

  tab.setColumnWidth(1, 140);
  tab.setColumnWidth(2, 250);
  tab.setColumnWidth(3, 160);
  tab.setColumnWidth(4, 140);
  tab.setColumnWidth(5, 150);
  tab.setColumnWidth(6, 100);
  tab.setColumnWidth(7, 130);
  tab.setColumnWidth(8, 100);
  tab.setColumnWidth(9, 100);
  tab.setColumnWidth(10, 80);

  // Status column (J = col 10)
  applyStatusColors_(tab, 10, 7);
}

function formatResults_(tab) {
  applyDarkBase_(tab);
  formatTitleArea_(tab);
  formatHeaderRow_(tab);

  // Stats colors: total=white, ok=green, error=red
  colorStats_(tab, [
    { col: 2, color: THEME.TEXT_WHITE },
    { col: 3, color: THEME.STAT_GREEN },
    { col: 4, color: THEME.STAT_RED }
  ]);

  tab.setColumnWidth(1, 160);
  tab.setColumnWidth(2, 80);
  tab.setColumnWidth(3, 120);
  tab.setColumnWidth(4, 140);
  tab.setColumnWidth(5, 150);
  tab.setColumnWidth(6, 90);
  tab.setColumnWidth(7, 130);
  tab.setColumnWidth(8, 70);
  tab.setColumnWidth(9, 200);
  tab.setColumnWidth(10, 250);

  // Status column (H = col 8)
  applyStatusColors_(tab, 8, 7);
}
