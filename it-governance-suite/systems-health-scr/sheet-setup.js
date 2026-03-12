/**
 * Systems Health Tracker — Sheet Setup
 * Formats System Registry with Wasp-tab style + GMP compliance columns.
 * Column layout matches the approved XLSX template (19 columns).
 *
 * Run: Health Monitor menu → Format Registry Sheet
 * Safe to re-run — reads and preserves all existing system rows.
 *
 * CRITICAL: GAS V8 — var only, no const/let, no arrow functions
 */

// ─── Maintenance Guide hyperlinks (system name → Drive/doc URL) ───────────
// Any system listed here gets a clickable link in the Maintenance Guide col.
var SOP_LINKS = {
  'FedEx Dispute Bot':                  'https://drive.google.com/drive/folders/1ZI40dAsLXL6lFjZaUQ6esc6rlGTBSvte',
  'Katana → Google Calendar PO Sync':   'https://drive.google.com/drive/folders/190tw277_odKMPwKdO5CsFNpHFKX9tEwT',
  'QA Escalation Automation':           'https://drive.google.com/drive/folders/1YkMQvwUwhLn6Bo5rJfTBY0iXRdt3LVBU'
};

// ─── Separator bar colours per Type ───────────────────────────────────────
var SEP_COLORS = {
  'FLOW':         '01579B',
  'SCRIPT':       '004D40',
  'BOT':          '4A148C',
  'DASHBOARD': '1B5E20'
};

// ─── Platform chip colours (col D) — light bg + dark text ────────────────
var PLATFORM_COLORS = {
  'SHOPIFY':       { bg: 'B2DFDB', fg: '00695C' },
  'GOOGLE':        { bg: 'BBDEFB', fg: '0D47A1' },
  'GOOGLE SHEET': { bg: 'BBDEFB', fg: '0D47A1' },
  'FEDEX':         { bg: 'E1BEE7', fg: '4A148C' },
  'KATANA':        { bg: 'FFE0B2', fg: 'BF360C' },
  'WASP/KATANA':   { bg: 'C8E6C9', fg: '1B5E20' }
};

// ─── Column definitions (19 cols — matches approved XLSX template) ─────────
//  A  System Name        B  Description         C  Type
//  D  Platform           E  GMP Critical        F  Run Frequency
//  G  Maintenance Guide  H  Auth Type           I  Auth Location
//  J  Expiry Date        K  Days Left           L  Owner
//  M  Validated          N  Last Validated      O  Heartbeat Method
//  P  Status             Q  Last Heartbeat      R  Last Run OK
//  S  Notes

var REGISTRY_HEADERS = [
  'System Name',        // A  1
  'Description',        // B  2
  'Type',               // C  3
  'Platform',           // D  4
  'GMP Critical',       // E  5
  'Run Frequency',      // F  6
  'Maintenance Guide',  // G  7
  'Auth Type',          // H  8
  'Auth Location',      // I  9
  'Expiry Date',        // J  10
  'Days Left',          // K  11
  'Owner',              // L  12
  'Validated',          // M  13
  'Last Validated',     // N  14
  'Heartbeat Method',   // O  15
  'Status',             // P  16
  'Last Heartbeat',     // Q  17
  'Last Run OK',        // R  18
  'Notes'               // S  19
];

var COL_WIDTHS = [200, 320, 110, 110, 95, 120, 200, 105, 200, 105, 75, 90, 95, 115, 165, 105, 165, 100, 200];
var TYPE_ORDER = ['FLOW', 'SCRIPT', 'BOT', 'DASHBOARD'];
var NUM_COLS   = 19;

// Pre-filled system data — edit these rows directly to change what
// setupHealthSheet() writes. All existing sheet data is preserved if
// col C (Type) matches a known type; only formatting is re-applied.
var SYSTEMS = {
  'FLOW': [
    // ── Internal Shopify Flows (no external API) ───────────────────────────
    ['Repeat High-Volume Purchase Detection',
     'Flags repeat customers placing high-volume orders; applies internal review tag on order',
     'FLOW', 'Shopify', 'Yes', 'Per order (conditional)',
     'SOP-009-Shopify-Fraud-Flows', 'None',
     'N/A',
     '', '', 'Felippe', 'Yes', '', 'Flow HTTP step (both branches)', 'Unknown', '', '',
     'Conditional flow — add Send HTTP Request on BOTH branches (condition_met:true/false). Fires on every order created.'],

    ['High Quantity Order Review',
     'Flags orders with high item quantities for manual review before fulfillment',
     'FLOW', 'Shopify', 'Yes', 'Per order (conditional)',
     'SOP-009-Shopify-Fraud-Flows', 'None',
     'N/A',
     '', '', 'Felippe', 'Yes', '', 'Flow HTTP step (both branches)', 'Unknown', '', '',
     'Conditional flow — add Send HTTP Request on BOTH branches (condition_met:true/false). Fires on every order created.'],

    ['Add Gift When 2+ 4oz Products Ordered',
     'Automatically adds a promotional gift line item when a customer orders 2 or more 4oz products',
     'FLOW', 'Shopify', 'No', 'Per order (conditional)',
     'SOP-010-Shopify-Promo-Flows', 'None',
     'N/A',
     '', '', 'Felippe', 'Yes', '', 'Flow HTTP step (both branches)', 'Unknown', '', '',
     'Conditional flow — add Send HTTP Request on BOTH branches (condition_met:true/false). Fires on every order created.'],

    // ── Shopify → ShipStation API Flows ───────────────────────────────────
    ['Shopify Hold → ShipStation Hold Sync',
     'Puts ShipStation order on hold when Shopify order fulfillment order is placed on hold',
     'FLOW', 'Shopify', 'Yes', 'Per fulfillment hold event',
     'SOP-008-ShipStation-Shopify-Flow', 'Bearer Token',
     'Shopify Secrets → Secrets Manager',
     '2027-01-01', '', 'Felippe', 'Yes', '', 'Flow HTTP step (pre-condition)', 'Unknown', '', '',
     'Add Send HTTP Request as FIRST step (before any condition) so it fires on every fulfillment hold event.'],

    ['Shopify Hold Release → ShipStation Release',
     'Releases ShipStation hold when Shopify fulfillment order hold is released',
     'FLOW', 'Shopify', 'Yes', 'Per hold-release event',
     'SOP-008-ShipStation-Shopify-Flow', 'Bearer Token',
     'Shopify Secrets → Secrets Manager',
     '2027-01-01', '', 'Felippe', 'Yes', '', 'Flow HTTP step (pre-condition)', 'Unknown', '', '',
     'Add Send HTTP Request as FIRST step (before any condition) so it fires on every hold-release event.'],

    ['Shopify Cancel → ShipStation Cancel',
     'Cancels ShipStation order and removes any active hold when Shopify order is cancelled',
     'FLOW', 'Shopify', 'Yes', 'Per order cancel',
     'SOP-008-ShipStation-Shopify-Flow', 'Bearer Token',
     'Shopify Secrets → Secrets Manager',
     '2027-01-01', '', 'Felippe', 'Yes', '', 'Flow HTTP step (pre-condition)', 'Unknown', '', '',
     'Add Send HTTP Request as FIRST step (before any condition) so it fires on every order cancel.']
  ],

  'SCRIPT': [
    ['Katana → Google Calendar PO Sync',
     'Syncs Katana purchase orders to Google Calendar as dated events',
     'SCRIPT', 'Katana', 'Yes', 'Per PO event',
     'SOP-005-Katana-Calendar-PO-Sync', 'API Key',
     'Apps Script → Script Properties',
     'NO EXPIRY', '', 'Erik', 'Yes', '2026-01-26', 'Script sendHeartbeat()', 'Healthy',
     '2026-01-28 00:23', 'Yes', ''],

    ['QA Escalation Automation',
     'Escalates QA issues via Google Sheet triggers and Slack notifications',
     'SCRIPT', 'Google', 'Yes', 'Daily + on edit',
     'SOP-007-QA-Escalation-Bot', 'None', 'N/A',
     '', '', 'Erik', 'Yes', '', 'Script sendHeartbeat()', 'Healthy',
     '2026-01-27', 'Yes', ''],

  ],

  'BOT': [
    ['FedEx Dispute Bot',
     'Automatically files FedEx duty/tax disputes. Runs locally every Monday at 10AM via Windows Task Scheduler.',
     'BOT', 'FedEx', 'No', 'Weekly',
     'SOP-006-FedEx-Dispute-Bot', 'None', 'N/A',
     '', '', 'Erik', 'Yes', '', 'Python health_monitor.py', 'Healthy',
     '2026-03-02', 'Yes', 'GitHub: https://github.com/MMH-IT-DEV/MMH_FEDEX_DISPUTE_BOT | Drive: https://drive.google.com/drive/folders/1ZI40dAsLXL6lFjZaUQ6esc6rlGTBSvte']
  ],

  'DASHBOARD': [
    ['2026_Security & GMP Connection Tracker',
     'GMP access log, incident register, and periodic user access reviews',
     'DASHBOARD', 'Google Sheet', 'Yes', 'On demand',
     'SOP-007-Security-Tracker', 'None', 'N/A',
     '', '', 'Erik', 'Yes', '2026-01-01', 'N/A', 'Healthy', 'N/A', 'N/A', ''],

    ['2026_Katana-Wasp Inventory Sync',
     'Inventory sync monitoring dashboard — Katana MRP ↔ WASP InventoryCloud',
     'DASHBOARD', 'Wasp/Katana', 'Yes', 'Hourly (via engin-src)',
     'SOP-WASP-Katana-Sync', 'API Key x2',
     'Script Properties',
     '2027-12-31', '', 'Erik', 'Yes', '2026-01-01', 'Script sendHeartbeat()', 'Healthy', 'N/A', 'N/A',
     'Katana API: No expiry | WASP API: expires 2027-12-31'],

    ['2026-Katana-WASP_DebugLog',
     'Audit log and debug dashboard for WASP-Katana webhook processing — Activity, Flow Detail, Webhook Queue tabs',
     'DASHBOARD', 'Wasp/Katana', 'Yes', 'Per webhook event',
     'SOP-WASP-Katana-Sync', 'None',
     'N/A',
     '', '', 'Erik', 'Yes', '2026-01-01', 'Script sendHeartbeat()', 'Healthy', 'N/A', 'N/A', '']
  ]
};

// ─────────────────────────────────────────────────────────────────────────
// MAIN SETUP FUNCTION
// ─────────────────────────────────────────────────────────────────────────

function setupHealthSheet() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('System Registry');

  if (!sheet) {
    try {
      SpreadsheetApp.getUi().alert(
        '"System Registry" tab not found. Please rename your main tab to "System Registry" and try again.',
        SpreadsheetApp.getUi().ButtonSet.OK);
    } catch (e) { Logger.log('Sheet not found'); }
    return;
  }

  // 1. Read & merge existing row data with template defaults
  var mergedData = mergeWithExisting_(sheet);

  // 2. Clear everything — including data validations (sheet.clear() misses these)
  sheet.clear();
  try { sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).clearDataValidations(); } catch (e) {}
  try { sheet.clearConditionalFormatRules(); } catch (e) {}

  // 3. Dark header row
  var hRange = sheet.getRange(1, 1, 1, NUM_COLS);
  hRange.setValues([REGISTRY_HEADERS]);
  hRange.setBackground('#263238');
  hRange.setFontColor('#ffffff');
  hRange.setFontWeight('bold');
  hRange.setFontSize(10);
  sheet.setFrozenRows(1);
  sheet.row_dimensions = undefined; // not needed in GAS

  // Accent headers on GMP/Status/Heartbeat columns
  sheet.getRange(1, 5).setBackground('#37474f');   // GMP Critical
  sheet.getRange(1, 13).setBackground('#37474f');  // Validated
  sheet.getRange(1, 16).setBackground('#1a3a4a');  // Status
  sheet.getRange(1, 17).setBackground('#1a3a4a');  // Last Heartbeat
  sheet.getRange(1, 18).setBackground('#1a3a4a');  // Last Run OK

  // 4. Column widths
  for (var c = 0; c < COL_WIDTHS.length; c++) {
    sheet.setColumnWidth(c + 1, COL_WIDTHS[c]);
  }

  // 5. Write rows by type group
  var writeRow     = 2;
  var totalSystems = 0;

  for (var ti = 0; ti < TYPE_ORDER.length; ti++) {
    var typeName  = TYPE_ORDER[ti];
    var typeRows  = mergedData[typeName];
    if (!typeRows || typeRows.length === 0) continue;

    var sepHex = SEP_COLORS[typeName] || '424242';

    // ── Separator bar ──────────────────────────────────────────────────
    var sepVals = [];
    for (var sc = 0; sc < NUM_COLS; sc++) sepVals.push('');
    sepVals[2] = typeName; // col C — visible under the Type header

    var sepRange = sheet.getRange(writeRow, 1, 1, NUM_COLS);
    sepRange.setValues([sepVals]);
    sepRange.setBackground('#' + sepHex);
    sepRange.setFontColor('#ffffff');
    sepRange.setFontWeight('bold');
    sepRange.setFontSize(9);
    writeRow++;

    // ── Data rows ──────────────────────────────────────────────────────
    for (var dr = 0; dr < typeRows.length; dr++) {
      var row   = typeRows[dr];
      var rowBg = (dr % 2 === 0) ? '#f5f7f9' : '#ffffff';

      sheet.getRange(writeRow, 1, 1, NUM_COLS).setValues([row]);
      sheet.getRange(writeRow, 1, 1, NUM_COLS).setBackground(rowBg);

      // Platform chip  (col D = 4)
      applyChip_(sheet, writeRow, 4, String(row[3] || '').toUpperCase(), PLATFORM_COLORS, rowBg);

      // GMP Critical   (col E = 5)
      applyGmpChip_(sheet, writeRow, 5, String(row[4] || ''));

      // Validated      (col M = 13)
      applyValidatedChip_(sheet, writeRow, 13, String(row[12] || ''));

      // Heartbeat Method (col O = 15)
      applyHbChip_(sheet, writeRow, 15, String(row[14] || ''));

      // Status         (col P = 16)
      applyStatusStyle_(sheet.getRange(writeRow, 16), String(row[15] || ''));

      // Last Run OK    (col R = 18)
      applyRunOkChip_(sheet, writeRow, 18, String(row[17] || ''));

      writeRow++;
      totalSystems++;
    }
  }

  // 6. Dropdowns
  var lastRow = writeRow - 1;
  if (lastRow >= 2) {
    addDropdown_(sheet, 2, 3,  lastRow, ['FLOW', 'SCRIPT', 'BOT', 'DASHBOARD']);
    addDropdown_(sheet, 2, 5,  lastRow, ['Yes', 'No']);
    addDropdown_(sheet, 2, 8,  lastRow, ['API Key', 'API Key x2', 'Bearer Token', 'OAuth', 'Service Account', 'Apps Script Property', 'None']);
    addDropdown_(sheet, 2, 13, lastRow, ['Yes', 'No', 'Pending']);
    addDropdown_(sheet, 2, 16, lastRow, ['Healthy', 'Degraded', 'Down', 'Unknown']);
    addDropdown_(sheet, 2, 18, lastRow, ['Yes', 'No', 'N/A']);
  }

  // 7. Hyperlinks on Maintenance Guide col (G = 7) for systems in SOP_LINKS
  setSopLinks_(sheet);

  // 8. Days Left formulas — auto-calculates from Expiry Date each day
  setDaysLeftFormulas_(sheet);

  SpreadsheetApp.flush();

  try {
    SpreadsheetApp.getUi().alert('Done',
      totalSystems + ' systems formatted.\n\n' +
      'Next step: set HEALTH_MONITOR_URL in each script\'s Script Properties\n' +
      'so heartbeats can reach this sheet.',
      SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (e) {
    Logger.log('setupHealthSheet complete: ' + totalSystems + ' systems');
  }
}

// ─────────────────────────────────────────────────────────────────────────
// SET SOP HYPERLINKS — sets Maintenance Guide cell (col G=7) as a clickable
// link for any system that has an entry in SOP_LINKS above.
// ─────────────────────────────────────────────────────────────────────────
function setSopLinks_(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  var names = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  for (var i = 0; i < names.length; i++) {
    var name = String(names[i][0] || '').trim();
    if (SOP_LINKS[name]) {
      var cell = sheet.getRange(i + 2, 7);
      cell.setFormula('=HYPERLINK("' + SOP_LINKS[name] + '","' + cell.getValue() + '")');
      cell.setFontColor('#1155CC').setFontLine('underline');
    }
  }
}

// ─────────────────────────────────────────────────────────────────────────
// DAYS LEFT FORMULAS — sets =J{n}-TODAY() on every row that has a date in
// col J (Expiry Date). Colours cell red (<0) or yellow (≤30 days).
// Runs after all rows are written so row numbers are stable.
// ─────────────────────────────────────────────────────────────────────────
function setDaysLeftFormulas_(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  var today = new Date();
  today.setHours(0, 0, 0, 0);

  for (var r = 2; r <= lastRow; r++) {
    var expiryVal = sheet.getRange(r, 10).getValue(); // col J
    var dCell     = sheet.getRange(r, 11);            // col K

    if (!expiryVal || expiryVal === '') {
      dCell.setValue('');
      continue;
    }
    if (expiryVal === 'NO EXPIRY') {
      dCell.setValue('—');
      dCell.setFontColor('#546E7A');
      dCell.setHorizontalAlignment('center');
      continue;
    }

    // Set live formula
    dCell.setFormula('=J' + r + '-TODAY()');
    dCell.setNumberFormat('0');
    dCell.setHorizontalAlignment('center');

    // Compute days now for colour (formula recalculates daily, colour is set once)
    var expDate  = expiryVal instanceof Date ? expiryVal : new Date(expiryVal);
    var daysLeft = Math.round((expDate - today) / 86400000);

    if (daysLeft < 0) {
      dCell.setBackground('#5c0011').setFontColor('#f87171').setFontWeight('bold');
    } else if (daysLeft <= 30) {
      dCell.setBackground('#553600').setFontColor('#fbbf24').setFontWeight('bold');
    } else {
      dCell.setBackground(null).setFontColor('#000000').setFontWeight('normal');
    }
  }
}

// ─────────────────────────────────────────────────────────────────────────
// MERGE: overlay template defaults with whatever is already in the sheet
// (preserves live Status, Last Heartbeat, Last Run OK from existing rows)
// ─────────────────────────────────────────────────────────────────────────
function mergeWithExisting_(sheet) {
  // Build lookup of existing rows by System Name
  var existing = {};
  var lastRow  = sheet.getLastRow();
  if (lastRow > 1) {
    var numCols = Math.min(sheet.getLastColumn(), NUM_COLS);
    var raw     = sheet.getRange(2, 1, lastRow - 1, numCols).getValues();
    for (var i = 0; i < raw.length; i++) {
      var name = String(raw[i][0] || '').trim();
      if (name) existing[name] = raw[i];
    }
  }

  var merged = {};
  for (var ti = 0; ti < TYPE_ORDER.length; ti++) {
    var typeName  = TYPE_ORDER[ti];
    var templates = SYSTEMS[typeName] || [];
    merged[typeName] = [];

    for (var ri = 0; ri < templates.length; ri++) {
      var tmpl = templates[ri].slice(0);
      while (tmpl.length < NUM_COLS) tmpl.push('');

      var sysName = String(tmpl[0] || '').trim();
      var live    = existing[sysName];

      if (live) {
        // Preserve live operational columns: Status(16), Last HB(17), Last Run OK(18)
        if (live[15] && live[15] !== '') tmpl[15] = live[15]; // Status
        if (live[16] && live[16] !== '') tmpl[16] = live[16]; // Last Heartbeat
        if (live[17] && live[17] !== '') tmpl[17] = live[17]; // Last Run OK
      }

      merged[typeName].push(tmpl);
    }
  }
  return merged;
}

// ─────────────────────────────────────────────────────────────────────────
// CELL STYLING HELPERS
// ─────────────────────────────────────────────────────────────────────────

// Platform chip — light bg + dark text
function applyChip_(sheet, row, col, key, colorMap, rowBg) {
  var c = colorMap[key];
  if (!c) return;
  var cell = sheet.getRange(row, col);
  cell.setBackground('#' + c.bg).setFontColor('#' + c.fg).setFontWeight('bold');
}

// GMP Critical chip
function applyGmpChip_(sheet, row, col, val) {
  var cell = sheet.getRange(row, col);
  cell.setFontWeight('bold');
  if (val === 'Yes') {
    cell.setBackground('#C8E6C9').setFontColor('#1B5E20');
  } else {
    cell.setBackground('#ECEFF1').setFontColor('#546E7A');
  }
}

// Validated chip
function applyValidatedChip_(sheet, row, col, val) {
  var cell = sheet.getRange(row, col);
  cell.setFontWeight('bold');
  if (val === 'Yes') {
    cell.setBackground('#C8E6C9').setFontColor('#1B5E20');
  } else if (val === 'Pending') {
    cell.setBackground('#E3F2FD').setFontColor('#1565C0');
  } else if (val === 'No') {
    cell.setBackground('#FFCDD2').setFontColor('#C62828');
  }
}

// Heartbeat Method chip
function applyHbChip_(sheet, row, col, val) {
  var cell = sheet.getRange(row, col);
  if (val === 'N/A' || val === '') return;
  cell.setBackground('#E0F2F1').setFontColor('#00695C').setFontWeight('bold');
}

// Last Run OK chip
function applyRunOkChip_(sheet, row, col, val) {
  var cell = sheet.getRange(row, col);
  if (val === 'Yes') {
    cell.setBackground('#C8E6C9').setFontColor('#1B5E20').setFontWeight('bold');
  } else if (val === 'No') {
    cell.setBackground('#FFCDD2').setFontColor('#C62828').setFontWeight('bold');
  }
}

/**
 * Status cell — softer light bg + dark text. Called here and from health-checks.js.
 */
function applyStatusStyle_(cell, status) {
  var s = String(status || '').toLowerCase().trim();
  cell.setFontWeight('bold');
  if (s === 'healthy') {
    cell.setBackground('#E8F5E9').setFontColor('#2E7D32');
  } else if (s === 'degraded') {
    cell.setBackground('#FFF8E1').setFontColor('#F57F17');
  } else if (s === 'down') {
    cell.setBackground('#FFEBEE').setFontColor('#C62828');
  } else {
    cell.setBackground('#ECEFF1').setFontColor('#546E7A');
  }
}

function addDropdown_(sheet, startRow, col, endRow, options) {
  if (endRow < startRow) return;
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(options, true)
    .setAllowInvalid(true)
    .build();
  sheet.getRange(startRow, col, endRow - startRow + 1, 1).setDataValidation(rule);
}
