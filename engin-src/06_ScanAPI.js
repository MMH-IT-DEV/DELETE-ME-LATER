/**
 * 06_ScanAPI — Terminal sheet scanning endpoint
 * Bound to engin-src (Sync Sheet) — SpreadsheetApp.getActiveSpreadsheet() returns sync sheet.
 *
 * Setup (one-time):
 *   1. Project Settings → Script Properties → Add:
 *        Key:   SCAN_TOKEN
 *        Value: katana-scan-2026
 *   2. Deploy → New Deployment → Web App
 *        Execute as: Me
 *        Who has access: Anyone
 *   3. Copy the deployment URL into config/sheet-scan.json as "syncUrl"
 *
 * CRITICAL GAS RULES: var only, no const/let, no arrow functions, no template literals.
 */

function doGet(e) {
  var params = e ? (e.parameter || {}) : {};

  // ── Token check ────────────────────────────────────────────────────────────
  var expectedToken = PropertiesService.getScriptProperties().getProperty('SCAN_TOKEN');
  if (!expectedToken) {
    return jsonScanResponse({ error: 'SCAN_TOKEN not set in Script Properties' });
  }
  if (!params.token || params.token !== expectedToken) {
    return jsonScanResponse({ error: 'Unauthorized' });
  }

  // ── Tab param ──────────────────────────────────────────────────────────────
  var tabName = params.tab ? String(params.tab).trim() : '';
  if (!tabName) {
    return jsonScanResponse({ error: 'Missing ?tab= parameter' });
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // ── List tabs mode ─────────────────────────────────────────────────────────
  if (tabName === '__list__') {
    var sheets = ss.getSheets();
    var names = [];
    for (var i = 0; i < sheets.length; i++) {
      names.push(sheets[i].getName());
    }
    return jsonScanResponse({ tabs: names, spreadsheetId: ss.getId() });
  }

  // ── Read tab ───────────────────────────────────────────────────────────────
  try {
    var sheet = ss.getSheetByName(tabName);
    if (!sheet) {
      return jsonScanResponse({ error: 'Tab not found: ' + tabName });
    }

    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();

    if (lastRow === 0 || lastCol === 0) {
      return jsonScanResponse({ tab: tabName, rows: [] });
    }

    var values = sheet.getRange(1, 1, lastRow, lastCol).getValues();

    var rows = [];
    for (var r = 0; r < values.length; r++) {
      var row = [];
      for (var c = 0; c < values[r].length; c++) {
        var cell = values[r][c];
        if (cell instanceof Date) {
          row.push(Utilities.formatDate(cell, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss'));
        } else {
          row.push(cell);
        }
      }
      rows.push(row);
    }

    return jsonScanResponse({ tab: tabName, rows: rows });

  } catch (err) {
    return jsonScanResponse({ error: err.message });
  }
}

function jsonScanResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
