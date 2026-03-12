/**
 * 06_ScanAPI — Terminal sheet scanning endpoint
 *
 * Allows reading any tab from the terminal via:
 *   node scripts/sheet-scan.js --tab "Katana"
 *
 * Setup (one-time):
 *   1. Paste this file into the GAS editor as 06_ScanAPI.gs
 *   2. Project Settings → Script Properties → Add:
 *        Key:   SCAN_TOKEN
 *        Value: (any strong secret string you choose)
 *   3. Deploy → New Deployment → Web App
 *        Execute as: Me
 *        Who has access: Anyone
 *   4. Copy the deployment URL into config/sheet-scan.json
 *
 * CRITICAL GAS RULES: var only, no const/let, no arrow functions,
 *   no template literals, no for-of loops.
 */

function doGet(e) {
  var params = e ? (e.parameter || {}) : {};

  // ── Token check ────────────────────────────────────────────────────────────
  var expectedToken = PropertiesService.getScriptProperties().getProperty('SCAN_TOKEN');
  if (!expectedToken) {
    return jsonResponse({ error: 'SCAN_TOKEN not set in Script Properties' }, 500);
  }
  if (!params.token || params.token !== expectedToken) {
    return jsonResponse({ error: 'Unauthorized' }, 401);
  }

  // ── Tab param ──────────────────────────────────────────────────────────────
  var tabName = params.tab ? String(params.tab).trim() : '';
  if (!tabName) {
    return jsonResponse({ error: 'Missing ?tab= parameter' }, 400);
  }

  // ── Open spreadsheet — by ID param or fall back to bound sheet ────────────
  var sheetId = params.sheetId ? String(params.sheetId).trim() : '';
  var ss;
  try {
    ss = sheetId ? SpreadsheetApp.openById(sheetId) : SpreadsheetApp.getActiveSpreadsheet();
  } catch (openErr) {
    return jsonResponse({ error: 'Cannot open spreadsheet: ' + openErr.message }, 403);
  }

  // ── List tabs mode ─────────────────────────────────────────────────────────
  if (tabName === '__list__') {
    var sheets = ss.getSheets();
    var names = [];
    for (var i = 0; i < sheets.length; i++) {
      names.push(sheets[i].getName());
    }
    return jsonResponse({ tabs: names, spreadsheetId: ss.getId() });
  }

  // ── Read tab ───────────────────────────────────────────────────────────────
  try {
    var sheet = ss.getSheetByName(tabName);
    if (!sheet) {
      return jsonResponse({ error: 'Tab not found: ' + tabName }, 404);
    }

    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();

    if (lastRow === 0 || lastCol === 0) {
      return jsonResponse({ tab: tabName, rows: [] });
    }

    var values = sheet.getRange(1, 1, lastRow, lastCol).getValues();

    // Convert Date objects to ISO strings so JSON serializes cleanly
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

    return jsonResponse({ tab: tabName, rows: rows });

  } catch (err) {
    return jsonResponse({ error: err.message }, 500);
  }
}

function jsonResponse(obj, statusCode) {
  var output = ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
  return output;
}
