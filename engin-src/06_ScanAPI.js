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

  var sheetId = params.sheetId ? String(params.sheetId).trim() : '';
  var ss;
  try {
    ss = sheetId ? SpreadsheetApp.openById(sheetId) : SpreadsheetApp.getActiveSpreadsheet();
  } catch (openErr) {
    return jsonScanResponse({ error: 'Cannot open spreadsheet: ' + openErr.message });
  }

  // ── List tabs mode ─────────────────────────────────────────────────────────
  if (tabName === '__list__') {
    var sheets = ss.getSheets();
    var names = [];
    var tabInfo = [];
    for (var i = 0; i < sheets.length; i++) {
      names.push(sheets[i].getName());
      tabInfo.push({
        name: sheets[i].getName(),
        sheetId: sheets[i].getSheetId(),
        rows: sheets[i].getLastRow(),
        columns: sheets[i].getLastColumn()
      });
    }
    return jsonScanResponse({
      tabs: names,
      tabInfo: tabInfo,
      spreadsheetId: ss.getId(),
      spreadsheetName: ss.getName()
    });
  }

  // ── Read tab ───────────────────────────────────────────────────────────────
  try {
    var sheet = ss.getSheetByName(tabName);
    if (!sheet) {
      return jsonScanResponse({ error: 'Tab not found: ' + tabName });
    }

    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();
    var startRow = parseScanPositiveInt_(params.startRow, 1);
    var startCol = parseScanPositiveInt_(params.startCol, 1);
    var maxRows = parseScanPositiveInt_(params.maxRows, lastRow || 1);
    var maxCols = parseScanPositiveInt_(params.maxCols, lastCol || 1);

    if (lastRow === 0 || lastCol === 0) {
      return jsonScanResponse({
        spreadsheetId: ss.getId(),
        spreadsheetName: ss.getName(),
        tab: tabName,
        totalRows: lastRow,
        totalColumns: lastCol,
        rowStart: startRow,
        columnStart: startCol,
        returnedRows: 0,
        returnedColumns: 0,
        truncated: false,
        rows: []
      });
    }

    if (startRow > lastRow || startCol > lastCol) {
      return jsonScanResponse({
        spreadsheetId: ss.getId(),
        spreadsheetName: ss.getName(),
        tab: tabName,
        totalRows: lastRow,
        totalColumns: lastCol,
        rowStart: startRow,
        columnStart: startCol,
        returnedRows: 0,
        returnedColumns: 0,
        truncated: true,
        rows: []
      });
    }

    var rowCount = Math.min(maxRows, (lastRow - startRow + 1));
    var colCount = Math.min(maxCols, (lastCol - startCol + 1));
    var values = sheet.getRange(startRow, startCol, rowCount, colCount).getValues();

    var rows = [];
    for (var r = 0; r < values.length; r++) {
      var row = [];
      for (var c = 0; c < values[r].length; c++) {
        row.push(serializeScanCell_(values[r][c]));
      }
      rows.push(row);
    }

    return jsonScanResponse({
      spreadsheetId: ss.getId(),
      spreadsheetName: ss.getName(),
      tab: tabName,
      sheetId: sheet.getSheetId(),
      totalRows: lastRow,
      totalColumns: lastCol,
      rowStart: startRow,
      rowEnd: startRow + rowCount - 1,
      columnStart: startCol,
      columnEnd: startCol + colCount - 1,
      returnedRows: rowCount,
      returnedColumns: colCount,
      truncated: startRow !== 1 || startCol !== 1 || rowCount !== lastRow || colCount !== lastCol,
      rows: rows
    });

  } catch (err) {
    return jsonScanResponse({ error: err.message });
  }
}

function jsonScanResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function parseScanPositiveInt_(rawValue, fallbackValue) {
  var parsed = parseInt(rawValue, 10);
  if (isNaN(parsed) || parsed < 1) {
    return fallbackValue;
  }
  return parsed;
}

function serializeScanCell_(cell) {
  if (cell instanceof Date) {
    return Utilities.formatDate(cell, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  }
  return cell;
}
