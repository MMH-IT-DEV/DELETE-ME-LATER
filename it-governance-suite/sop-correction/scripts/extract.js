/**
 * dumpToFixToLog
 * Prints the full structured content of IT-014 and IT-015 to the execution log.
 * Copy the log output and paste it to Claude for analysis.
 */
function dumpToFixToLog() {
  _dumpDocsToLog(SOP_TO_FIX);
}


/**
 * _dumpDocsToLog (internal)
 */
function _dumpDocsToLog(docList) {
  docList.forEach(function(sop) {
    var doc  = DocumentApp.openById(sop.id);
    var body = doc.getBody();
    var lines = ['', '════════════════════════════════════════', 'DOC: ' + doc.getName(), 'ID:  ' + sop.sopId, '════════════════════════════════════════'];

    for (var i = 0; i < body.getNumChildren(); i++) {
      var child = body.getChild(i);
      var type  = child.getType();

      if (type === DocumentApp.ElementType.PARAGRAPH) {
        var para  = child.asParagraph();
        var style = para.getHeading().toString();
        var text  = para.getText().trim();
        if (!text) continue;
        var prefix = style === 'NORMAL' ? '      ' : '[' + style + '] ';
        lines.push(prefix + text);

      } else if (type === DocumentApp.ElementType.LIST_ITEM) {
        var item  = child.asListItem();
        var level = item.getNestingLevel();
        var text  = item.getText().trim();
        if (text) lines.push(Array(level * 2 + 3).join(' ') + '• ' + text);

      } else if (type === DocumentApp.ElementType.TABLE) {
        var table = child.asTable();
        lines.push('  [TABLE]');
        for (var r = 0; r < table.getNumRows(); r++) {
          var tableRow = table.getRow(r);
          var cells = [];
          for (var c = 0; c < tableRow.getNumCells(); c++) {
            cells.push(tableRow.getCell(c).getText().trim());
          }
          lines.push('  | ' + cells.join(' | ') + ' |');
        }
      }
    }
    Logger.log(lines.join('\n'));
  });
}


/**
 * dumpExamplesToLog
 * Prints the full structured content of both example SOPs to the execution log.
 * Copy the log output and paste it to Claude for analysis.
 */
function dumpExamplesToLog() {
  _dumpDocsToLog(SOP_EXAMPLES);
}


/**
 * scanAndCompare
 * Reads all 4 focus docs (2 examples + 2 to-fix) and writes their full
 * structured content into separate tabs for side-by-side comparison.
 * Run this FIRST so we can analyze the example format vs current state.
 */
function scanAndCompare() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  _extractDocsToSheet(ss, SOP_EXAMPLES, 'Examples', '#0d6e2e', '#ffffff');
  _extractDocsToSheet(ss, SOP_TO_FIX,   'ToFix',    '#7f1d1d', '#ffffff');

  Logger.log('Scan complete! "Examples" tab → LOG-006, QA-012 | "ToFix" tab → IT-014, IT-015');
}


/**
 * _extractDocsToSheet (internal)
 * Extracts full paragraph/list/table content from a list of docs
 * and writes them to a named sheet tab.
 */
function _extractDocsToSheet(ss, docList, sheetName, headerBg, headerFg) {
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  } else {
    sheet.clearContents();
    sheet.clearFormats();
  }

  var headers = ['SOP ID', 'Doc Title', 'Style', 'Content'];
  sheet.getRange(1, 1, 1, 4)
    .setValues([headers])
    .setFontWeight('bold')
    .setBackground(headerBg)
    .setFontColor(headerFg);

  var rows = [];

  docList.forEach(function(sop) {
    try {
      var doc   = DocumentApp.openById(sop.id);
      var title = doc.getName();
      var body  = doc.getBody();
      Logger.log('Scanning: ' + title);

      for (var i = 0; i < body.getNumChildren(); i++) {
        var child = body.getChild(i);
        var type  = child.getType();

        if (type === DocumentApp.ElementType.PARAGRAPH) {
          var para  = child.asParagraph();
          var style = para.getHeading().toString();
          var text  = para.getText().trim();
          if (text) rows.push([sop.sopId, title, style, text]);

        } else if (type === DocumentApp.ElementType.LIST_ITEM) {
          var item  = child.asListItem();
          var level = item.getNestingLevel();
          var text  = item.getText().trim();
          if (text) {
            var indent = Array(level + 1).join('    ');
            rows.push([sop.sopId, title, 'LIST_L' + level, indent + '• ' + text]);
          }

        } else if (type === DocumentApp.ElementType.TABLE) {
          var table = child.asTable();
          for (var r = 0; r < table.getNumRows(); r++) {
            var tableRow = table.getRow(r);
            var cells = [];
            for (var c = 0; c < tableRow.getNumCells(); c++) {
              cells.push(tableRow.getCell(c).getText().trim());
            }
            rows.push([sop.sopId, title, 'TABLE_ROW_' + r, cells.join(' | ')]);
          }
        }
      }

      rows.push(['', '', '────────', '──── END: ' + sop.sopId + ' — ' + title + ' ────']);

    } catch (e) {
      Logger.log('ERROR ' + sop.sopId + ': ' + e.message);
      rows.push([sop.sopId, 'ERROR', 'ERROR', e.message]);
    }
  });

  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, 4).setValues(rows);
  }
  sheet.autoResizeColumns(1, 4);
  sheet.setColumnWidth(4, 600);
}


/**
 * extractAllSOPs
 * Bulk extractor — reads all 15 IT SOPs into the "Extraction" tab.
 */
function extractAllSOPs() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Extraction');
  if (!sheet) sheet = ss.insertSheet('Extraction');
  else sheet.clearContents();

  var headers = ['SOP ID', 'Doc Title', 'Doc URL', 'Element Type', 'Heading Style', 'Content', 'Extracted At'];
  sheet.getRange(1, 1, 1, headers.length)
    .setValues([headers])
    .setFontWeight('bold')
    .setBackground('#1a1a2e')
    .setFontColor('#ffffff');

  var rows = [];
  var timestamp = new Date().toISOString();
  var ok = 0, err = 0;

  SOP_DOCS.forEach(function(sop) {
    try {
      var doc   = DocumentApp.openById(sop.id);
      var title = doc.getName();
      var url   = doc.getUrl();
      var body  = doc.getBody();

      for (var i = 0; i < body.getNumChildren(); i++) {
        var child = body.getChild(i);
        var type  = child.getType();

        if (type === DocumentApp.ElementType.PARAGRAPH) {
          var para  = child.asParagraph();
          var style = para.getHeading().toString();
          var text  = para.getText().trim();
          if (text) rows.push([sop.sopId, title, url, 'PARAGRAPH', style, text, timestamp]);

        } else if (type === DocumentApp.ElementType.LIST_ITEM) {
          var item  = child.asListItem();
          var level = item.getNestingLevel();
          var text  = item.getText().trim();
          if (text) rows.push([sop.sopId, title, url, 'LIST_ITEM', 'LEVEL_' + level, Array(level+1).join('  ') + '• ' + text, timestamp]);

        } else if (type === DocumentApp.ElementType.TABLE) {
          var table = child.asTable();
          for (var r = 0; r < table.getNumRows(); r++) {
            var tableRow = table.getRow(r);
            var cells = [];
            for (var c = 0; c < tableRow.getNumCells(); c++) {
              cells.push(tableRow.getCell(c).getText().trim());
            }
            rows.push([sop.sopId, title, url, 'TABLE_ROW', 'ROW_' + r, cells.join(' | '), timestamp]);
          }
        }
      }
      rows.push(['', '', '', '─────', '─────', '──── END: ' + sop.sopId + ' ────', '']);
      ok++;
    } catch (e) {
      Logger.log('✗ ' + sop.sopId + ': ' + e.message);
      rows.push([sop.sopId, 'ERROR', '', 'ERROR', 'ERROR', e.message, timestamp]);
      err++;
    }
  });

  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  }
  sheet.autoResizeColumns(1, headers.length);
  Logger.log('Done. ✓ ' + ok + ' extracted, ✗ ' + err + ' errors. ' + rows.length + ' rows total.');
}


/**
 * listSOPTitles — quick access check for all 15 IT SOPs
 */
function listSOPTitles() {
  var results = SOP_DOCS.map(function(sop) {
    try {
      return sop.sopId + ' | ' + DocumentApp.openById(sop.id).getName();
    } catch (e) {
      return sop.sopId + ' | ERROR: ' + e.message;
    }
  });
  Logger.log(results.join('\n'));
  Logger.log('SOP Title Check:\n' + results.join('\n'));
}
