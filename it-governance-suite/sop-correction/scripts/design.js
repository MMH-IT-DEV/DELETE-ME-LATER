/**
 * design.js
 * Extracts full visual design from example SOPs using DocumentApp (no REST API needed).
 * Also tries HTML export via Drive API as a backup.
 *
 * Run extractDesignDocumentApp() first — writes full style spec to "DesignSpec" tab.
 * Run exportAsHTML() to save each example as an HTML file to Google Drive.
 */


// ─── PRIMARY: DocumentApp style extraction ────────────────────────────────────

function extractDesignDocumentApp() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('DesignSpec');
  if (!sheet) sheet = ss.insertSheet('DesignSpec');
  else { sheet.clearContents(); sheet.clearFormats(); }

  var rows = [['SOP', 'Element', 'Property', 'Value']];

  SOP_EXAMPLES.forEach(function(sop) {
    var doc  = DocumentApp.openById(sop.id);
    var body = doc.getBody();
    var id   = sop.sopId;
    Logger.log('Extracting design: ' + doc.getName());

    for (var i = 0; i < body.getNumChildren(); i++) {
      var child = body.getChild(i);
      var type  = child.getType();

      // ── PARAGRAPH ──
      if (type === DocumentApp.ElementType.PARAGRAPH) {
        var para  = child.asParagraph();
        var text  = para.getText().trim().substring(0, 50);
        var lbl   = 'PARA[' + i + '] heading=' + para.getHeading().toString();

        rows.push([id, lbl, 'text_preview', text]);
        rows.push([id, lbl, 'alignment',    _val(para.getAlignment())]);
        rows.push([id, lbl, 'spacingBefore', _pt(para.getSpacingBefore())]);
        rows.push([id, lbl, 'spacingAfter',  _pt(para.getSpacingAfter())]);
        rows.push([id, lbl, 'lineSpacing',   para.getLineSpacing() !== null ? para.getLineSpacing() + '%' : 'null']);
        rows.push([id, lbl, 'indentStart',   _pt(para.getIndentStart())]);
        rows.push([id, lbl, 'indentFirstLine', _pt(para.getIndentFirstLine())]);

        if (text.length > 0) {
          var t = para.editAsText();
          rows.push([id, lbl, 'fontFamily',  t.getFontFamily(0)   || 'inherited']);
          rows.push([id, lbl, 'fontSize',    t.getFontSize(0) !== null ? t.getFontSize(0) + 'pt' : 'inherited']);
          rows.push([id, lbl, 'textColor',   t.getForegroundColor(0) || 'inherited']);
          rows.push([id, lbl, 'bgColor',     t.getBackgroundColor(0) || 'none']);
          rows.push([id, lbl, 'bold',        t.isBold(0)]);
          rows.push([id, lbl, 'italic',      t.isItalic(0)]);
          rows.push([id, lbl, 'underline',   t.isUnderline(0)]);
        }

      // ── LIST ITEM ──
      } else if (type === DocumentApp.ElementType.LIST_ITEM) {
        var item = child.asListItem();
        var text = item.getText().trim().substring(0, 50);
        var lbl  = 'LIST[' + i + '] level=' + item.getNestingLevel();

        rows.push([id, lbl, 'text_preview', text]);
        rows.push([id, lbl, 'glyphType',    _val(item.getGlyphType())]);
        rows.push([id, lbl, 'alignment',    _val(item.getAlignment())]);
        rows.push([id, lbl, 'indentStart',  _pt(item.getIndentStart())]);

        if (text.length > 0) {
          var t = item.editAsText();
          rows.push([id, lbl, 'fontFamily', t.getFontFamily(0)   || 'inherited']);
          rows.push([id, lbl, 'fontSize',   t.getFontSize(0) !== null ? t.getFontSize(0) + 'pt' : 'inherited']);
          rows.push([id, lbl, 'textColor',  t.getForegroundColor(0) || 'inherited']);
          rows.push([id, lbl, 'bold',       t.isBold(0)]);
        }

      // ── TABLE ──
      } else if (type === DocumentApp.ElementType.TABLE) {
        var table = child.asTable();
        var numCols = table.getNumRows() > 0 ? table.getRow(0).getNumCells() : 0;
        var lbl   = 'TABLE[' + i + '] ' + table.getNumRows() + 'x' + numCols;

        rows.push([id, lbl, 'borderColor', table.getBorderColor() || 'none']);
        rows.push([id, lbl, 'borderWidth', table.getBorderWidth() !== null ? table.getBorderWidth() + 'pt' : 'null']);

        var maxRows = Math.min(table.getNumRows(), 3);
        for (var r = 0; r < maxRows; r++) {
          var numCells = table.getRow(r).getNumCells();
          for (var c = 0; c < numCells; c++) {
            var cell     = table.getCell(r, c);
            var cellText = cell.getText().trim().substring(0, 25);
            var clbl     = lbl + ' [r' + r + 'c' + c + '] "' + cellText + '"';

            rows.push([id, clbl, 'bgColor',        cell.getBackgroundColor() || 'none']);
            rows.push([id, clbl, 'verticalAlign',  _val(cell.getVerticalAlignment())]);
            rows.push([id, clbl, 'paddingTop',     _pt(cell.getPaddingTop())]);
            rows.push([id, clbl, 'paddingLeft',    _pt(cell.getPaddingLeft())]);

            if (cellText.length > 0) {
              var t = cell.editAsText();
              rows.push([id, clbl, 'fontFamily', t.getFontFamily(0)   || 'inherited']);
              rows.push([id, clbl, 'fontSize',   t.getFontSize(0) !== null ? t.getFontSize(0) + 'pt' : 'inherited']);
              rows.push([id, clbl, 'textColor',  t.getForegroundColor(0) || 'inherited']);
              rows.push([id, clbl, 'bgColor_txt',t.getBackgroundColor(0) || 'none']);
              rows.push([id, clbl, 'bold',       t.isBold(0)]);
            }
          }
        }
      }
    }

    rows.push([id, '────────', '────────', '──── END ' + id + ' ────']);
    Logger.log('✓ Done: ' + id);
  });

  sheet.getRange(1, 1, rows.length, 4).setValues(rows);
  sheet.getRange(1, 1, 1, 4).setFontWeight('bold').setBackground('#1a1a2e').setFontColor('#fff');
  sheet.autoResizeColumns(1, 3);
  sheet.setColumnWidth(4, 350);
  Logger.log('DesignSpec tab updated with ' + rows.length + ' rows.');
}


// ─── SUMMARIZE: Read DesignSpec and extract key findings ─────────────────────

function summarizeDesign() {
  var ss       = SpreadsheetApp.getActiveSpreadsheet();
  var specSheet = ss.getSheetByName('DesignSpec');
  if (!specSheet) { Logger.log('Run extractDesignDocumentApp first.'); return; }

  var data = specSheet.getDataRange().getValues();

  // Buckets
  var fonts   = {};
  var colors  = {};
  var sizes   = {};
  var bolds   = {};
  var tables  = {};
  var spacing = {};
  var aligns  = {};

  data.forEach(function(row) {
    var sop  = row[0];
    var elem = row[1].toString();
    var prop = row[2].toString();
    var val  = row[3].toString();
    if (!val || val === 'null' || val === 'inherited' || val === 'none' || val === '────────') return;

    var isHeading = elem.indexOf('heading=HEADING') !== -1;
    var headingType = isHeading ? elem.match(/heading=(\w+)/)[1] : 'NORMAL_TEXT';
    var isTable  = elem.indexOf('TABLE') !== -1;
    var isList   = elem.indexOf('LIST') !== -1;

    if (prop === 'fontFamily') {
      var key = (isTable ? 'TABLE' : isHeading ? headingType : isList ? 'LIST' : 'NORMAL') + ':' + val;
      fonts[key] = (fonts[key] || 0) + 1;
    }
    if (prop === 'fontSize') {
      var key = (isTable ? 'TABLE' : isHeading ? headingType : isList ? 'LIST' : 'NORMAL') + ':' + val;
      sizes[key] = (sizes[key] || 0) + 1;
    }
    if (prop === 'textColor') {
      var key = (isTable ? 'TABLE' : isHeading ? headingType : isList ? 'LIST' : 'NORMAL') + ':' + val;
      colors[key] = (colors[key] || 0) + 1;
    }
    if (prop === 'bold' && val === 'true') {
      var key = isTable ? 'TABLE' : isHeading ? headingType : isList ? 'LIST' : 'NORMAL';
      bolds[key] = (bolds[key] || 0) + 1;
    }
    if (prop === 'bgColor' && isTable) {
      tables[val] = (tables[val] || 0) + 1;
    }
    if (prop === 'borderColor') {
      tables['border:' + val] = (tables['border:' + val] || 0) + 1;
    }
    if (prop === 'spacingBefore' || prop === 'spacingAfter' || prop === 'lineSpacing') {
      var key = prop + ':' + val;
      spacing[key] = (spacing[key] || 0) + 1;
    }
    if (prop === 'alignment') {
      var key = (isHeading ? headingType : isTable ? 'TABLE' : 'NORMAL') + ':' + val;
      aligns[key] = (aligns[key] || 0) + 1;
    }
  });

  // Write summary sheet
  var sum = ss.getSheetByName('DesignSummary');
  if (!sum) sum = ss.insertSheet('DesignSummary');
  else sum.clearContents();

  var out = [['Category', 'Key', 'Count']];
  function _write(label, obj) {
    Object.keys(obj).sort().forEach(function(k) { out.push([label, k, obj[k]]); });
  }
  _write('FONT',    fonts);
  _write('SIZE',    sizes);
  _write('COLOR',   colors);
  _write('BOLD',    bolds);
  _write('TABLE_BG', tables);
  _write('SPACING', spacing);
  _write('ALIGN',   aligns);

  sum.getRange(1, 1, out.length, 3).setValues(out);
  sum.getRange(1, 1, 1, 3).setFontWeight('bold').setBackground('#1a1a2e').setFontColor('#fff');
  sum.autoResizeColumns(1, 3);

  Logger.log('Summary written to DesignSummary tab (' + out.length + ' rows). Paste that tab here.');
}


// ─── SECONDARY: HTML export via Drive API ─────────────────────────────────────

function exportAsHTML() {
  var token = ScriptApp.getOAuthToken();

  SOP_EXAMPLES.forEach(function(sop) {
    try {
      var url  = 'https://www.googleapis.com/drive/v3/files/' + sop.id + '/export?mimeType=text%2Fhtml';
      var resp = UrlFetchApp.fetch(url, {
        headers: { 'Authorization': 'Bearer ' + token },
        muteHttpExceptions: true
      });

      if (resp.getResponseCode() === 200) {
        var file = DriveApp.createFile(sop.sopId + '_design.html', resp.getContentText(), MimeType.HTML);
        Logger.log('✓ ' + sop.sopId + ' → ' + file.getUrl());
      } else {
        Logger.log('✗ ' + sop.sopId + ' HTTP ' + resp.getResponseCode() + ': ' + resp.getContentText().substring(0, 300));
      }
    } catch(e) {
      Logger.log('✗ ' + sop.sopId + ': ' + e.message);
    }
  });
}


// ─── HELPERS ─────────────────────────────────────────────────────────────────

function _pt(val) { return val !== null && val !== undefined ? val + 'pt' : 'null'; }
function _val(val) { return val !== null && val !== undefined ? val.toString() : 'null'; }
