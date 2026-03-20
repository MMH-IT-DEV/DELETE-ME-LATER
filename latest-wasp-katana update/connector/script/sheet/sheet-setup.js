var GMP_SHEET_COLORS = {
  TITLE: '#0f172a',
  HEADER: '#163b65',
  REQUEST_SECTION: '#8a5a00',
  ACCOUNT_SECTION: '#0f5fa8',
  REVIEW_SECTION: '#0f5fa8',
  INCIDENT_SECTION: '#9a3412',
  TEXT_LIGHT: '#ffffff',
  NOTE_BG: '#f8fafc',
  BORDER: '#d0d7de'
};

var LEGACY_ACCOUNT_HEADERS = [
  'Access ID',
  'System',
  'Person',
  'Company Email',
  'Platform Account',
  'Access Level',
  'Access Status',
  'MFA',
  'Privileged',
  'Last Verified In System',
  'Next Review Due',
  'Review Status',
  'Source Request ID',
  'Audit Log'
];

var LEGACY_REVIEW_HEADERS = [
  'Review ID',
  'Priority',
  'Person',
  'System',
  'Access Level',
  'Last Login',
  '30d Logins',
  'Why Flagged',
  'Reviewer Action',
  'Decision Status',
  'Next Review Due',
  'Notes',
  'Trigger Type',
  'Access ID',
  'Days Since Login',
  'Activity Score',
  'Decision Date',
  'Linked Request',
  'Slack Thread',
  'Next Slack Trigger'
];

var LEGACY_INCIDENT_HEADERS = [
  'Incident ID',
  'Reported At',
  'Severity',
  'System',
  'Reported By',
  'Issue Type',
  'Summary',
  'Linked Access ID',
  'Linked Request ID',
  'Assigned To',
  'Status',
  'Next Slack Trigger',
  'Response Target',
  'Root Cause',
  'Resolution',
  'Resolved At',
  'Slack Thread',
  'Notes'
];

var HEADER_ALIASES = {
  'Access Status': ['Presence Status', 'Current State'],
  'Last Verified In System': ['Granted On'],
  'Activity Status': ['Activity Band'],
  'Last Activity Evidence': ['Last Login'],
  'Review Cycle': ['Next Review Due'],
  'Next Check Date': ['Check Date', 'Next Review Due'],
  'Check Needed': ['Why In Review', 'Why Flagged'],
  'Review Status': ['Decision Status'],
  'Reviewed At': ['Completed At', 'Decision Date'],
  'IT Owner': ['Provisioning Owner', 'Owner'],
  'Audit Log': ['Notes']
};

function forceRebuildAccountSection_() {
  var sheet = SpreadsheetApp.getActive().getSheetByName(GMP_SCHEMA.tabs.ACCESS_CONTROL);
  if (!sheet) { return; }

  var ctx;
  try { ctx = getAccessControlContext_(); } catch (err) { return; }

  var startRow = ctx.accountHeaderRow + 1;
  var endRow = ctx.archiveSectionRow === -1 ? sheet.getLastRow() : ctx.archiveSectionRow - 1;
  if (endRow < startRow) { return; }

  // Known valid Access Status values — used to detect which column layout the DATA is in
  var accessStatusValues = { 'Setting Up': 1, 'Access Granted': 1, 'Present': 1, Unmanaged: 1, Missing: 1, Revoked: 1 };
  var accessLevelValues = { Admin: 1, 'Full Access': 1, 'Read-Write': 1, 'Read-Only': 1, Limited: 1 };

  // All possible header layouts the data might be in (try each until one validates)
  var knownLayouts = [
    // Current new 10-column layout
    ['Access ID', 'System', 'Person', 'Company Email', 'Platform Account', 'Access Level', 'Access Status', 'MFA', 'Last Verified', 'Next Review Due'],
    // Previous 14-column layout
    ['Access ID', 'System', 'Person', 'Company Email', 'Platform Account', 'Access Level', 'Access Status', 'MFA', 'Privileged', 'Last Verified In System', 'Next Review Due', 'Review Status', 'Source Request ID', 'Audit Log'],
    // Previous 20-column layout
    ['Access ID', 'System', 'Person', 'Company Email', 'Dept', 'Platform Account', 'Access Level', 'Access Status', 'Last Verified In System', 'Next Review Due', 'Last Review Date', 'Review Status', 'IT Owner', 'Audit Log', 'Source Request ID', 'Granted On', 'MFA', 'Privileged', 'Review Trigger', 'Slack Thread'],
    // Original 31-column layout
    ['Access ID', 'System', 'Person', 'Company Email', 'Dept', 'Platform Account', 'Access Level', 'Account Status', 'Last Verified In System', 'Activity Status', 'Last Activity Evidence', 'Next Review Due', 'Last Review Date', 'Review Status', 'IT Owner', 'Audit Log', 'Source Request ID', 'Granted On', 'Days Since Login', '30d Logins', '90d Logins', 'MFA', 'Privileged', 'Activity Score', 'Review Trigger', 'Expected Use', 'Min 30d Logins', 'Stale After (Days)', 'Activity Source', 'Activity Confidence', 'Slack Thread']
  ];

  var dataValues = sheet.getRange(startRow, 1, endRow - startRow + 1, sheet.getMaxColumns()).getValues();

  // Find the first non-blank, non-group data row to test layouts against
  var testRow = null;
  for (var t = 0; t < dataValues.length; t++) {
    if (!isBlankRow_(dataValues[t]) && !isDisplayGroupRow_(dataValues[t])) {
      testRow = dataValues[t];
      break;
    }
  }
  if (!testRow) { return; }

  // Try each layout and pick the one where Access Level and Access Status columns have valid values
  var bestLayout = knownLayouts[0];
  for (var l = 0; l < knownLayouts.length; l++) {
    var layout = knownLayouts[l];
    var alIdx = -1, asIdx = -1;
    for (var c = 0; c < layout.length; c++) {
      if (layout[c] === 'Access Level') { alIdx = c; }
      if (layout[c] === 'Access Status' || layout[c] === 'Account Status') { asIdx = c; }
    }
    if (alIdx >= 0 && asIdx >= 0 && alIdx < testRow.length && asIdx < testRow.length) {
      var alVal = normalizeText_(testRow[alIdx]);
      var asVal = normalizeText_(testRow[asIdx]);
      if (accessLevelValues[alVal] && accessStatusValues[asVal]) {
        bestLayout = layout;
        break;
      }
    }
  }

  // Build header map from the detected layout
  var headerMap = {};
  for (var h = 0; h < bestLayout.length; h++) {
    headerMap[bestLayout[h]] = h;
  }

  // Aliases: new header name → old header names to try
  var nameAliases = {
    'Access Status': ['Account Status', 'Current State', 'Presence Status'],
    'Last Verified': ['Last Verified In System', 'Granted On'],
    'MFA': ['MFA']
  };

  var newHeaders = GMP_SCHEMA.accounts.headers;
  var rebuilt = [];
  for (var i = 0; i < dataValues.length; i++) {
    var oldRow = dataValues[i];
    if (isBlankRow_(oldRow) || isDisplayGroupRow_(oldRow)) {
      rebuilt.push(buildBlankRow_(newHeaders.length));
      continue;
    }
    var newRow = [];
    for (var j = 0; j < newHeaders.length; j++) {
      var target = newHeaders[j];
      var value = '';
      if (headerMap.hasOwnProperty(target)) {
        value = oldRow[headerMap[target]];
      }
      if (!value && nameAliases[target]) {
        for (var a = 0; a < nameAliases[target].length; a++) {
          if (headerMap.hasOwnProperty(nameAliases[target][a])) {
            value = oldRow[headerMap[nameAliases[target][a]]];
            if (value) { break; }
          }
        }
      }
      // Normalize Access Status values
      if (target === 'Access Status' && normalizeText_(value) === 'Present') {
        value = 'Access Granted';
      }
      newRow.push(value !== null && value !== undefined ? value : '');
    }
    rebuilt.push(newRow);
  }

  // Write new headers + rebuilt data, clear all excess
  var fullWidth = sheet.getMaxColumns();
  var newHeaderRow = newHeaders.slice();
  while (newHeaderRow.length < fullWidth) { newHeaderRow.push(''); }
  sheet.getRange(ctx.accountHeaderRow, 1, 1, fullWidth).setValues([newHeaderRow]);

  for (var k = 0; k < rebuilt.length; k++) {
    while (rebuilt[k].length < fullWidth) { rebuilt[k].push(''); }
  }
  if (rebuilt.length > 0) {
    sheet.getRange(startRow, 1, rebuilt.length, fullWidth).setValues(rebuilt);
  }
}

function refreshSheet() {
  forceRebuildAccountSection_();
  ensureThreeTabWorkbookReady_();

  sortRequestSection_();
  sortArchivedRequestSection_();
  sortAccountSection_();
  sortReviewSection_();
  sortIncidentSection_();

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var accessSheet = ss.getSheetByName(GMP_SCHEMA.tabs.ACCESS_CONTROL);
  var reviewSheet = ss.getSheetByName(GMP_SCHEMA.tabs.REVIEW_CHECKS);
  var incidentSheet = ss.getSheetByName(GMP_SCHEMA.tabs.INCIDENTS);

  if (accessSheet) { formatAccessControlSheet_(accessSheet); }
  if (reviewSheet) { formatReviewChecksSheet_(reviewSheet); }
  if (incidentSheet) { formatIncidentsSheet_(incidentSheet); }

  SpreadsheetApp.flush();

  try {
    SpreadsheetApp.getUi().alert('Refresh Complete', 'All tabs sorted and formatted.', SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (err) {
    Logger.log('Sheet refresh complete');
  }
}

function setupSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var archivedSheets = [];

  setupAccessControlSheet_(ss, archivedSheets);
  setupReviewChecksSheet_(ss, archivedSheets);
  setupIncidentsSheet_(ss, archivedSheets);
  sortRequestSection_();
  sortAccountSection_();
  sortReviewSection_();
  sortIncidentSection_();
  normalizeStoredSlackPeopleLabels_();

  setWorkbookSchemaVersion_();
  SpreadsheetApp.flush();

  var message = 'Three-tab system workbook is ready.\n\n' +
    'Sheets:\n' +
    '- ' + GMP_SCHEMA.tabs.ACCESS_CONTROL + '\n' +
    '- ' + GMP_SCHEMA.tabs.REVIEW_CHECKS + '\n' +
    '- ' + GMP_SCHEMA.tabs.INCIDENTS;

  if (archivedSheets.length > 0) {
    message += '\n\nArchived legacy tabs:\n- ' + archivedSheets.join('\n- ');
  }

  try {
    SpreadsheetApp.getUi().alert('Setup Complete', message, SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (err) {
    Logger.log(message);
  }
}

function setupAccessControlSheet_(ss, archivedSheets) {
  var sheet = ss.getSheetByName(GMP_SCHEMA.tabs.ACCESS_CONTROL);
  if (!sheet) {
    sheet = ss.insertSheet(GMP_SCHEMA.tabs.ACCESS_CONTROL);
  }

  if (!accessControlMatchesSchema_(sheet) && !migrateLegacyAccessControlAccountSection_(sheet)) {
    if (sheet.getLastRow() > 0 || sheet.getLastColumn() > 0) {
      archivedSheets.push(archiveSheet_(sheet));
    }
    resetSheet_(sheet);
    buildAccessControlSkeleton_(sheet);
  }
  formatAccessControlSheet_(sheet);
}

function setupReviewChecksSheet_(ss, archivedSheets) {
  var sheet = ss.getSheetByName(GMP_SCHEMA.tabs.REVIEW_CHECKS);
  if (!sheet) {
    sheet = ss.insertSheet(GMP_SCHEMA.tabs.REVIEW_CHECKS);
  }

  if (!sheetHeaderMatches_(sheet, GMP_SCHEMA.layout.REVIEW_HEADER_ROW, GMP_SCHEMA.reviews.headers) &&
      !migrateLegacyReviewChecksSheet_(sheet)) {
    if (sheet.getLastRow() > 0 || sheet.getLastColumn() > 0) {
      archivedSheets.push(archiveSheet_(sheet));
    }
    resetSheet_(sheet);
    buildReviewChecksSkeleton_(sheet);
  }
  formatReviewChecksSheet_(sheet);
}

function setupIncidentsSheet_(ss, archivedSheets) {
  var sheet = ss.getSheetByName(GMP_SCHEMA.tabs.INCIDENTS);
  if (!sheet) {
    sheet = ss.insertSheet(GMP_SCHEMA.tabs.INCIDENTS);
  }

  if (!sheetHeaderMatches_(sheet, GMP_SCHEMA.layout.INCIDENT_HEADER_ROW, GMP_SCHEMA.incidents.headers) &&
      !migrateLegacyIncidentsSheet_(sheet)) {
    if (sheet.getLastRow() > 0 || sheet.getLastColumn() > 0) {
      archivedSheets.push(archiveSheet_(sheet));
    }
    resetSheet_(sheet);
    buildIncidentsSkeleton_(sheet);
  }
  formatIncidentsSheet_(sheet);
}

function prepareTargetSheet_(ss, sheetName, archivedSheets, matchesSchemaFn) {
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    return ss.insertSheet(sheetName);
  }

  if (matchesSchemaFn(sheet)) {
    return sheet;
  }

  if (sheet.getLastRow() > 0 || sheet.getLastColumn() > 0) {
    archivedSheets.push(archiveSheet_(sheet));
  }

  resetSheet_(sheet);
  return sheet;
}

function accessControlMatchesSchema_(sheet) {
  var ctx;
  try {
    ctx = getAccessControlContext_();
  } catch (err) {
    return false;
  }
  return sheetHeaderMatches_(sheet, ctx.requestHeaderRow, GMP_SCHEMA.requests.headers) &&
    sheetHeaderMatches_(sheet, ctx.accountHeaderRow, GMP_SCHEMA.accounts.headers);
}

function archiveSheet_(sheet) {
  var ss = sheet.getParent();
  var suffix = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd-HHmmss');
  var archivedName = sheet.getName() + ' Legacy ' + suffix;
  var copy = sheet.copyTo(ss);
  copy.setName(archivedName);
  return archivedName;
}

function resetSheet_(sheet) {
  var dataRange = sheet.getDataRange();
  if (dataRange.getNumRows() > 0 && dataRange.getNumColumns() > 0) {
    dataRange.breakApart();
  }
  sheet.clear();
  sheet.setConditionalFormatRules([]);
  sheet.setFrozenRows(0);
}

function buildAccessControlSkeleton_(sheet) {
  var width = Math.max(GMP_SCHEMA.requests.headers.length, GMP_SCHEMA.accounts.headers.length);
  writeTitleRows_(
    sheet,
    width,
    'System Access Control',
    'Approve open requests at the top. Verify active accounts by system in the middle. Use archived requests at the bottom for history only.'
  );
  writeSectionRow_(sheet, GMP_SCHEMA.layout.REQUEST_SECTION_ROW, GMP_SCHEMA.sectionLabels.REQUESTS, width, GMP_SHEET_COLORS.REQUEST_SECTION);
  writeHeaderRow_(sheet, GMP_SCHEMA.layout.REQUEST_HEADER_ROW, GMP_SCHEMA.requests.headers, GMP_SHEET_COLORS.HEADER);
  writeSectionRow_(sheet, GMP_SCHEMA.layout.ACCOUNT_SECTION_ROW, GMP_SCHEMA.sectionLabels.ACCOUNTS, width, GMP_SHEET_COLORS.ACCOUNT_SECTION);
  writeHeaderRow_(sheet, GMP_SCHEMA.layout.ACCOUNT_HEADER_ROW, GMP_SCHEMA.accounts.headers, GMP_SHEET_COLORS.HEADER);
  writeSectionRow_(sheet, GMP_SCHEMA.layout.ARCHIVE_SECTION_ROW, GMP_SCHEMA.sectionLabels.ARCHIVED_REQUESTS, width, GMP_SHEET_COLORS.REQUEST_SECTION);
  writeHeaderRow_(sheet, GMP_SCHEMA.layout.ARCHIVE_HEADER_ROW, GMP_SCHEMA.requests.headers, GMP_SHEET_COLORS.HEADER);
}

function buildReviewChecksSkeleton_(sheet) {
  writeTitleRows_(
    sheet,
    GMP_SCHEMA.reviews.headers.length,
    'Review Checks',
    'Use the orange cell directly under Next Check Date to schedule the whole open queue. Change a row date only when one account needs a different follow-up day. Completed reviews are history only.'
  );
  writeSectionRow_(
    sheet,
    GMP_SCHEMA.layout.REVIEW_SECTION_ROW,
    GMP_SCHEMA.sectionLabels.REVIEWS,
    GMP_SCHEMA.reviews.headers.length,
    GMP_SHEET_COLORS.REVIEW_SECTION
  );
  writeHeaderRow_(sheet, GMP_SCHEMA.layout.REVIEW_HEADER_ROW, GMP_SCHEMA.reviews.headers, GMP_SHEET_COLORS.HEADER);
  writeReviewBulkControlRow_(sheet);
}

function buildIncidentsSkeleton_(sheet) {
  writeTitleRows_(
    sheet,
    GMP_SCHEMA.incidents.headers.length,
    'Incidents',
    'Use this tab for /system-issue and linked system access follow-up.'
  );
  writeSectionRow_(
    sheet,
    GMP_SCHEMA.layout.INCIDENT_SECTION_ROW,
    GMP_SCHEMA.sectionLabels.INCIDENTS,
    GMP_SCHEMA.incidents.headers.length,
    GMP_SHEET_COLORS.INCIDENT_SECTION
  );
  writeHeaderRow_(sheet, GMP_SCHEMA.layout.INCIDENT_HEADER_ROW, GMP_SCHEMA.incidents.headers, GMP_SHEET_COLORS.HEADER);
}

function formatAccessControlSheet_(sheet) {
  var ctx = ensureArchivedRequestsSection_();
  var width = Math.max(GMP_SCHEMA.requests.headers.length, GMP_SCHEMA.accounts.headers.length);

  writeTitleRows_(
    sheet,
    width,
    'System Access Control',
    'Approve open requests at the top. Use active accounts in the middle to verify who has access by system. Archived requests at the bottom are history only.'
  );
  writeSectionRow_(sheet, ctx.requestSectionRow, GMP_SCHEMA.sectionLabels.REQUESTS, width, GMP_SHEET_COLORS.REQUEST_SECTION);
  writeHeaderRow_(sheet, ctx.requestHeaderRow, GMP_SCHEMA.requests.headers, GMP_SHEET_COLORS.HEADER);
  writeSectionRow_(sheet, ctx.accountSectionRow, GMP_SCHEMA.sectionLabels.ACCOUNTS, width, GMP_SHEET_COLORS.ACCOUNT_SECTION);
  writeHeaderRow_(sheet, ctx.accountHeaderRow, GMP_SCHEMA.accounts.headers, GMP_SHEET_COLORS.HEADER);
  writeSectionRow_(sheet, ctx.archiveSectionRow, GMP_SCHEMA.sectionLabels.ARCHIVED_REQUESTS, width, GMP_SHEET_COLORS.REQUEST_SECTION);
  writeHeaderRow_(sheet, ctx.archiveHeaderRow, GMP_SCHEMA.requests.headers, GMP_SHEET_COLORS.HEADER);

  sheet.setFrozenRows(GMP_SCHEMA.layout.REQUEST_HEADER_ROW);
  var requestRowCount = Math.max(ctx.accountSectionRow - ctx.requestHeaderRow - 1, 1);
  var accountRowCount = Math.max(ctx.archiveSectionRow - ctx.accountHeaderRow - 1, 1);
  var archiveRowCount = Math.max(sheet.getMaxRows() - ctx.archiveHeaderRow, 1);
  applyDropdowns_(sheet, ctx.requestHeaderRow + 1, requestRowCount, GMP_SCHEMA.requests);
  applyDropdowns_(sheet, ctx.accountHeaderRow + 1, accountRowCount, GMP_SCHEMA.accounts);
  applyDropdowns_(sheet, ctx.archiveHeaderRow + 1, archiveRowCount, GMP_SCHEMA.requests);
  applyDateFormatsByKeys_(sheet, GMP_SCHEMA.requests.columns, ['SUBMITTED_AT', 'SLA_DUE', 'DECISION_DATE'], ctx.requestHeaderRow + 1, requestRowCount, 'yyyy-mm-dd hh:mm');
  applyDateFormatsByKeys_(sheet, GMP_SCHEMA.accounts.columns, ['LAST_VERIFIED_AT', 'LAST_REVIEW_DATE'], ctx.accountHeaderRow + 1, accountRowCount, 'yyyy-mm-dd hh:mm');
  applyDateFormatsByKeys_(sheet, GMP_SCHEMA.accounts.columns, ['NEXT_REVIEW_DUE'], ctx.accountHeaderRow + 1, accountRowCount, 'yyyy-mm-dd');
  applyDateFormatsByKeys_(sheet, GMP_SCHEMA.requests.columns, ['SUBMITTED_AT', 'SLA_DUE', 'DECISION_DATE'], ctx.archiveHeaderRow + 1, archiveRowCount, 'yyyy-mm-dd hh:mm');
  setColumnWidths_(sheet, GMP_SCHEMA.requests.widths, GMP_SCHEMA.accounts.widths);
  applyStandardRowHeights_(sheet, [
    GMP_SCHEMA.layout.TITLE_ROW,
    GMP_SCHEMA.layout.NOTES_ROW,
    GMP_SCHEMA.layout.REQUEST_SECTION_ROW,
    ctx.requestHeaderRow,
    ctx.accountSectionRow,
    ctx.accountHeaderRow,
    ctx.archiveSectionRow,
    ctx.archiveHeaderRow
  ]);
  applyDataRowFormatting_(sheet, ctx.requestHeaderRow + 1, ctx.accountSectionRow - 1, GMP_SCHEMA.requests.keys.length);
  applyDataRowFormatting_(sheet, ctx.accountHeaderRow + 1, Math.max(ctx.archiveSectionRow - 1, ctx.accountHeaderRow + 1), GMP_SCHEMA.accounts.keys.length);
  applyDataRowFormatting_(sheet, ctx.archiveHeaderRow + 1, Math.max(sheet.getLastRow(), ctx.archiveHeaderRow + 1), GMP_SCHEMA.requests.keys.length);
  setAccessControlColumnVisibility_(sheet);
  applyAccessControlConditionalFormatting_(sheet, ctx);
}

function formatReviewChecksSheet_(sheet) {
  writeTitleRows_(
    sheet,
    GMP_SCHEMA.reviews.headers.length,
    'Review Checks',
    'Use the orange cell directly under Next Check Date to schedule all open reviews. Change a row date only when one account needs a different follow-up day. Scheduled follow-ups stay separate from the open queue.'
  );
  writeSectionRow_(
    sheet,
    GMP_SCHEMA.layout.REVIEW_SECTION_ROW,
    GMP_SCHEMA.sectionLabels.REVIEWS,
    GMP_SCHEMA.reviews.headers.length,
    GMP_SHEET_COLORS.REVIEW_SECTION
  );
  writeHeaderRow_(sheet, GMP_SCHEMA.layout.REVIEW_HEADER_ROW, GMP_SCHEMA.reviews.headers, GMP_SHEET_COLORS.HEADER);
  writeReviewBulkControlRow_(sheet);

  sheet.setFrozenRows(GMP_SCHEMA.layout.REVIEW_CONTROL_ROW);
  var reviewDataStartRow = GMP_SCHEMA.layout.REVIEW_CONTROL_ROW + 1;
  applyDropdowns_(sheet, reviewDataStartRow, 2000, GMP_SCHEMA.reviews);
  applyDateFormatsByKeys_(sheet, GMP_SCHEMA.reviews.columns, ['LAST_LOGIN', 'DECISION_DATE'], reviewDataStartRow, 2000, 'yyyy-mm-dd hh:mm');
  applyDateFormatsByKeys_(sheet, GMP_SCHEMA.reviews.columns, ['NEXT_REVIEW_DUE'], reviewDataStartRow, 2000, 'yyyy-mm-dd');
  applyDateValidation_(sheet.getRange(GMP_SCHEMA.layout.REVIEW_CONTROL_ROW, GMP_SCHEMA.reviews.columns.NEXT_REVIEW_DUE, 1, 1));
  applyDateValidation_(sheet.getRange(reviewDataStartRow, GMP_SCHEMA.reviews.columns.NEXT_REVIEW_DUE, 2000, 1));
  setColumnWidths_(sheet, GMP_SCHEMA.reviews.widths);
  applyStandardRowHeights_(sheet, [
    GMP_SCHEMA.layout.TITLE_ROW,
    GMP_SCHEMA.layout.REVIEW_SECTION_ROW,
    GMP_SCHEMA.layout.REVIEW_CONTROL_ROW,
    GMP_SCHEMA.layout.REVIEW_HEADER_ROW
  ]);
  applyDataRowFormatting_(sheet, reviewDataStartRow, Math.max(sheet.getLastRow(), reviewDataStartRow), GMP_SCHEMA.reviews.keys.length);
  setReviewChecksColumnVisibility_(sheet);
  applyReviewConditionalFormatting_(sheet);
  highlightReviewCheckDateColumn_(sheet);
  formatReviewGroupRows_(sheet, reviewDataStartRow, Math.max(sheet.getLastRow() - GMP_SCHEMA.layout.REVIEW_CONTROL_ROW, 0));
}

function formatIncidentsSheet_(sheet) {
  writeTitleRows_(
    sheet,
    GMP_SCHEMA.incidents.headers.length,
    'Incidents',
    'Report system issues from Slack and keep linked status updates in the same thread.'
  );
  writeSectionRow_(
    sheet,
    GMP_SCHEMA.layout.INCIDENT_SECTION_ROW,
    GMP_SCHEMA.sectionLabels.INCIDENTS,
    GMP_SCHEMA.incidents.headers.length,
    GMP_SHEET_COLORS.INCIDENT_SECTION
  );
  writeHeaderRow_(sheet, GMP_SCHEMA.layout.INCIDENT_HEADER_ROW, GMP_SCHEMA.incidents.headers, GMP_SHEET_COLORS.HEADER);

  sheet.setFrozenRows(GMP_SCHEMA.layout.INCIDENT_HEADER_ROW);
  applyDropdowns_(sheet, GMP_SCHEMA.layout.INCIDENT_HEADER_ROW + 1, 2000, GMP_SCHEMA.incidents);
  applyDateFormatsByKeys_(sheet, GMP_SCHEMA.incidents.columns, ['REPORTED_AT', 'RESOLVED_AT'], GMP_SCHEMA.layout.INCIDENT_HEADER_ROW + 1, 2000, 'yyyy-mm-dd hh:mm');
  setColumnWidths_(sheet, GMP_SCHEMA.incidents.widths);
  applyStandardRowHeights_(sheet, [
    GMP_SCHEMA.layout.TITLE_ROW,
    GMP_SCHEMA.layout.INCIDENT_SECTION_ROW,
    GMP_SCHEMA.layout.INCIDENT_HEADER_ROW
  ]);
  applyDataRowFormatting_(sheet, GMP_SCHEMA.layout.INCIDENT_HEADER_ROW + 1, Math.max(sheet.getLastRow(), GMP_SCHEMA.layout.INCIDENT_HEADER_ROW + 1), GMP_SCHEMA.incidents.keys.length);
  if (sheet.getLastRow() > GMP_SCHEMA.layout.INCIDENT_HEADER_ROW) {
    sheet.setRowHeights(GMP_SCHEMA.layout.INCIDENT_HEADER_ROW + 1, sheet.getLastRow() - GMP_SCHEMA.layout.INCIDENT_HEADER_ROW, 18);
  }
  setIncidentsColumnVisibility_(sheet);
  applyIncidentConditionalFormatting_(sheet);
}

function writeTitleRows_(sheet, width, title, noteText) {
  mergeAcross_(sheet, GMP_SCHEMA.layout.TITLE_ROW, width);
  mergeAcross_(sheet, GMP_SCHEMA.layout.NOTES_ROW, width);

  var titleCell = sheet.getRange(GMP_SCHEMA.layout.TITLE_ROW, 1);
  titleCell.clearDataValidations();
  titleCell.setValue(title);
  titleCell.setBackground(GMP_SHEET_COLORS.TITLE);
  titleCell.setFontColor(GMP_SHEET_COLORS.TEXT_LIGHT);
  titleCell.setFontWeight('bold');
  titleCell.setFontSize(13);
  titleCell.setHorizontalAlignment('left');
  sheet.setRowHeight(GMP_SCHEMA.layout.TITLE_ROW, 24);

  var noteCell = sheet.getRange(GMP_SCHEMA.layout.NOTES_ROW, 1);
  noteCell.clearDataValidations();
  noteCell.setValue(noteText);
  noteCell.setBackground(GMP_SHEET_COLORS.NOTE_BG);
  noteCell.setFontColor('#0f172a');
  noteCell.setWrap(false);
  noteCell.setHorizontalAlignment('left');
  sheet.setRowHeight(GMP_SCHEMA.layout.NOTES_ROW, 20);
}

function writeSectionRow_(sheet, rowNumber, text, width, backgroundColor) {
  mergeAcross_(sheet, rowNumber, width);
  var cell = sheet.getRange(rowNumber, 1);
  cell.clearDataValidations();
  cell.setValue(text);
  cell.setBackground(backgroundColor);
  cell.setFontColor(GMP_SHEET_COLORS.TEXT_LIGHT);
  cell.setFontWeight('bold');
  cell.setHorizontalAlignment('left');
  sheet.setRowHeight(rowNumber, 20);
}

function writeHeaderRow_(sheet, rowNumber, headers, backgroundColor) {
  var range = sheet.getRange(rowNumber, 1, 1, headers.length);
  range.clearDataValidations();
  range.setValues([headers]);
  range.setBackground(backgroundColor);
  range.setFontColor(GMP_SHEET_COLORS.TEXT_LIGHT);
  range.setFontWeight('bold');
  range.setWrap(true);
  range.setHorizontalAlignment('center');
  range.setVerticalAlignment('middle');
  sheet.setRowHeight(rowNumber, 34);
}

function writeReviewBulkControlRow_(sheet) {
  var dateColumn = GMP_SCHEMA.reviews.columns.NEXT_REVIEW_DUE;
  var labelWidth = Math.max(dateColumn - 1, 1);
  var leadingRange = sheet.getRange(GMP_SCHEMA.layout.REVIEW_CONTROL_ROW, 1, 1, labelWidth);
  leadingRange.breakApart();
  leadingRange.clearDataValidations();
  if (labelWidth > 1) {
    leadingRange.merge();
  }

  var labelCell = sheet.getRange(GMP_SCHEMA.layout.REVIEW_CONTROL_ROW, 1);
  labelCell
    .setValue('Set all open reviews to:')
    .setBackground('#fff7ed')
    .setFontColor('#9a3412')
    .setFontWeight('bold')
    .setHorizontalAlignment('right')
    .setVerticalAlignment('middle');

  var dateCell = sheet.getRange(GMP_SCHEMA.layout.REVIEW_CONTROL_ROW, dateColumn);
  dateCell
    .clearDataValidations()
    .setValue('')
    .setBackground('#fdba74')
    .setFontColor('#7c2d12')
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setNumberFormat('yyyy-mm-dd')
    .setNote('Set a future date here to move all open reviews into Scheduled Follow-Ups.');
  dateCell.setBorder(true, true, true, true, false, false, '#c2410c', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  applyDateValidation_(dateCell);

  if (sheet.getMaxColumns() > dateColumn) {
    sheet.getRange(
      GMP_SCHEMA.layout.REVIEW_CONTROL_ROW,
      dateColumn + 1,
      1,
      sheet.getMaxColumns() - dateColumn
    ).clearContent().clearDataValidations().setBackground('#fff7ed');
  }

  sheet.setRowHeight(GMP_SCHEMA.layout.REVIEW_CONTROL_ROW, 22);
}

function mergeAcross_(sheet, rowNumber, width) {
  var maxColumns = Math.max(width, sheet.getMaxColumns());
  if (sheet.getMaxColumns() < width) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), width - sheet.getMaxColumns());
  }
  var rowRange = sheet.getRange(rowNumber, 1, 1, maxColumns);
  rowRange.breakApart();
  rowRange.clearDataValidations();
  sheet.getRange(rowNumber, 1, 1, width).merge();
}

function applyDropdowns_(sheet, startRow, rowCount, tableConfig) {
  if (rowCount <= 0) {
    return;
  }

  var fullRange = sheet.getRange(startRow, 1, rowCount, tableConfig.keys.length);
  fullRange.clearDataValidations();
  try {
    fullRange.removeCheckboxes();
  } catch (err) {}

  for (var i = 0; i < tableConfig.dropdowns.length; i++) {
    var dropdown = tableConfig.dropdowns[i];
    var column = tableConfig.columns[dropdown.key];
    var rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(dropdown.values, true)
      .setAllowInvalid(false)
      .build();
    sheet.getRange(startRow, column, rowCount, 1).setDataValidation(rule);
  }

  var checkboxes = tableConfig.checkboxes || [];
  for (var j = 0; j < checkboxes.length; j++) {
    var checkboxColumn = tableConfig.columns[checkboxes[j].key];
    sheet.getRange(startRow, checkboxColumn, rowCount, 1).insertCheckboxes();
  }
}

function applyDateValidation_(range) {
  if (!range) {
    return;
  }
  var rule = SpreadsheetApp.newDataValidation()
    .requireDate()
    .setAllowInvalid(false)
    .build();
  range.setDataValidation(rule);
}

function applyDateFormats_(sheet, columns, startRow, rowCount, formatPattern) {
  for (var i = 0; i < columns.length; i++) {
    sheet.getRange(startRow, columns[i], rowCount, 1).setNumberFormat(formatPattern);
  }
}

function applyDateFormatsByKeys_(sheet, columnMap, keys, startRow, rowCount, formatPattern) {
  var columns = [];
  for (var i = 0; i < keys.length; i++) {
    columns.push(columnMap[keys[i]]);
  }
  applyDateFormats_(sheet, columns, startRow, rowCount, formatPattern);
}

function setColumnWidths_(sheet, primaryWidths, secondaryWidths) {
  var widths = [];
  var maxLength = Math.max(primaryWidths.length, secondaryWidths ? secondaryWidths.length : 0);
  for (var i = 0; i < maxLength; i++) {
    widths.push(
      primaryWidths[i] ||
      (secondaryWidths && secondaryWidths[i]) ||
      120
    );
  }
  for (var i = 0; i < widths.length; i++) {
    sheet.setColumnWidth(i + 1, widthUnitsToPixels_(widths[i]));
  }
}

function widthUnitsToPixels_(widthUnits) {
  return Math.max(Math.round(widthUnits * 8), 72);
}

function applyStandardRowHeights_(sheet, rowNumbers) {
  for (var i = 0; i < rowNumbers.length; i++) {
    var rowNumber = rowNumbers[i];
    if (rowNumber === GMP_SCHEMA.layout.TITLE_ROW) {
      sheet.setRowHeight(rowNumber, 24);
    } else if (rowNumber === GMP_SCHEMA.layout.NOTES_ROW) {
      sheet.setRowHeight(rowNumber, 20);
    } else if (rowNumber === GMP_SCHEMA.layout.REQUEST_HEADER_ROW ||
               rowNumber === GMP_SCHEMA.layout.REVIEW_HEADER_ROW ||
               rowNumber === GMP_SCHEMA.layout.INCIDENT_HEADER_ROW ||
               rowNumber === rowNumbers[rowNumbers.length - 1]) {
      sheet.setRowHeight(rowNumber, 34);
    } else {
      sheet.setRowHeight(rowNumber, 20);
    }
  }
}

function applyAccessControlConditionalFormatting_(sheet, ctx) {
  var rules = [];
  rules = rules.concat(buildTextRules_(sheet, ctx.requestHeaderRow + 1, 2000, GMP_SCHEMA.requests.columns.QUEUE_PRIORITY, {
    High: '#fee2e2',
    Medium: '#fef3c7',
    Low: '#dbeafe'
  }));
  rules = rules.concat(buildTextRules_(sheet, ctx.requestHeaderRow + 1, 2000, GMP_SCHEMA.requests.columns.APPROVAL, {
    Submitted: '#fef3c7',
    Approved: '#dcfce7',
    Denied: '#fee2e2',
    Escalated: '#fce7f3'
  }));
  rules = rules.concat(buildTextRules_(sheet, ctx.requestHeaderRow + 1, 2000, GMP_SCHEMA.requests.columns.PROVISIONING, {
    Open: '#fee2e2',
    'In Progress': '#fef3c7',
    Completed: '#dcfce7',
    'Removing Access': '#fce7f3',
    Closed: '#e2e8f0'
  }));
  rules = rules.concat(buildTextRules_(sheet, ctx.accountHeaderRow + 1, 2000, GMP_SCHEMA.accounts.columns.CURRENT_STATE, {
    'Setting Up': '#dbeafe',
    'Access Granted': '#dcfce7',
    Unmanaged: '#fef3c7',
    Revoked: '#e2e8f0'
  }));
  rules = rules.concat(buildTextRules_(sheet, ctx.accountHeaderRow + 1, 2000, GMP_SCHEMA.accounts.columns.MFA, {
    Yes: '#dcfce7',
    No: '#fee2e2'
  }));
  sheet.setConditionalFormatRules(rules);
}

function applyReviewConditionalFormatting_(sheet) {
  var startRow = GMP_SCHEMA.layout.REVIEW_HEADER_ROW + 1;
  var rules = [];
  rules = rules.concat(buildTextRules_(sheet, startRow, 2000, GMP_SCHEMA.reviews.columns.REVIEW_PRIORITY, {
    High: '#fee2e2',
    Medium: '#fef3c7',
    Low: '#dbeafe'
  }));
  rules = rules.concat(buildTextRules_(sheet, startRow, 2000, GMP_SCHEMA.reviews.columns.PRESENCE_STATUS, {
    'Access Granted': '#dcfce7',
    'Setting Up': '#dbeafe',
    Unmanaged: '#fef3c7',
    Revoked: '#e2e8f0'
  }));
  rules = rules.concat(buildTextRules_(sheet, startRow, 2000, GMP_SCHEMA.reviews.columns.REVIEWER_ACTION, {
    Keep: '#dcfce7',
    Reduce: '#fef3c7',
    Remove: '#fee2e2',
    'Need Info': '#dbeafe'
  }));
  rules = rules.concat(buildTextRules_(sheet, startRow, 2000, GMP_SCHEMA.reviews.columns.DECISION_STATUS, {
    Open: '#fee2e2',
    Waiting: '#fef3c7',
    Done: '#dcfce7'
  }));
  sheet.setConditionalFormatRules(rules);
}

function highlightReviewCheckDateColumn_(sheet) {
  var column = GMP_SCHEMA.reviews.columns.NEXT_REVIEW_DUE;
  var headerCell = sheet.getRange(GMP_SCHEMA.layout.REVIEW_HEADER_ROW, column);
  headerCell
    .setBackground('#b45309')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  var rowCount = Math.max(sheet.getMaxRows() - GMP_SCHEMA.layout.REVIEW_CONTROL_ROW, 1);
  var range = sheet.getRange(GMP_SCHEMA.layout.REVIEW_CONTROL_ROW + 1, column, rowCount, 1);
  range
    .setBackground('#fff7ed')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setNumberFormat('yyyy-mm-dd');
}

function applyIncidentConditionalFormatting_(sheet) {
  var rules = [];
  rules = rules.concat(buildTextRules_(sheet, GMP_SCHEMA.layout.INCIDENT_HEADER_ROW + 1, 2000, GMP_SCHEMA.incidents.columns.SEVERITY, {
    High: '#fee2e2',
    Medium: '#fef3c7',
    Low: '#dbeafe'
  }));
  rules = rules.concat(buildTextRules_(sheet, GMP_SCHEMA.layout.INCIDENT_HEADER_ROW + 1, 2000, GMP_SCHEMA.incidents.columns.STATUS, {
    Open: '#fee2e2',
    Investigating: '#fef3c7',
    Resolved: '#dcfce7'
  }));
  sheet.setConditionalFormatRules(rules);
}

function buildTextRules_(sheet, startRow, rowCount, column, colorMap) {
  var rules = [];
  var keys = Object.keys(colorMap);
  var range = sheet.getRange(startRow, column, rowCount, 1);
  for (var i = 0; i < keys.length; i++) {
    rules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo(keys[i])
        .setBackground(colorMap[keys[i]])
        .setRanges([range])
        .build()
    );
  }
  return rules;
}

function applyDataRowFormatting_(sheet, startRow, endRow, columnCount) {
  if (startRow > endRow || columnCount <= 0) {
    return;
  }
  applyOperationalRowFormatting_(sheet.getRange(startRow, 1, endRow - startRow + 1, columnCount));
}

function setAccessControlColumnVisibility_(sheet) {
  var maxColumns = Math.max(sheet.getMaxColumns(), GMP_SCHEMA.requests.headers.length);
  if (sheet.getMaxColumns() < maxColumns) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), maxColumns - sheet.getMaxColumns());
  }

  sheet.showColumns(1, maxColumns);
  // Show first 11 columns (visible request + account columns), hide 12+ (internal automation)
  var visibleColumns = Math.max(11, GMP_SCHEMA.accounts.headers.length);
  if (maxColumns > visibleColumns) {
    sheet.hideColumns(visibleColumns + 1, maxColumns - visibleColumns);
  }
}

function setReviewChecksColumnVisibility_(sheet) {
  var maxColumns = Math.max(sheet.getMaxColumns(), GMP_SCHEMA.reviews.headers.length);
  sheet.showColumns(1, maxColumns);
  sheet.hideColumns(2, 1);                          // Review Cycle
  sheet.hideColumns(8, 1);                          // Activity Status
  var emailColumn = GMP_SCHEMA.reviews.columns.COMPANY_EMAIL;
  var reviewedByColumn = GMP_SCHEMA.reviews.columns.REVIEWED_BY;
  var lastVisibleColumn = Math.max(emailColumn || 13, reviewedByColumn || 13);
  if (lastVisibleColumn > 14) {
    sheet.hideColumns(14, lastVisibleColumn - 14);  // Hide Notes through Activity Score
  }
  if (maxColumns > lastVisibleColumn) {
    sheet.hideColumns(lastVisibleColumn + 1, maxColumns - lastVisibleColumn);
  }
}

function setIncidentsColumnVisibility_(sheet) {
  var maxColumns = Math.max(sheet.getMaxColumns(), GMP_SCHEMA.incidents.headers.length);
  sheet.showColumns(1, maxColumns);
  if (maxColumns > 12) {
    sheet.hideColumns(13, maxColumns - 12);
  }
}

function migrateLegacyAccessControlAccountSection_(sheet) {
  var requestSectionRow = findLabelRow_(sheet, GMP_SCHEMA.sectionLabels.REQUESTS);
  var accountSectionRow = findLabelRow_(sheet, GMP_SCHEMA.sectionLabels.ACCOUNTS);
  if (requestSectionRow === -1 || accountSectionRow === -1) {
    return false;
  }

  var requestHeaderRow = requestSectionRow + 1;
  var accountHeaderRow = accountSectionRow + 1;
  if (!sheetHeaderMatches_(sheet, requestHeaderRow, GMP_SCHEMA.requests.headers)) {
    return false;
  }

  if (!sheetHeaderMatches_(sheet, accountHeaderRow, LEGACY_ACCOUNT_HEADERS)) {
    return false;
  }

  remapSectionRowsByHeaders_(
    sheet,
    accountHeaderRow + 1,
    GMP_SCHEMA.accounts.headers.length,
    LEGACY_ACCOUNT_HEADERS,
    GMP_SCHEMA.accounts.headers
  );
  return true;
}

function migrateLegacyReviewChecksSheet_(sheet) {
  if (!sheetHeaderMatches_(sheet, GMP_SCHEMA.layout.REVIEW_HEADER_ROW, LEGACY_REVIEW_HEADERS)) {
    return false;
  }

  remapSectionRowsByHeaders_(
    sheet,
    GMP_SCHEMA.layout.REVIEW_HEADER_ROW + 1,
    LEGACY_REVIEW_HEADERS.length,
    LEGACY_REVIEW_HEADERS,
    GMP_SCHEMA.reviews.headers
  );
  return true;
}

function migrateLegacyIncidentsSheet_(sheet) {
  if (!sheetHeaderMatches_(sheet, GMP_SCHEMA.layout.INCIDENT_HEADER_ROW, LEGACY_INCIDENT_HEADERS)) {
    return false;
  }

  remapSectionRowsByHeaders_(
    sheet,
    GMP_SCHEMA.layout.INCIDENT_HEADER_ROW + 1,
    LEGACY_INCIDENT_HEADERS.length,
    LEGACY_INCIDENT_HEADERS,
    GMP_SCHEMA.incidents.headers
  );
  return true;
}

function remapSectionRowsByHeaders_(sheet, startRow, columnCount, oldHeaders, newHeaders) {
  var endRow = sheet.getLastRow();
  if (endRow < startRow) {
    return;
  }

  if (sheet.getMaxColumns() < newHeaders.length) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), newHeaders.length - sheet.getMaxColumns());
  }

  var values = sheet.getRange(startRow, 1, endRow - startRow + 1, columnCount).getValues();
  var remapped = [];

  for (var i = 0; i < values.length; i++) {
    var row = values[i];
    if (isBlankRow_(row)) {
      remapped.push(new Array(newHeaders.length).fill(''));
      continue;
    }

    var valueByHeader = {};
    for (var j = 0; j < oldHeaders.length; j++) {
      valueByHeader[oldHeaders[j]] = row[j];
    }

    var nextRow = [];
    for (var k = 0; k < newHeaders.length; k++) {
      var value = valueByHeader.hasOwnProperty(newHeaders[k]) ? valueByHeader[newHeaders[k]] : '';
      if (value === '' && HEADER_ALIASES[newHeaders[k]]) {
        for (var m = 0; m < HEADER_ALIASES[newHeaders[k]].length; m++) {
          var alias = HEADER_ALIASES[newHeaders[k]][m];
          if (valueByHeader.hasOwnProperty(alias) && valueByHeader[alias] !== '') {
            value = valueByHeader[alias];
            break;
          }
        }
      }
      nextRow.push(value);
    }
    remapped.push(nextRow);
  }

  sheet.getRange(startRow, 1, remapped.length, newHeaders.length).setValues(remapped);
}
