function normalizeText_(value) {
  if (value === null || value === undefined) {
    return '';
  }
  return String(value).trim();
}

function normalizeIdentifier_(value) {
  var text = normalizeText_(value);
  var lowered = text.toLowerCase();
  if (!text ||
      lowered === 'none' ||
      lowered === 'null' ||
      lowered === 'undefined' ||
      lowered === 'n/a' ||
      lowered === 'na') {
    return '';
  }
  return text;
}

function accountUniqueSignatureKey_(email, systemName, accessLevel, platformAccount, person) {
  var normalizedSystem = canonicalSystemName_(systemName).toLowerCase();
  var normalizedAccess = normalizeText_(accessLevel).toLowerCase();
  var normalizedEmail = normalizeText_(email).toLowerCase();
  if (normalizedEmail) {
    return 'email:' + normalizedEmail + '|' + normalizedSystem + '|' + normalizedAccess;
  }

  var normalizedPlatform = normalizeText_(platformAccount).toLowerCase();
  if (normalizedPlatform) {
    return 'platform:' + normalizedPlatform + '|' + normalizedSystem + '|' + normalizedAccess;
  }

  return 'person:' + normalizeText_(person).toLowerCase() + '|' + normalizedSystem + '|' + normalizedAccess;
}

function reviewSignatureKey_(person, systemName, accessLevel) {
  var normalizedPerson = normalizeText_(person).toLowerCase();
  var normalizedSystem = canonicalSystemName_(systemName).toLowerCase();
  var normalizedAccess = normalizeText_(accessLevel).toLowerCase();
  if (!normalizedPerson || !normalizedSystem) {
    return '';
  }
  return 'sig:' + normalizedPerson + '|' + normalizedSystem + '|' + normalizedAccess;
}

function reviewAccountIdentityKey_(accessId, person, systemName, accessLevel) {
  var normalizedAccessId = normalizeIdentifier_(accessId);
  if (normalizedAccessId) {
    return 'id:' + normalizedAccessId;
  }
  return reviewSignatureKey_(person, systemName, accessLevel);
}

function reviewIdentityKey_(review) {
  return reviewAccountIdentityKey_(review.ACCESS_ID, review.PERSON, review.SYSTEM, review.ACCESS_LEVEL);
}

function accountIdentityKey_(account) {
  return reviewAccountIdentityKey_(account.ACCESS_ID, account.PERSON, account.GMP_SYSTEM, account.ACCESS_LEVEL);
}

function accountUniqueKey_(account) {
  return accountUniqueSignatureKey_(account.COMPANY_EMAIL, account.GMP_SYSTEM, account.ACCESS_LEVEL, account.PLATFORM_ACCOUNT, account.PERSON);
}

function reviewReusableIdentityKey_(review) {
  return reviewSignatureKey_(review.PERSON, review.SYSTEM, review.ACCESS_LEVEL) ||
    reviewAccountIdentityKey_(review.ACCESS_ID, review.PERSON, review.SYSTEM, review.ACCESS_LEVEL);
}

function accountReusableReviewIdentityKey_(account) {
  return reviewSignatureKey_(account.PERSON, account.GMP_SYSTEM, account.ACCESS_LEVEL) ||
    reviewAccountIdentityKey_(account.ACCESS_ID, account.PERSON, account.GMP_SYSTEM, account.ACCESS_LEVEL);
}

var ACCOUNT_SYSTEM_GROUPS = [
  'Katana',
  'Wasp',
  'Shopify',
  'ShipStation',
  '1Password',
  'Other Systems'
];

var LEGACY_REVIEW_HEADERS_WITH_ACTION_BUTTONS = [
  'Review ID',
  'Review Cycle',
  'Priority',
  'Person',
  'System',
  'Access Level',
  'Account Status',
  'Activity Status',
  'Check Needed',
  'Reviewer Action',
  'Review Status',
  '30d',
  'Done',
  'Reviewed At',
  'Next Check Date',
  'Notes',
  'Trigger Type',
  'Access ID',
  'Linked Request',
  'Slack Thread',
  'Next Slack Trigger',
  'Last Activity Evidence',
  '30d Logins',
  'Days Since Login',
  'Activity Score'
];

function formatDateOnly_(value) {
  if (!(value instanceof Date) || isNaN(value)) {
    return '';
  }
  return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function formatTimestamp_(value) {
  if (!(value instanceof Date) || isNaN(value)) {
    return '';
  }
  return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
}

function toInteger_(value, defaultValue) {
  var parsed = parseInt(value, 10);
  return isNaN(parsed) ? (defaultValue || 0) : parsed;
}

function addDays_(date, days) {
  var value = new Date(date.getTime());
  value.setDate(value.getDate() + days);
  return value;
}

function addHours_(date, hours) {
  var value = new Date(date.getTime());
  value.setHours(value.getHours() + hours);
  return value;
}

function daysBetween_(olderDate, newerDate) {
  if (!(olderDate instanceof Date) || isNaN(olderDate)) {
    return '';
  }
  var start = new Date(olderDate.getTime());
  var end = newerDate instanceof Date && !isNaN(newerDate) ? new Date(newerDate.getTime()) : new Date();
  start.setHours(0, 0, 0, 0);
  end.setHours(0, 0, 0, 0);
  return Math.floor((end - start) / 86400000);
}

function isYesValue_(value) {
  return normalizeText_(value).toLowerCase() === 'yes';
}

function isBlankRow_(rowValues) {
  for (var i = 0; i < rowValues.length; i++) {
    if (normalizeText_(rowValues[i]) !== '') {
      return false;
    }
  }
  return true;
}

function objectToRow_(data, keys) {
  var row = [];
  for (var i = 0; i < keys.length; i++) {
    var key = keys[i];
    row.push(data.hasOwnProperty(key) ? data[key] : '');
  }
  return row;
}

function replaceRangeValues_(range, values) {
  if (!range) {
    return;
  }
  range.clearDataValidations();
  range.setValues(values);
}

function rowToObject_(rowValues, keys, columnMap) {
  var item = {};
  for (var i = 0; i < keys.length; i++) {
    var key = keys[i];
    item[key] = rowValues[columnMap[key] - 1];
  }
  return item;
}

function nextStableId_(prefix) {
  var props = PropertiesService.getScriptProperties();
  var year = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy');
  var propertyName = 'COUNTER_' + prefix + '_' + year;
  var next = toInteger_(props.getProperty(propertyName), 0) + 1;
  props.setProperty(propertyName, String(next));
  return prefix + '-' + year + '-' + String(next).padStart(4, '0');
}

function setWorkbookSchemaVersion_() {
  PropertiesService.getDocumentProperties().setProperty(
    GMP_CONFIG.SCHEMA_VERSION_PROPERTY,
    GMP_CONFIG.SCHEMA_VERSION
  );
}

function ensureThreeTabWorkbookReady_() {
  try {
    normalizeAccessControlHeaders_();
  } catch (err) {}
  try {
    dedupeAccountRowsByUniqueSignature_();
  } catch (err) {}
  try {
    ensureReviewControlRow_();
  } catch (err) {}
  try {
    normalizeReviewCheckHeaders_();
  } catch (err) {}
  try {
    normalizeReviewRowLayouts_();
  } catch (err) {}
  try {
    dedupeReviewRowsByReusableIdentity_();
  } catch (err) {}
  try {
    dedupeActiveReviewRowsByAccessId_();
  } catch (err) {}
  try {
    dedupeReviewRowsByCycleSignature_();
  } catch (err) {}
  try {
    dedupeReviewRowsById_();
  } catch (err) {}
  var issues = getThreeTabWorkbookIssues_();
  if (issues.length > 0) {
    throw new Error(issues.join(' '));
  }
  normalizeAccessControlTerminology_();
  normalizeReviewRecords_();
  normalizeIncidentRecords_();
  syncAccountReviewRollupsFromReviews_();
  normalizeReviewRecords_();
}

function getThreeTabWorkbookIssues_() {
  var issues = [];
  var version = PropertiesService.getDocumentProperties().getProperty(GMP_CONFIG.SCHEMA_VERSION_PROPERTY);
  if (version !== GMP_CONFIG.SCHEMA_VERSION) {
    issues.push('Workbook schema is not initialized. Run setupSheet() before using Slack intake or automation.');
    return issues;
  }

  var accessSheet = getRequiredSheet_(GMP_SCHEMA.tabs.ACCESS_CONTROL);
  var reviewSheet = getRequiredSheet_(GMP_SCHEMA.tabs.REVIEW_CHECKS);
  var incidentSheet = getRequiredSheet_(GMP_SCHEMA.tabs.INCIDENTS);

  if (findLabelRow_(accessSheet, GMP_SCHEMA.sectionLabels.REQUESTS) === -1 ||
      findLabelRow_(accessSheet, GMP_SCHEMA.sectionLabels.ACCOUNTS) === -1) {
    issues.push('Access Control sheet does not match the current schema.');
  }
  if (!sheetHeaderMatches_(reviewSheet, GMP_SCHEMA.layout.REVIEW_HEADER_ROW, GMP_SCHEMA.reviews.headers)) {
    issues.push('Review Checks sheet does not match the current schema.');
  }
  if (!sheetHeaderMatches_(incidentSheet, GMP_SCHEMA.layout.INCIDENT_HEADER_ROW, GMP_SCHEMA.incidents.headers)) {
    issues.push('Incidents sheet does not match the current schema.');
  }
  return issues;
}

function getRequiredSheet_(sheetName) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) {
    throw new Error('Missing required sheet: ' + sheetName);
  }
  return sheet;
}

function findLabelRow_(sheet, label) {
  var maxRows = Math.max(sheet.getLastRow(), 20);
  var values = sheet.getRange(1, 1, maxRows, 1).getDisplayValues();
  for (var i = 0; i < values.length; i++) {
    if (normalizeText_(values[i][0]) === label) {
      return i + 1;
    }
  }
  return -1;
}

function sheetHeaderMatches_(sheet, rowNumber, expectedHeaders) {
  if (sheet.getLastRow() < rowNumber) {
    return false;
  }
  var values = sheet.getRange(rowNumber, 1, 1, expectedHeaders.length).getDisplayValues()[0];
  for (var i = 0; i < expectedHeaders.length; i++) {
    if (!headerMatchesExpected_(values[i], expectedHeaders[i])) {
      return false;
    }
  }
  return true;
}

function headerMatchesExpected_(actualValue, expectedHeader) {
  var actual = normalizeText_(actualValue);
  if (expectedHeader === 'IT Owner' && (actual === 'Provisioning Owner' || actual === 'Owner')) {
    return true;
  }
  if (expectedHeader === 'Access Setup Status' && actual === 'Provisioning') {
    return true;
  }
  if (expectedHeader === 'Account Status' && (actual === 'Presence Status' || actual === 'Current State')) {
    return true;
  }
  if (expectedHeader === 'Last Verified In System' && actual === 'Granted On') {
    return true;
  }
  if (expectedHeader === 'Activity Status' && actual === 'Activity Band') {
    return true;
  }
  if (expectedHeader === 'Last Activity Evidence' && actual === 'Last Login') {
    return true;
  }
  if (expectedHeader === 'Review Cycle' && actual === 'Next Review Due') {
    return true;
  }
  if (expectedHeader === 'Next Check Date' && (actual === 'Check Date' || actual === 'Next Review Due')) {
    return true;
  }
  if (expectedHeader === 'Check Needed' && (actual === 'Why In Review' || actual === 'Why Flagged')) {
    return true;
  }
  if (expectedHeader === 'Review Status' && actual === 'Decision Status') {
    return true;
  }
  if (expectedHeader === 'Reviewed At' && (actual === 'Completed At' || actual === 'Decision Date')) {
    return true;
  }
  if (expectedHeader === 'Audit Log' && actual === 'Notes') {
    return true;
  }
  return actual === expectedHeader || actual.indexOf(expectedHeader + ' ') === 0;
}

function normalizeHeaderText_(value) {
  var actual = normalizeText_(value);
  var knownHeaders = getKnownHeaders_();
  for (var i = 0; i < knownHeaders.length; i++) {
    if (headerMatchesExpected_(actual, knownHeaders[i])) {
      return knownHeaders[i];
    }
  }
  return actual;
}

function getKnownHeaders_() {
  return []
    .concat(GMP_SCHEMA.requests.headers)
    .concat(GMP_SCHEMA.accounts.headers)
    .concat(GMP_SCHEMA.reviews.headers)
    .concat(GMP_SCHEMA.incidents.headers);
}

function normalizeAccessControlHeaders_() {
  var ctx = getAccessControlContext_();
  replaceRangeValues_(
    ctx.sheet.getRange(ctx.requestHeaderRow, 1, 1, GMP_SCHEMA.requests.headers.length),
    [GMP_SCHEMA.requests.headers]
  );
  replaceRangeValues_(
    ctx.sheet.getRange(ctx.accountHeaderRow, 1, 1, GMP_SCHEMA.accounts.headers.length),
    [GMP_SCHEMA.accounts.headers]
  );
  if (ctx.archiveSectionRow !== -1) {
    replaceRangeValues_(
      ctx.sheet.getRange(ctx.archiveHeaderRow, 1, 1, GMP_SCHEMA.requests.headers.length),
      [GMP_SCHEMA.requests.headers]
    );
  }
}

function normalizeReviewCheckHeaders_() {
  var sheet = getRequiredSheet_(GMP_SCHEMA.tabs.REVIEW_CHECKS);
  if (sheet.getLastRow() >= GMP_SCHEMA.layout.REVIEW_HEADER_ROW) {
    var currentHeaders = sheet.getRange(
      GMP_SCHEMA.layout.REVIEW_HEADER_ROW,
      1,
      1,
      LEGACY_REVIEW_HEADERS_WITH_ACTION_BUTTONS.length
    ).getDisplayValues()[0];
    var matchesLegacyActionLayout = true;
    for (var i = 0; i < LEGACY_REVIEW_HEADERS_WITH_ACTION_BUTTONS.length; i++) {
      if (normalizeHeaderText_(currentHeaders[i]) !== LEGACY_REVIEW_HEADERS_WITH_ACTION_BUTTONS[i]) {
        matchesLegacyActionLayout = false;
        break;
      }
    }

    if (matchesLegacyActionLayout) {
      remapSectionRowsByHeaders_(
        sheet,
        GMP_SCHEMA.layout.REVIEW_CONTROL_ROW + 1,
        LEGACY_REVIEW_HEADERS_WITH_ACTION_BUTTONS.length,
        LEGACY_REVIEW_HEADERS_WITH_ACTION_BUTTONS,
        GMP_SCHEMA.reviews.headers
      );
    }
  }
  replaceRangeValues_(
    sheet.getRange(GMP_SCHEMA.layout.REVIEW_HEADER_ROW, 1, 1, GMP_SCHEMA.reviews.headers.length),
    [GMP_SCHEMA.reviews.headers]
  );
}

function ensureReviewControlRow_() {
  var sheet = getRequiredSheet_(GMP_SCHEMA.tabs.REVIEW_CHECKS);
  var controlLabel = normalizeText_(sheet.getRange(GMP_SCHEMA.layout.REVIEW_CONTROL_ROW, 1).getDisplayValue());
  if (controlLabel === 'Set all open reviews to:') {
    return;
  }

  var rowValues = sheet.getRange(
    GMP_SCHEMA.layout.REVIEW_CONTROL_ROW,
    1,
    1,
    Math.min(sheet.getMaxColumns(), GMP_SCHEMA.reviews.headers.length)
  ).getDisplayValues()[0];
  if (isBlankRow_(rowValues)) {
    return;
  }

  sheet.insertRowsBefore(GMP_SCHEMA.layout.REVIEW_CONTROL_ROW, 1);
}

function normalizeReviewRowLayouts_() {
  var sheet = getRequiredSheet_(GMP_SCHEMA.tabs.REVIEW_CHECKS);
  var startRow = GMP_SCHEMA.layout.REVIEW_CONTROL_ROW + 1;
  var endRow = sheet.getLastRow();
  if (endRow < startRow || sheet.getLastColumn() < LEGACY_REVIEW_HEADERS_WITH_ACTION_BUTTONS.length) {
    return;
  }

  var values = sheet.getRange(
    startRow,
    1,
    endRow - startRow + 1,
    LEGACY_REVIEW_HEADERS_WITH_ACTION_BUTTONS.length
  ).getValues();

  var shouldRemap = false;
  for (var i = 0; i < values.length; i++) {
    var row = values[i];
    if (isBlankRow_(row) || isDisplayGroupRow_(row)) {
      continue;
    }

    if (typeof row[11] === 'boolean' || typeof row[12] === 'boolean') {
      shouldRemap = true;
      break;
    }

    if (isReviewTriggerLabel_(row[12]) || isReviewTriggerLabel_(row[14])) {
      shouldRemap = true;
      break;
    }
  }

  if (!shouldRemap) {
    return;
  }

  remapSectionRowsByHeaders_(
    sheet,
    startRow,
    LEGACY_REVIEW_HEADERS_WITH_ACTION_BUTTONS.length,
    LEGACY_REVIEW_HEADERS_WITH_ACTION_BUTTONS,
    GMP_SCHEMA.reviews.headers
  );
}

function isReviewTriggerLabel_(value) {
  var text = normalizeText_(value);
  return text === '120-Day Review Due' ||
    text === 'Unmanaged Account' ||
    text === 'Missing From System' ||
    text === 'No Activity Evidence';
}

function normalizeAccessControlTerminology_() {
  normalizeRequestTerminology_(getRequestRows_());
  normalizeRequestTerminology_(getArchivedRequestRows_());
  normalizeAccountTerminology_(getAccountRows_());
}

function normalizeReviewRecords_() {
  var rows = getReviewRows_();
  var today = new Date();
  var todayStart = new Date(today.getFullYear(), today.getMonth(), today.getDate());
  for (var i = 0; i < rows.length; i++) {
    var record = rows[i].data;
    var shouldUpdate = false;
    var nextCheckValue = record.NEXT_REVIEW_DUE;
    var decisionStatus = normalizeText_(record.DECISION_STATUS);
    var recoveredDecisionParts = recoverReviewMetadataFromDecisionCell_(record.DECISION_DATE);
    if (recoveredDecisionParts.accessId && !normalizeIdentifier_(record.ACCESS_ID)) {
      record.ACCESS_ID = recoveredDecisionParts.accessId;
      shouldUpdate = true;
    }
    if (recoveredDecisionParts.notes) {
      var mergedNotes = mergeNoteBlocks_(record.NOTES, recoveredDecisionParts.notes);
      if (mergedNotes !== normalizeText_(record.NOTES)) {
        record.NOTES = mergedNotes;
        shouldUpdate = true;
      }
    }
    if (recoveredDecisionParts.hadNonDateContent && decisionStatus !== 'Done') {
      record.DECISION_DATE = '';
      shouldUpdate = true;
    }

    if (isReviewTriggerLabel_(nextCheckValue)) {
      if (!normalizeText_(record.TRIGGER_TYPE)) {
        record.TRIGGER_TYPE = normalizeText_(nextCheckValue);
      }
      record.NEXT_REVIEW_DUE = '';
      shouldUpdate = true;
    } else if (typeof nextCheckValue === 'boolean') {
      record.NEXT_REVIEW_DUE = '';
      shouldUpdate = true;
    }

    if (nextCheckValue instanceof Date && !isNaN(nextCheckValue)) {
      var nextCheckDay = new Date(nextCheckValue.getFullYear(), nextCheckValue.getMonth(), nextCheckValue.getDate());
      if (nextCheckDay.getFullYear() < 2025) {
        record.NEXT_REVIEW_DUE = '';
        shouldUpdate = true;
      } else if (nextCheckDay <= todayStart) {
        record.DECISION_STATUS = 'Open';
        record.NEXT_REVIEW_DUE = '';
        shouldUpdate = true;
      } else if (decisionStatus === 'Waiting') {
        record.DECISION_STATUS = 'Open';
        shouldUpdate = true;
      }
    } else if (normalizeText_(nextCheckValue) && decisionStatus !== 'Done') {
      record.NEXT_REVIEW_DUE = '';
      shouldUpdate = true;
    }

    if (decisionStatus === 'Done' && normalizeText_(record.NEXT_REVIEW_DUE)) {
      record.NEXT_REVIEW_DUE = '';
      shouldUpdate = true;
    }

    var accountRow = findAccountRowForReview_(record);
    if (accountRow) {
      var accountPresenceStatus = normalizePresenceState_(accountRow.data.CURRENT_STATE);
      if (accountPresenceStatus && normalizeText_(record.PRESENCE_STATUS) !== accountPresenceStatus) {
        record.PRESENCE_STATUS = accountPresenceStatus;
        shouldUpdate = true;
      }
      if (normalizeText_(record.ACCESS_LEVEL) !== normalizeText_(accountRow.data.ACCESS_LEVEL)) {
        record.ACCESS_LEVEL = accountRow.data.ACCESS_LEVEL;
        shouldUpdate = true;
      }
      if (!normalizeIdentifier_(record.ACCESS_ID) && normalizeIdentifier_(accountRow.data.ACCESS_ID)) {
        record.ACCESS_ID = accountRow.data.ACCESS_ID;
        shouldUpdate = true;
      }
      if (decisionStatus !== 'Done' &&
          (!(record.NEXT_REVIEW_DUE instanceof Date) || isNaN(record.NEXT_REVIEW_DUE)) &&
          accountRow.data.NEXT_REVIEW_DUE instanceof Date &&
          !isNaN(accountRow.data.NEXT_REVIEW_DUE)) {
        record.NEXT_REVIEW_DUE = accountRow.data.NEXT_REVIEW_DUE;
        shouldUpdate = true;
      }
    }

    if (decisionStatus === 'Done' && !(record.DECISION_DATE instanceof Date)) {
      if (accountRow && accountRow.data.LAST_REVIEW_DATE instanceof Date && !isNaN(accountRow.data.LAST_REVIEW_DATE)) {
        record.DECISION_DATE = accountRow.data.LAST_REVIEW_DATE;
        shouldUpdate = true;
      } else if (normalizeText_(record.DECISION_DATE)) {
        record.DECISION_DATE = '';
        shouldUpdate = true;
      }
    }

    if (shouldUpdate) {
      updateReviewRow_(rows[i].rowNumber, record);
    }
  }
}

function recoverReviewMetadataFromDecisionCell_(value) {
  var text = normalizeText_(value);
  if (!text || value instanceof Date) {
    return {
      accessId: '',
      notes: '',
      hadNonDateContent: false
    };
  }

  var parts = text.split(/\r?\n+/);
  var accessId = '';
  var noteLines = [];
  for (var i = 0; i < parts.length; i++) {
    var line = normalizeText_(parts[i]);
    if (!line) {
      continue;
    }
    if (!accessId && /^ACC-\d{4}-\d{4}$/i.test(line)) {
      accessId = line;
      continue;
    }
    noteLines.push(line);
  }

  return {
    accessId: accessId,
    notes: noteLines.join('\n'),
    hadNonDateContent: true
  };
}

function normalizeIncidentRecords_() {
  var rows = getIncidentRows_();
  for (var i = 0; i < rows.length; i++) {
    var record = rows[i].data;
    var compactNotes = compactIncidentNotes_(record.NOTES);
    if (compactNotes !== normalizeText_(record.NOTES)) {
      record.NOTES = compactNotes;
      updateIncidentRow_(rows[i].rowNumber, record);
    }
  }
}

function dedupeReviewRowsById_() {
  var sheet = getRequiredSheet_(GMP_SCHEMA.tabs.REVIEW_CHECKS);
  var rows = getReviewRows_();
  var grouped = {};
  var ids = [];

  for (var i = 0; i < rows.length; i++) {
    var reviewId = normalizeText_(rows[i].data.REVIEW_ID);
    if (!reviewId) {
      continue;
    }
    if (!grouped[reviewId]) {
      grouped[reviewId] = [];
      ids.push(reviewId);
    }
    grouped[reviewId].push(rows[i]);
  }

  var rowsToDelete = [];
  for (var j = 0; j < ids.length; j++) {
    var duplicates = grouped[ids[j]];
    if (!duplicates || duplicates.length < 2) {
      continue;
    }

    var keeper = chooseReviewRowToKeep_(duplicates);
    for (var k = 0; k < duplicates.length; k++) {
      if (duplicates[k].rowNumber !== keeper.rowNumber) {
        rowsToDelete.push(duplicates[k].rowNumber);
      }
    }
  }

  if (!rowsToDelete.length) {
    return;
  }

  rowsToDelete.sort(function(a, b) { return b - a; });
  for (var m = 0; m < rowsToDelete.length; m++) {
    sheet.deleteRow(rowsToDelete[m]);
  }
}

function dedupeReviewRowsByCycleSignature_() {
  var sheet = getRequiredSheet_(GMP_SCHEMA.tabs.REVIEW_CHECKS);
  var rows = getReviewRows_();
  var grouped = {};
  var keys = [];

  for (var i = 0; i < rows.length; i++) {
    var reviewIdentity = reviewIdentityKey_(rows[i].data);
    var reviewCycle = normalizeText_(rows[i].data.REVIEW_CYCLE);
    if (!reviewIdentity || !reviewCycle) {
      continue;
    }

    var groupKey = reviewIdentity + '|' + reviewCycle;
    if (!grouped[groupKey]) {
      grouped[groupKey] = [];
      keys.push(groupKey);
    }
    grouped[groupKey].push(rows[i]);
  }

  var rowsToDelete = [];
  for (var j = 0; j < keys.length; j++) {
    var duplicates = grouped[keys[j]];
    if (!duplicates || duplicates.length < 2) {
      continue;
    }

    var keeper = chooseReviewRowToKeep_(duplicates);
    for (var k = 0; k < duplicates.length; k++) {
      if (duplicates[k].rowNumber !== keeper.rowNumber) {
        rowsToDelete.push(duplicates[k].rowNumber);
      }
    }
  }

  if (!rowsToDelete.length) {
    return;
  }

  rowsToDelete.sort(function(a, b) { return b - a; });
  for (var m = 0; m < rowsToDelete.length; m++) {
    sheet.deleteRow(rowsToDelete[m]);
  }
}

function dedupeActiveReviewRowsByAccessId_() {
  var sheet = getRequiredSheet_(GMP_SCHEMA.tabs.REVIEW_CHECKS);
  var rows = getReviewRows_();
  var grouped = {};
  var keys = [];

  for (var i = 0; i < rows.length; i++) {
    var review = rows[i].data;
    if (normalizeText_(review.DECISION_STATUS) === 'Done') {
      continue;
    }

    var reviewIdentity = reviewIdentityKey_(review);
    if (!reviewIdentity) {
      continue;
    }

    if (!grouped[reviewIdentity]) {
      grouped[reviewIdentity] = [];
      keys.push(reviewIdentity);
    }
    grouped[reviewIdentity].push(rows[i]);
  }

  var rowsToDelete = [];
  for (var j = 0; j < keys.length; j++) {
    var duplicates = grouped[keys[j]];
    if (!duplicates || duplicates.length < 2) {
      continue;
    }

    var keeper = chooseReviewRowToKeep_(duplicates);
    for (var k = 0; k < duplicates.length; k++) {
      if (duplicates[k].rowNumber !== keeper.rowNumber) {
        rowsToDelete.push(duplicates[k].rowNumber);
      }
    }
  }

  if (!rowsToDelete.length) {
    return;
  }

  rowsToDelete.sort(function(a, b) { return b - a; });
  for (var m = 0; m < rowsToDelete.length; m++) {
    sheet.deleteRow(rowsToDelete[m]);
  }
}

function dedupeReviewRowsByReusableIdentity_() {
  var sheet = getRequiredSheet_(GMP_SCHEMA.tabs.REVIEW_CHECKS);
  var rows = getReviewRows_();
  var grouped = {};
  var keys = [];

  for (var i = 0; i < rows.length; i++) {
    var identityKey = reviewReusableIdentityKey_(rows[i].data);
    if (!identityKey) {
      continue;
    }

    if (!grouped[identityKey]) {
      grouped[identityKey] = [];
      keys.push(identityKey);
    }
    grouped[identityKey].push(rows[i]);
  }

  var rowsToDelete = [];
  for (var j = 0; j < keys.length; j++) {
    var duplicates = grouped[keys[j]];
    if (!duplicates || duplicates.length < 2) {
      continue;
    }

    var keeper = chooseReusableReviewRowForIdentity_(duplicates);
    for (var k = 0; k < duplicates.length; k++) {
      if (duplicates[k].rowNumber !== keeper.rowNumber) {
        rowsToDelete.push(duplicates[k].rowNumber);
      }
    }
  }

  if (!rowsToDelete.length) {
    return;
  }

  rowsToDelete.sort(function(a, b) { return b - a; });
  for (var m = 0; m < rowsToDelete.length; m++) {
    sheet.deleteRow(rowsToDelete[m]);
  }
}

function chooseReusableReviewRowForIdentity_(rows) {
  var current = chooseCurrentReviewRowForAccount_(rows);
  if (current) {
    return current;
  }
  return chooseLatestDoneReviewRowForAccount_(rows) || rows[0];
}

function dedupeAccountRowsByUniqueSignature_() {
  var sheet = getRequiredSheet_(GMP_SCHEMA.tabs.ACCESS_CONTROL);
  var rows = getAccountRows_();
  var grouped = {};
  var keys = [];

  for (var i = 0; i < rows.length; i++) {
    var uniqueKey = accountUniqueKey_(rows[i].data);
    if (!uniqueKey) {
      continue;
    }
    if (!grouped[uniqueKey]) {
      grouped[uniqueKey] = [];
      keys.push(uniqueKey);
    }
    grouped[uniqueKey].push(rows[i]);
  }

  var rowsToDelete = [];
  for (var j = 0; j < keys.length; j++) {
    var duplicates = grouped[keys[j]];
    if (!duplicates || duplicates.length < 2) {
      continue;
    }

    var keeper = chooseAccountRowToKeep_(duplicates);
    var merged = mergeDuplicateAccountRows_(duplicates, keeper);
    updateAccountRow_(keeper.rowNumber, merged);
    relinkRowsToMergedAccount_(merged, duplicates, keeper.rowNumber);

    for (var k = 0; k < duplicates.length; k++) {
      if (duplicates[k].rowNumber !== keeper.rowNumber) {
        rowsToDelete.push(duplicates[k].rowNumber);
      }
    }
  }

  if (!rowsToDelete.length) {
    return;
  }

  rowsToDelete.sort(function(a, b) { return b - a; });
  for (var m = 0; m < rowsToDelete.length; m++) {
    sheet.deleteRow(rowsToDelete[m]);
  }
}

function chooseAccountRowToKeep_(rows) {
  var bestRow = rows[0];
  var bestScore = accountRowKeepScore_(rows[0]);
  for (var i = 1; i < rows.length; i++) {
    var score = accountRowKeepScore_(rows[i]);
    if (score > bestScore || (score === bestScore && rows[i].rowNumber > bestRow.rowNumber)) {
      bestRow = rows[i];
      bestScore = score;
    }
  }
  return bestRow;
}

function accountRowKeepScore_(row) {
  var account = row.data;
  var score = 0;
  if (normalizeIdentifier_(account.ACCESS_ID)) {
    score += 400;
  }
  if (normalizeText_(account.SOURCE_REQUEST_ID)) {
    score += 160;
  }
  if (normalizeText_(account.REVIEW_STATUS) === 'Open') {
    score += 120;
  } else if (normalizeText_(account.REVIEW_STATUS) === 'Waiting') {
    score += 100;
  } else if (normalizeText_(account.REVIEW_STATUS) === 'Done') {
    score += 80;
  }

  var state = normalizePresenceState_(account.CURRENT_STATE);
  if (state === 'Present') {
    score += 70;
  } else if (state === 'Unmanaged') {
    score += 60;
  } else if (state === 'Missing') {
    score += 50;
  } else if (state === 'Setting Up') {
    score += 40;
  } else if (state === 'Revoked') {
    score += 20;
  }

  if (account.LAST_VERIFIED_AT instanceof Date && !isNaN(account.LAST_VERIFIED_AT)) {
    score += 20;
  }
  if (account.LAST_REVIEW_DATE instanceof Date && !isNaN(account.LAST_REVIEW_DATE)) {
    score += 15;
  }
  if (account.LAST_LOGIN instanceof Date && !isNaN(account.LAST_LOGIN)) {
    score += 10;
  }
  if (normalizeText_(account.NOTES)) {
    score += 5;
  }
  return score;
}

function cloneRecord_(record, keys) {
  var copy = {};
  for (var i = 0; i < keys.length; i++) {
    var key = keys[i];
    copy[key] = record.hasOwnProperty(key) ? record[key] : '';
  }
  return copy;
}

function mergeDuplicateAccountRows_(rows, keeperRow) {
  var merged = cloneRecord_(keeperRow.data, GMP_SCHEMA.accounts.keys);
  for (var i = 0; i < rows.length; i++) {
    var account = rows[i].data;
    merged.ACCESS_ID = normalizeIdentifier_(merged.ACCESS_ID) || normalizeIdentifier_(account.ACCESS_ID);
    merged.SOURCE_REQUEST_ID = chooseNonBlankText_(merged.SOURCE_REQUEST_ID, account.SOURCE_REQUEST_ID);
    merged.GMP_SYSTEM = canonicalSystemName_(chooseNonBlankText_(merged.GMP_SYSTEM, account.GMP_SYSTEM));
    merged.PERSON = chooseLongerText_(merged.PERSON, account.PERSON);
    merged.COMPANY_EMAIL = chooseNonBlankText_(merged.COMPANY_EMAIL, account.COMPANY_EMAIL);
    merged.DEPT = chooseNonBlankText_(merged.DEPT, account.DEPT);
    merged.PLATFORM_ACCOUNT = chooseNonBlankText_(merged.PLATFORM_ACCOUNT, account.PLATFORM_ACCOUNT);
    merged.ACCESS_LEVEL = chooseNonBlankText_(merged.ACCESS_LEVEL, account.ACCESS_LEVEL);
    merged.CURRENT_STATE = choosePreferredAccountState_(merged.CURRENT_STATE, account.CURRENT_STATE);
    merged.LAST_VERIFIED_AT = chooseLatestDateValue_(merged.LAST_VERIFIED_AT, account.LAST_VERIFIED_AT);
    merged.ACTIVITY_BAND = choosePreferredActivityBand_(merged.ACTIVITY_BAND, account.ACTIVITY_BAND);
    merged.LAST_LOGIN = chooseLatestDateValue_(merged.LAST_LOGIN, account.LAST_LOGIN);
    merged.NEXT_REVIEW_DUE = chooseAccountDueDate_(merged.NEXT_REVIEW_DUE, merged.REVIEW_STATUS, account.NEXT_REVIEW_DUE, account.REVIEW_STATUS);
    merged.LAST_REVIEW_DATE = chooseLatestDateValue_(merged.LAST_REVIEW_DATE, account.LAST_REVIEW_DATE);
    merged.REVIEW_STATUS = choosePreferredReviewStatus_(merged.REVIEW_STATUS, account.REVIEW_STATUS);
    merged.OWNER = chooseNonBlankText_(merged.OWNER, account.OWNER);
    merged.NOTES = mergeNoteBlocks_(merged.NOTES, account.NOTES);
    merged.GRANTED_ON = chooseEarliestDateValue_(merged.GRANTED_ON, account.GRANTED_ON);
    merged.DAYS_SINCE_LOGIN = chooseNumericValue_(merged.DAYS_SINCE_LOGIN, account.DAYS_SINCE_LOGIN, true);
    merged.LOGINS_30D = chooseNumericValue_(merged.LOGINS_30D, account.LOGINS_30D, false);
    merged.LOGINS_90D = chooseNumericValue_(merged.LOGINS_90D, account.LOGINS_90D, false);
    merged.MFA = chooseNonBlankText_(merged.MFA, account.MFA);
    merged.PRIVILEGED = chooseNonBlankText_(merged.PRIVILEGED, account.PRIVILEGED);
    merged.ACTIVITY_SCORE = chooseNumericValue_(merged.ACTIVITY_SCORE, account.ACTIVITY_SCORE, false);
    merged.REVIEW_TRIGGER = chooseNonBlankText_(merged.REVIEW_TRIGGER, account.REVIEW_TRIGGER);
    merged.EXPECTED_USE = chooseNonBlankText_(merged.EXPECTED_USE, account.EXPECTED_USE);
    merged.MIN_LOGINS_30D = chooseNumericValue_(merged.MIN_LOGINS_30D, account.MIN_LOGINS_30D, false);
    merged.STALE_AFTER_DAYS = chooseNumericValue_(merged.STALE_AFTER_DAYS, account.STALE_AFTER_DAYS, false);
    merged.ACTIVITY_SOURCE = chooseNonBlankText_(merged.ACTIVITY_SOURCE, account.ACTIVITY_SOURCE);
    merged.ACTIVITY_CONFIDENCE = chooseNonBlankText_(merged.ACTIVITY_CONFIDENCE, account.ACTIVITY_CONFIDENCE);
    merged.SLACK_THREAD = chooseNonBlankText_(merged.SLACK_THREAD, account.SLACK_THREAD);
  }

  if (!normalizeIdentifier_(merged.ACCESS_ID)) {
    merged.ACCESS_ID = nextStableId_('ACC');
  }
  return merged;
}

function relinkRowsToMergedAccount_(mergedAccount, duplicateRows, keeperRowNumber) {
  var mergedAccessId = normalizeIdentifier_(mergedAccount.ACCESS_ID);
  if (!mergedAccessId) {
    return;
  }

  var duplicateAccessIds = {};
  for (var i = 0; i < duplicateRows.length; i++) {
    if (duplicateRows[i].rowNumber === keeperRowNumber) {
      continue;
    }
    var duplicateAccessId = normalizeIdentifier_(duplicateRows[i].data.ACCESS_ID);
    if (duplicateAccessId) {
      duplicateAccessIds[duplicateAccessId] = true;
    }
  }

  var reviewSignature = reviewSignatureKey_(mergedAccount.PERSON, mergedAccount.GMP_SYSTEM, mergedAccount.ACCESS_LEVEL);
  var reviewRows = getReviewRows_();
  for (var j = 0; j < reviewRows.length; j++) {
    var review = reviewRows[j].data;
    var reviewAccessId = normalizeIdentifier_(review.ACCESS_ID);
    var reviewFallbackIdentity = reviewSignatureKey_(review.PERSON, review.SYSTEM, review.ACCESS_LEVEL);
    if (duplicateAccessIds[reviewAccessId] || (!reviewAccessId && reviewFallbackIdentity === reviewSignature)) {
      review.ACCESS_ID = mergedAccessId;
      review.SYSTEM = canonicalSystemName_(mergedAccount.GMP_SYSTEM);
      review.PERSON = mergedAccount.PERSON;
      review.ACCESS_LEVEL = mergedAccount.ACCESS_LEVEL;
      updateReviewRow_(reviewRows[j].rowNumber, review);
    }
  }

  relinkRequestRowsToMergedAccount_(getRequestRows_(), mergedAccount, duplicateAccessIds);
  relinkRequestRowsToMergedAccount_(getArchivedRequestRows_(), mergedAccount, duplicateAccessIds);
}

function relinkRequestRowsToMergedAccount_(rows, mergedAccount, duplicateAccessIds) {
  var mergedAccessId = normalizeIdentifier_(mergedAccount.ACCESS_ID);
  var requestSignature = accountUniqueSignatureKey_(mergedAccount.COMPANY_EMAIL, mergedAccount.GMP_SYSTEM, mergedAccount.ACCESS_LEVEL, mergedAccount.PLATFORM_ACCOUNT, mergedAccount.PERSON);
  for (var i = 0; i < rows.length; i++) {
    var request = rows[i].data;
    var requestAccessId = normalizeIdentifier_(request.ACCESS_ID);
    var currentSignature = accountUniqueSignatureKey_(request.COMPANY_EMAIL, request.GMP_SYSTEM, request.ACCESS_LEVEL, request.COMPANY_EMAIL, request.TARGET_USER);
    if (!duplicateAccessIds[requestAccessId] && currentSignature !== requestSignature) {
      continue;
    }
    request.ACCESS_ID = mergedAccessId;
    request.GMP_SYSTEM = canonicalSystemName_(mergedAccount.GMP_SYSTEM);
    updateRequestRow_(rows[i].rowNumber, request, rows[i].sheetName);
  }
}

function chooseNonBlankText_(currentValue, incomingValue) {
  return normalizeText_(currentValue) || normalizeText_(incomingValue);
}

function chooseLongerText_(currentValue, incomingValue) {
  var current = normalizeText_(currentValue);
  var incoming = normalizeText_(incomingValue);
  if (!current) {
    return incoming;
  }
  if (!incoming) {
    return current;
  }
  return incoming.length > current.length ? incoming : current;
}

function chooseLatestDateValue_(currentValue, incomingValue) {
  var currentIsDate = currentValue instanceof Date && !isNaN(currentValue);
  var incomingIsDate = incomingValue instanceof Date && !isNaN(incomingValue);
  if (!currentIsDate) {
    return incomingIsDate ? incomingValue : currentValue;
  }
  if (!incomingIsDate) {
    return currentValue;
  }
  return incomingValue.getTime() > currentValue.getTime() ? incomingValue : currentValue;
}

function chooseEarliestDateValue_(currentValue, incomingValue) {
  var currentIsDate = currentValue instanceof Date && !isNaN(currentValue);
  var incomingIsDate = incomingValue instanceof Date && !isNaN(incomingValue);
  if (!currentIsDate) {
    return incomingIsDate ? incomingValue : currentValue;
  }
  if (!incomingIsDate) {
    return currentValue;
  }
  return incomingValue.getTime() < currentValue.getTime() ? incomingValue : currentValue;
}

function chooseNumericValue_(currentValue, incomingValue, preferLower) {
  var currentText = normalizeText_(currentValue);
  var incomingText = normalizeText_(incomingValue);
  if (currentText === '') {
    return incomingText === '' ? currentValue : incomingValue;
  }
  if (incomingText === '') {
    return currentValue;
  }

  var currentNumber = parseFloat(currentText);
  var incomingNumber = parseFloat(incomingText);
  if (isNaN(currentNumber) || isNaN(incomingNumber)) {
    return currentValue;
  }

  if (preferLower) {
    return incomingNumber < currentNumber ? incomingValue : currentValue;
  }
  return incomingNumber > currentNumber ? incomingValue : currentValue;
}

function choosePreferredAccountState_(currentState, incomingState) {
  var current = normalizePresenceState_(currentState);
  var incoming = normalizePresenceState_(incomingState);
  var rank = {
    'Present': 0,
    'Unmanaged': 1,
    'Missing': 2,
    'Setting Up': 3,
    'Revoked': 4,
    '': 5
  };
  return (rank.hasOwnProperty(incoming) ? rank[incoming] : 5) < (rank.hasOwnProperty(current) ? rank[current] : 5)
    ? incoming
    : current;
}

function choosePreferredActivityBand_(currentBand, incomingBand) {
  var current = normalizeActivityStatus_(currentBand);
  var incoming = normalizeActivityStatus_(incomingBand);
  var rank = {
    'Verified Active': 0,
    'Some Evidence': 1,
    'Unknown': 2,
    'No Evidence': 3,
    '': 4
  };
  return (rank.hasOwnProperty(incoming) ? rank[incoming] : 4) < (rank.hasOwnProperty(current) ? rank[current] : 4)
    ? incoming
    : current;
}

function choosePreferredReviewStatus_(currentStatus, incomingStatus) {
  var current = normalizeText_(currentStatus);
  var incoming = normalizeText_(incomingStatus);
  var rank = {
    'Open': 0,
    'Waiting': 1,
    'Done': 2,
    '': 3
  };
  return (rank.hasOwnProperty(incoming) ? rank[incoming] : 3) < (rank.hasOwnProperty(current) ? rank[current] : 3)
    ? incoming
    : current;
}

function chooseAccountDueDate_(currentDate, currentStatus, incomingDate, incomingStatus) {
  var current = currentDate instanceof Date && !isNaN(currentDate) ? currentDate : '';
  var incoming = incomingDate instanceof Date && !isNaN(incomingDate) ? incomingDate : '';
  if (!current) {
    return incoming || currentDate;
  }
  if (!incoming) {
    return current;
  }

  var currentWaiting = normalizeText_(currentStatus) === 'Waiting';
  var incomingWaiting = normalizeText_(incomingStatus) === 'Waiting';
  if (currentWaiting && !incomingWaiting) {
    return current;
  }
  if (!currentWaiting && incomingWaiting) {
    return incoming;
  }
  return incoming.getTime() < current.getTime() ? incoming : current;
}

function mergeNoteBlocks_(currentValue, incomingValue) {
  var merged = normalizeText_(currentValue);
  var incoming = normalizeText_(incomingValue);
  if (!incoming) {
    return merged;
  }
  var lines = incoming.split(/\r?\n/);
  for (var i = 0; i < lines.length; i++) {
    var line = normalizeText_(lines[i]);
    if (!line) {
      continue;
    }
    merged = appendUniqueNoteLine_(merged, line);
  }
  return merged;
}

function chooseReviewRowToKeep_(rows) {
  var bestRow = rows[0];
  var bestScore = reviewRowKeepScore_(rows[0]);
  for (var i = 1; i < rows.length; i++) {
    var score = reviewRowKeepScore_(rows[i]);
    if (score > bestScore || (score === bestScore && rows[i].rowNumber > bestRow.rowNumber)) {
      bestRow = rows[i];
      bestScore = score;
    }
  }
  return bestRow;
}

function reviewRowKeepScore_(row) {
  var review = row.data;
  var score = 0;
  var status = normalizeText_(review.DECISION_STATUS);
  var accountRow = findAccountRowForReview_(review);
  var accountStatus = accountRow ? normalizeText_(accountRow.data.REVIEW_STATUS) : '';

  if (accountStatus && status === accountStatus) {
    score += 500;
  }

  if (status === 'Done') {
    score += review.DECISION_DATE instanceof Date && !isNaN(review.DECISION_DATE) ? 260 : 240;
  } else if (status === 'Waiting') {
    score += review.NEXT_REVIEW_DUE instanceof Date && !isNaN(review.NEXT_REVIEW_DUE) ? 210 : 170;
  } else if (status === 'Open') {
    score += 160;
  }

  if (normalizeText_(review.REVIEWER_ACTION)) {
    score += 20;
  }
  if (normalizeText_(review.WHY_FLAGGED)) {
    score += 10;
  }
  if (normalizeText_(review.SLACK_THREAD)) {
    score += 5;
  }

  return score;
}

function removeSiblingReviewRowsForCycle_(accessId, reviewCycle, keepReviewId, person, systemName, accessLevel) {
  var reviewIdentity = reviewSignatureKey_(person, systemName, accessLevel) ||
    reviewAccountIdentityKey_(accessId, person, systemName, accessLevel);
  var normalizedReviewCycle = normalizeText_(reviewCycle);
  if (!reviewIdentity || !normalizedReviewCycle) {
    return;
  }

  var sheet = getRequiredSheet_(GMP_SCHEMA.tabs.REVIEW_CHECKS);
  var rows = getReviewRows_();
  var rowsToDelete = [];

  for (var i = 0; i < rows.length; i++) {
    var review = rows[i].data;
    if (reviewReusableIdentityKey_(review) !== reviewIdentity) {
      continue;
    }
    if (normalizeText_(review.REVIEW_CYCLE) !== normalizedReviewCycle) {
      continue;
    }
    if (normalizeText_(review.REVIEW_ID) === normalizeText_(keepReviewId)) {
      continue;
    }
    rowsToDelete.push(rows[i].rowNumber);
  }

  rowsToDelete.sort(function(a, b) { return b - a; });
  for (var j = 0; j < rowsToDelete.length; j++) {
    sheet.deleteRow(rowsToDelete[j]);
  }
}

function syncAccountReviewRollupsFromReviews_() {
  var accountRows = getAccountRows_();
  var reviewRows = getReviewRows_();
  var reviewsByIdentity = {};

  for (var i = 0; i < reviewRows.length; i++) {
    var identityKey = reviewReusableIdentityKey_(reviewRows[i].data);
    if (!identityKey) {
      continue;
    }
    if (!reviewsByIdentity[identityKey]) {
      reviewsByIdentity[identityKey] = [];
    }
    reviewsByIdentity[identityKey].push(reviewRows[i]);
  }

  for (var j = 0; j < accountRows.length; j++) {
    var account = accountRows[j].data;
    var accountIdentity = accountReusableReviewIdentityKey_(account);
    if (!accountIdentity || !reviewsByIdentity[accountIdentity]) {
      if (normalizeText_(account.REVIEW_STATUS) !== '') {
        account.REVIEW_STATUS = '';
        updateAccountRow_(accountRows[j].rowNumber, account);
      }
      continue;
    }

    var relatedReviews = reviewsByIdentity[accountIdentity];
    var currentReview = chooseCurrentReviewRowForAccount_(relatedReviews);
    var latestDoneReview = chooseLatestDoneReviewRowForAccount_(relatedReviews);
    var shouldUpdate = false;

    if (currentReview) {
      var currentStatus = normalizeText_(currentReview.data.DECISION_STATUS);
      if (normalizeText_(account.REVIEW_STATUS) !== currentStatus) {
        account.REVIEW_STATUS = currentStatus;
        shouldUpdate = true;
      }
      if (currentReview.data.NEXT_REVIEW_DUE instanceof Date &&
          !isNaN(currentReview.data.NEXT_REVIEW_DUE) &&
          currentReview.data.NEXT_REVIEW_DUE instanceof Date &&
          !sameDateValue_(account.NEXT_REVIEW_DUE, currentReview.data.NEXT_REVIEW_DUE)) {
        account.NEXT_REVIEW_DUE = currentReview.data.NEXT_REVIEW_DUE;
        shouldUpdate = true;
      }
    } else if (latestDoneReview) {
      if (normalizeText_(account.REVIEW_STATUS) !== 'Done') {
        account.REVIEW_STATUS = 'Done';
        shouldUpdate = true;
      }
    }

    if (latestDoneReview && latestDoneReview.data.DECISION_DATE instanceof Date && !isNaN(latestDoneReview.data.DECISION_DATE)) {
      if (!sameDateValue_(account.LAST_REVIEW_DATE, latestDoneReview.data.DECISION_DATE)) {
        account.LAST_REVIEW_DATE = latestDoneReview.data.DECISION_DATE;
        shouldUpdate = true;
      }

      var desiredNextReviewDue = deriveNextReviewDueFromDoneReview_(latestDoneReview.data);
      if (desiredNextReviewDue &&
          (!(account.NEXT_REVIEW_DUE instanceof Date) ||
           isNaN(account.NEXT_REVIEW_DUE) ||
           account.NEXT_REVIEW_DUE.getTime() <= latestDoneReview.data.DECISION_DATE.getTime())) {
        account.NEXT_REVIEW_DUE = desiredNextReviewDue;
        shouldUpdate = true;
      }
    }

    if (shouldUpdate) {
      updateAccountRow_(accountRows[j].rowNumber, account);
    }
  }
}

function chooseCurrentReviewRowForAccount_(rows) {
  var best = null;
  for (var i = 0; i < rows.length; i++) {
    var status = normalizeText_(rows[i].data.DECISION_STATUS);
    if (status === 'Done') {
      continue;
    }
    if (!best) {
      best = rows[i];
      continue;
    }

    var bestStatusRank = reviewStatusRank_(best.data.DECISION_STATUS);
    var rowStatusRank = reviewStatusRank_(rows[i].data.DECISION_STATUS);
    if (rowStatusRank < bestStatusRank) {
      best = rows[i];
      continue;
    }
    if (rowStatusRank === bestStatusRank) {
      var dateCompare = compareByDateDesc_(rows[i].data.NEXT_REVIEW_DUE, best.data.NEXT_REVIEW_DUE);
      if (dateCompare < 0) {
        best = rows[i];
      } else if (dateCompare === 0 && rows[i].rowNumber > best.rowNumber) {
        best = rows[i];
      }
    }
  }
  return best;
}

function chooseLatestDoneReviewRowForAccount_(rows) {
  var best = null;
  for (var i = 0; i < rows.length; i++) {
    if (normalizeText_(rows[i].data.DECISION_STATUS) !== 'Done') {
      continue;
    }
    if (!best) {
      best = rows[i];
      continue;
    }
    var dateCompare = compareByDateDesc_(rows[i].data.DECISION_DATE, best.data.DECISION_DATE);
    if (dateCompare < 0 || (dateCompare === 0 && rows[i].rowNumber > best.rowNumber)) {
      best = rows[i];
    }
  }
  return best;
}

function deriveNextReviewDueFromDoneReview_(review) {
  return nextReviewDueFromDecision_(review.DECISION_DATE, review.SYSTEM, review.ACCESS_LEVEL, review.REVIEWER_ACTION);
}

function sameDateValue_(left, right) {
  if (!(left instanceof Date) || isNaN(left)) {
    return !(right instanceof Date) || isNaN(right);
  }
  if (!(right instanceof Date) || isNaN(right)) {
    return false;
  }
  return left.getTime() === right.getTime();
}

function normalizeRequestTerminology_(rows) {
  for (var i = 0; i < rows.length; i++) {
    var record = rows[i].data;
    var normalizedStatus = normalizeSetupStatus_(record.PROVISIONING);
    if (normalizedStatus && normalizedStatus !== normalizeText_(record.PROVISIONING)) {
      record.PROVISIONING = normalizedStatus;
      updateRequestRow_(rows[i].rowNumber, record, rows[i].sheetName);
    }
  }
}

function normalizeAccountTerminology_(rows) {
  for (var i = 0; i < rows.length; i++) {
    var record = rows[i].data;
    var normalizedState = normalizePresenceState_(record.CURRENT_STATE);
    var compactNotes = compactAuditLog_(record.NOTES);
    var shouldUpdate = false;
    if (normalizedState && normalizedState !== normalizeText_(record.CURRENT_STATE)) {
      record.CURRENT_STATE = normalizedState;
      shouldUpdate = true;
    }
    if (compactNotes !== normalizeText_(record.NOTES)) {
      record.NOTES = compactNotes;
      shouldUpdate = true;
    }
    if (shouldUpdate) {
      updateAccountRow_(rows[i].rowNumber, record);
    }
  }
}

function getAccessControlContext_() {
  var sheet = getRequiredSheet_(GMP_SCHEMA.tabs.ACCESS_CONTROL);
  var requestSectionRow = findLabelRow_(sheet, GMP_SCHEMA.sectionLabels.REQUESTS);
  var accountSectionRow = findLabelRow_(sheet, GMP_SCHEMA.sectionLabels.ACCOUNTS);
  var archiveSectionRow = findLabelRow_(sheet, GMP_SCHEMA.sectionLabels.ARCHIVED_REQUESTS);
  if (requestSectionRow === -1 || accountSectionRow === -1) {
    throw new Error('Access Control sheet sections are missing.');
  }
  return {
    sheet: sheet,
    requestSectionRow: requestSectionRow,
    requestHeaderRow: requestSectionRow + 1,
    accountSectionRow: accountSectionRow,
    accountHeaderRow: accountSectionRow + 1,
    archiveSectionRow: archiveSectionRow,
    archiveHeaderRow: archiveSectionRow === -1 ? -1 : archiveSectionRow + 1
  };
}

function getRequestRows_() {
  var ctx = getAccessControlContext_();
  var startRow = ctx.requestHeaderRow + 1;
  var endRow = ctx.accountSectionRow - 1;
  if (endRow < startRow) {
    return [];
  }
  var values = ctx.sheet.getRange(startRow, 1, endRow - startRow + 1, GMP_SCHEMA.requests.keys.length).getValues();
  return buildTableRows_(values, startRow, GMP_SCHEMA.requests.keys, GMP_SCHEMA.requests.columns, GMP_SCHEMA.tabs.ACCESS_CONTROL);
}

function ensureArchivedRequestsSection_() {
  var ctx = getAccessControlContext_();
  if (ctx.archiveSectionRow !== -1) {
    return ctx;
  }

  var sheet = ctx.sheet;
  var width = Math.max(GMP_SCHEMA.requests.headers.length, GMP_SCHEMA.accounts.headers.length);
  var sectionRow = Math.max(sheet.getLastRow(), ctx.accountHeaderRow + 1) + 1;
  ensureSheetHasRows_(sheet, sectionRow + 1);

  writeSectionRow_(sheet, sectionRow, GMP_SCHEMA.sectionLabels.ARCHIVED_REQUESTS, width, GMP_SHEET_COLORS.REQUEST_SECTION);
  writeHeaderRow_(sheet, sectionRow + 1, GMP_SCHEMA.requests.headers, GMP_SHEET_COLORS.HEADER);
  applyStandardRowHeights_(sheet, [sectionRow, sectionRow + 1]);
  applyDropdowns_(sheet, sectionRow + 2, Math.max(sheet.getMaxRows() - (sectionRow + 1), 1), GMP_SCHEMA.requests);
  applyDateFormatsByKeys_(sheet, GMP_SCHEMA.requests.columns, ['SUBMITTED_AT', 'SLA_DUE', 'DECISION_DATE'], sectionRow + 2, Math.max(sheet.getMaxRows() - (sectionRow + 1), 1), 'yyyy-mm-dd hh:mm');

  return getAccessControlContext_();
}

function getArchivedRequestRows_() {
  var ctx = getAccessControlContext_();
  if (ctx.archiveSectionRow === -1) {
    return [];
  }
  var startRow = ctx.archiveHeaderRow + 1;
  var endRow = ctx.sheet.getLastRow();
  if (endRow < startRow) {
    return [];
  }
  var values = ctx.sheet.getRange(startRow, 1, endRow - startRow + 1, GMP_SCHEMA.requests.keys.length).getValues();
  return buildTableRows_(values, startRow, GMP_SCHEMA.requests.keys, GMP_SCHEMA.requests.columns, GMP_SCHEMA.tabs.ACCESS_CONTROL);
}

function getAccountRows_() {
  var ctx = getAccessControlContext_();
  var startRow = ctx.accountHeaderRow + 1;
  var endRow = ctx.archiveSectionRow === -1 ? ctx.sheet.getLastRow() : ctx.archiveSectionRow - 1;
  if (endRow < startRow) {
    return [];
  }
  var values = ctx.sheet.getRange(startRow, 1, endRow - startRow + 1, GMP_SCHEMA.accounts.keys.length).getValues();
  return buildTableRows_(values, startRow, GMP_SCHEMA.accounts.keys, GMP_SCHEMA.accounts.columns, GMP_SCHEMA.tabs.ACCESS_CONTROL);
}

function getReviewRows_() {
  var sheet = getRequiredSheet_(GMP_SCHEMA.tabs.REVIEW_CHECKS);
  var startRow = GMP_SCHEMA.layout.REVIEW_CONTROL_ROW + 1;
  var endRow = sheet.getLastRow();
  if (endRow < startRow) {
    return [];
  }
  var values = sheet.getRange(startRow, 1, endRow - startRow + 1, GMP_SCHEMA.reviews.keys.length).getValues();
  return buildTableRows_(values, startRow, GMP_SCHEMA.reviews.keys, GMP_SCHEMA.reviews.columns, GMP_SCHEMA.tabs.REVIEW_CHECKS);
}

function getIncidentRows_() {
  var sheet = getRequiredSheet_(GMP_SCHEMA.tabs.INCIDENTS);
  var startRow = GMP_SCHEMA.layout.INCIDENT_HEADER_ROW + 1;
  var endRow = sheet.getLastRow();
  if (endRow < startRow) {
    return [];
  }
  var values = sheet.getRange(startRow, 1, endRow - startRow + 1, GMP_SCHEMA.incidents.keys.length).getValues();
  return buildTableRows_(values, startRow, GMP_SCHEMA.incidents.keys, GMP_SCHEMA.incidents.columns, GMP_SCHEMA.tabs.INCIDENTS);
}

function buildTableRows_(values, startRow, keys, columnMap, sheetName) {
  var rows = [];
  for (var i = 0; i < values.length; i++) {
    if (isBlankRow_(values[i]) || isDisplayGroupRow_(values[i])) {
      continue;
    }
    rows.push({
      rowNumber: startRow + i,
      sheetName: sheetName || '',
      values: values[i],
      data: rowToObject_(values[i], keys, columnMap)
    });
  }
  return rows;
}

function appendAccessRequestRecord_(record) {
  ensureThreeTabWorkbookReady_();
  var ctx = getAccessControlContext_();
  var targetRow = ctx.accountSectionRow;
  ctx.sheet.insertRowsBefore(targetRow, 1);
  var targetRange = ctx.sheet.getRange(targetRow, 1, 1, GMP_SCHEMA.requests.keys.length);
  replaceRangeValues_(targetRange, [objectToRow_(record, GMP_SCHEMA.requests.keys)]);
  applyOperationalRowFormatting_(targetRange);
  applyDropdowns_(ctx.sheet, targetRow, 1, GMP_SCHEMA.requests);
  applyDateFormatsByKeys_(ctx.sheet, GMP_SCHEMA.requests.columns, ['SUBMITTED_AT', 'SLA_DUE', 'DECISION_DATE'], targetRow, 1, 'yyyy-mm-dd hh:mm');
  sortRequestSection_();
  var requestRow = findRequestRowById_(record.REQUEST_ID);
  return requestRow ? requestRow.rowNumber : targetRow;
}

function appendAccountRecord_(record) {
  ensureThreeTabWorkbookReady_();
  record.NOTES = compactAuditLog_(record.NOTES);
  var ctx = getAccessControlContext_();
  var targetRow = ctx.archiveSectionRow === -1 ? Math.max(ctx.accountHeaderRow + 1, ctx.sheet.getLastRow() + 1) : ctx.archiveSectionRow;
  if (ctx.archiveSectionRow !== -1) {
    ctx.sheet.insertRowsBefore(targetRow, 1);
  } else if (targetRow > ctx.sheet.getMaxRows()) {
    ctx.sheet.insertRowsAfter(ctx.sheet.getMaxRows(), 1);
  }
  var targetRange = ctx.sheet.getRange(targetRow, 1, 1, GMP_SCHEMA.accounts.keys.length);
  replaceRangeValues_(targetRange, [objectToRow_(record, GMP_SCHEMA.accounts.keys)]);
  applyOperationalRowFormatting_(targetRange);
  applyDropdowns_(ctx.sheet, targetRow, 1, GMP_SCHEMA.accounts);
  applyDateFormatsByKeys_(ctx.sheet, GMP_SCHEMA.accounts.columns, ['LAST_VERIFIED_AT', 'GRANTED_ON', 'LAST_LOGIN'], targetRow, 1, 'yyyy-mm-dd hh:mm');
  applyDateFormatsByKeys_(ctx.sheet, GMP_SCHEMA.accounts.columns, ['NEXT_REVIEW_DUE', 'LAST_REVIEW_DATE'], targetRow, 1, 'yyyy-mm-dd');
  sortAccountSection_();
  var accountRow = findAccountRowByAccessId_(record.ACCESS_ID);
  return accountRow ? accountRow.rowNumber : targetRow;
}

function appendReviewRecord_(record) {
  ensureThreeTabWorkbookReady_();
  var sheet = getRequiredSheet_(GMP_SCHEMA.tabs.REVIEW_CHECKS);
  var targetRow = Math.max(GMP_SCHEMA.layout.REVIEW_CONTROL_ROW + 1, sheet.getLastRow() + 1);
  if (targetRow > sheet.getMaxRows()) {
    sheet.insertRowsAfter(sheet.getMaxRows(), 1);
  }
  var targetRange = sheet.getRange(targetRow, 1, 1, GMP_SCHEMA.reviews.keys.length);
  replaceRangeValues_(targetRange, [objectToRow_(record, GMP_SCHEMA.reviews.keys)]);
  applyOperationalRowFormatting_(targetRange);
  applyDropdowns_(sheet, targetRow, 1, GMP_SCHEMA.reviews);
  applyDateFormatsByKeys_(sheet, GMP_SCHEMA.reviews.columns, ['LAST_LOGIN', 'DECISION_DATE'], targetRow, 1, 'yyyy-mm-dd hh:mm');
  applyDateFormatsByKeys_(sheet, GMP_SCHEMA.reviews.columns, ['NEXT_REVIEW_DUE'], targetRow, 1, 'yyyy-mm-dd');
  sortReviewSection_();
  var reviewRow = findReviewRowById_(record.REVIEW_ID);
  return reviewRow ? reviewRow.rowNumber : targetRow;
}

function appendIncidentRecord_(record) {
  ensureThreeTabWorkbookReady_();
  record.NOTES = compactIncidentNotes_(record.NOTES);
  var sheet = getRequiredSheet_(GMP_SCHEMA.tabs.INCIDENTS);
  var targetRow = Math.max(GMP_SCHEMA.layout.INCIDENT_HEADER_ROW + 1, sheet.getLastRow() + 1);
  if (targetRow > sheet.getMaxRows()) {
    sheet.insertRowsAfter(sheet.getMaxRows(), 1);
  }
  var targetRange = sheet.getRange(targetRow, 1, 1, GMP_SCHEMA.incidents.keys.length);
  replaceRangeValues_(targetRange, [objectToRow_(record, GMP_SCHEMA.incidents.keys)]);
  applyOperationalRowFormatting_(targetRange);
  applyDropdowns_(sheet, targetRow, 1, GMP_SCHEMA.incidents);
  applyDateFormatsByKeys_(sheet, GMP_SCHEMA.incidents.columns, ['REPORTED_AT', 'RESOLVED_AT'], targetRow, 1, 'yyyy-mm-dd hh:mm');
  sortIncidentSection_();
  var incidentRow = findIncidentRowById_(record.INCIDENT_ID);
  return incidentRow ? incidentRow.rowNumber : targetRow;
}

function updateRequestRow_(rowNumber, record, sheetName) {
  var sheet = getRequiredSheet_(sheetName || GMP_SCHEMA.tabs.ACCESS_CONTROL);
  var targetRange = sheet.getRange(rowNumber, 1, 1, GMP_SCHEMA.requests.keys.length);
  replaceRangeValues_(targetRange, [objectToRow_(record, GMP_SCHEMA.requests.keys)]);
  applyOperationalRowFormatting_(targetRange);
  applyDropdowns_(sheet, rowNumber, 1, GMP_SCHEMA.requests);
  applyDateFormatsByKeys_(sheet, GMP_SCHEMA.requests.columns, ['SUBMITTED_AT', 'SLA_DUE', 'DECISION_DATE'], rowNumber, 1, 'yyyy-mm-dd hh:mm');
}

function updateAccountRow_(rowNumber, record) {
  var sheet = getRequiredSheet_(GMP_SCHEMA.tabs.ACCESS_CONTROL);
  record.NOTES = compactAuditLog_(record.NOTES);
  var targetRange = sheet.getRange(rowNumber, 1, 1, GMP_SCHEMA.accounts.keys.length);
  replaceRangeValues_(targetRange, [objectToRow_(record, GMP_SCHEMA.accounts.keys)]);
  applyOperationalRowFormatting_(targetRange);
  applyDropdowns_(sheet, rowNumber, 1, GMP_SCHEMA.accounts);
  applyDateFormatsByKeys_(sheet, GMP_SCHEMA.accounts.columns, ['LAST_VERIFIED_AT', 'GRANTED_ON', 'LAST_LOGIN'], rowNumber, 1, 'yyyy-mm-dd hh:mm');
  applyDateFormatsByKeys_(sheet, GMP_SCHEMA.accounts.columns, ['NEXT_REVIEW_DUE', 'LAST_REVIEW_DATE'], rowNumber, 1, 'yyyy-mm-dd');
}

function updateReviewRow_(rowNumber, record) {
  var sheet = getRequiredSheet_(GMP_SCHEMA.tabs.REVIEW_CHECKS);
  var targetRange = sheet.getRange(rowNumber, 1, 1, GMP_SCHEMA.reviews.keys.length);
  replaceRangeValues_(targetRange, [objectToRow_(record, GMP_SCHEMA.reviews.keys)]);
  applyOperationalRowFormatting_(targetRange);
  applyDropdowns_(sheet, rowNumber, 1, GMP_SCHEMA.reviews);
  applyDateFormatsByKeys_(sheet, GMP_SCHEMA.reviews.columns, ['LAST_LOGIN', 'DECISION_DATE'], rowNumber, 1, 'yyyy-mm-dd hh:mm');
  applyDateFormatsByKeys_(sheet, GMP_SCHEMA.reviews.columns, ['NEXT_REVIEW_DUE'], rowNumber, 1, 'yyyy-mm-dd');
}

function updateIncidentRow_(rowNumber, record) {
  var sheet = getRequiredSheet_(GMP_SCHEMA.tabs.INCIDENTS);
  record.NOTES = compactIncidentNotes_(record.NOTES);
  var targetRange = sheet.getRange(rowNumber, 1, 1, GMP_SCHEMA.incidents.keys.length);
  replaceRangeValues_(targetRange, [objectToRow_(record, GMP_SCHEMA.incidents.keys)]);
  applyOperationalRowFormatting_(targetRange);
  applyDropdowns_(sheet, rowNumber, 1, GMP_SCHEMA.incidents);
  applyDateFormatsByKeys_(sheet, GMP_SCHEMA.incidents.columns, ['REPORTED_AT', 'RESOLVED_AT'], rowNumber, 1, 'yyyy-mm-dd hh:mm');
}

function findRequestRowById_(requestId) {
  return findActiveRequestRowById_(requestId) || findArchivedRequestRowById_(requestId);
}

function findActiveRequestRowById_(requestId) {
  return findRowByValue_(getRequestRows_(), 'REQUEST_ID', requestId);
}

function findArchivedRequestRowById_(requestId) {
  return findRowByValue_(getArchivedRequestRows_(), 'REQUEST_ID', requestId);
}

function archiveActiveRequestById_(requestId) {
  var requestRow = findActiveRequestRowById_(requestId);
  if (!requestRow) {
    return findArchivedRequestRowById_(requestId);
  }

  return archiveRequestRow_(requestRow);
}

function archiveRequestRow_(requestRow) {
  if (!requestRow || !requestRow.data || !normalizeText_(requestRow.data.REQUEST_ID)) {
    return null;
  }

  upsertArchivedRequestRecord_(requestRow.data);
  var activeRow = findActiveRequestRowById_(requestRow.data.REQUEST_ID);
  if (activeRow) {
    getRequiredSheet_(activeRow.sheetName || GMP_SCHEMA.tabs.ACCESS_CONTROL).deleteRow(activeRow.rowNumber);
  }
  return findArchivedRequestRowById_(requestRow.data.REQUEST_ID);
}

function upsertArchivedRequestRecord_(record) {
  var ctx = ensureArchivedRequestsSection_();
  var sheet = ctx.sheet;
  var existingRow = findArchivedRequestRowById_(record.REQUEST_ID);
  var targetRow = existingRow ? existingRow.rowNumber : ctx.archiveHeaderRow + 1;

  if (!existingRow) {
    var firstDataRow = ctx.archiveHeaderRow + 1;
    if (!isBlankRow_(sheet.getRange(firstDataRow, 1, 1, GMP_SCHEMA.requests.keys.length).getValues()[0])) {
      sheet.insertRowsBefore(firstDataRow, 1);
    }
    ensureSheetHasRows_(sheet, targetRow);
  }

  var targetRange = sheet.getRange(targetRow, 1, 1, GMP_SCHEMA.requests.keys.length);
  replaceRangeValues_(targetRange, [objectToRow_(record, GMP_SCHEMA.requests.keys)]);
  applyOperationalRowFormatting_(targetRange);
  applyDropdowns_(sheet, targetRow, 1, GMP_SCHEMA.requests);
  applyDateFormatsByKeys_(sheet, GMP_SCHEMA.requests.columns, ['SUBMITTED_AT', 'SLA_DUE', 'DECISION_DATE'], targetRow, 1, 'yyyy-mm-dd hh:mm');
  sortArchivedRequestSection_();
}

function findAccountRowByAccessId_(accessId) {
  var normalizedAccessId = normalizeIdentifier_(accessId);
  if (!normalizedAccessId) {
    return null;
  }

  var rows = getAccountRows_();
  for (var i = 0; i < rows.length; i++) {
    if (normalizeIdentifier_(rows[i].data.ACCESS_ID) === normalizedAccessId) {
      return rows[i];
    }
  }
  return null;
}

function findAccountRowForReview_(review) {
  var directMatch = findAccountRowByAccessId_(review.ACCESS_ID);
  if (directMatch) {
    return directMatch;
  }

  var identityKey = reviewReusableIdentityKey_(review);
  var accountRows = getAccountRows_();
  for (var i = 0; i < accountRows.length; i++) {
    if (accountReusableReviewIdentityKey_(accountRows[i].data) === identityKey) {
      return accountRows[i];
    }
  }
  return null;
}

function findReviewRowById_(reviewId) {
  return findRowByValue_(getReviewRows_(), 'REVIEW_ID', reviewId);
}

function findIncidentRowById_(incidentId) {
  return findRowByValue_(getIncidentRows_(), 'INCIDENT_ID', incidentId);
}

function findAccountRowsByEmailAndSystem_(email, systemName) {
  var rows = getAccountRows_();
  var normalizedEmail = normalizeText_(email).toLowerCase();
  var normalizedSystem = canonicalSystemName_(systemName).toLowerCase();
  var matches = [];
  for (var i = 0; i < rows.length; i++) {
    var data = rows[i].data;
    if (normalizeText_(data.COMPANY_EMAIL).toLowerCase() === normalizedEmail &&
        canonicalSystemName_(data.GMP_SYSTEM).toLowerCase() === normalizedSystem) {
      matches.push(rows[i]);
    }
  }
  return matches;
}

function findAccountRowByEmailAndSystem_(email, systemName, accessLevel) {
  var matches = findAccountRowsByEmailAndSystem_(email, systemName);
  if (arguments.length < 3 || normalizeText_(accessLevel) === '') {
    return matches.length ? matches[0] : null;
  }

  var normalizedAccessLevel = normalizeText_(accessLevel).toLowerCase();
  for (var i = 0; i < matches.length; i++) {
    if (normalizeText_(matches[i].data.ACCESS_LEVEL).toLowerCase() === normalizedAccessLevel) {
      return matches[i];
    }
  }
  return null;
}

function findBestAccountRowForRequest_(request) {
  var accessId = normalizeIdentifier_(request.ACCESS_ID);
  if (accessId) {
    var directMatch = findAccountRowByAccessId_(accessId);
    if (directMatch) {
      return directMatch;
    }
  }

  var canonicalSystem = canonicalSystemName_(request.GMP_SYSTEM);
  var exactMatch = findAccountRowByEmailAndSystem_(request.COMPANY_EMAIL, canonicalSystem, request.ACCESS_LEVEL);
  if (exactMatch) {
    return exactMatch;
  }

  var requestType = normalizeText_(request.REQUEST_TYPE);
  if (requestType !== 'Change Access' && requestType !== 'Remove Access') {
    return null;
  }

  var systemMatches = findAccountRowsByEmailAndSystem_(request.COMPANY_EMAIL, canonicalSystem);
  if (systemMatches.length === 1) {
    return systemMatches[0];
  }

  var normalizedTargetUser = normalizeText_(request.TARGET_USER).toLowerCase();
  var personMatches = [];
  for (var i = 0; i < systemMatches.length; i++) {
    if (normalizeText_(systemMatches[i].data.PERSON).toLowerCase() === normalizedTargetUser) {
      personMatches.push(systemMatches[i]);
    }
  }

  return personMatches.length === 1 ? personMatches[0] : null;
}

function findOpenReviewRowForAccess_(accessId, triggerType) {
  var rows = getReviewRows_();
  for (var i = 0; i < rows.length; i++) {
    var data = rows[i].data;
    if (normalizeText_(data.ACCESS_ID) !== normalizeText_(accessId)) {
      continue;
    }
    if (normalizeText_(data.TRIGGER_TYPE) !== normalizeText_(triggerType)) {
      continue;
    }
    if (normalizeText_(data.DECISION_STATUS) === 'Done') {
      continue;
    }
    return rows[i];
  }
  return null;
}

function findRowByValue_(rows, key, expectedValue) {
  var needle = normalizeText_(expectedValue);
  if (!needle) {
    return null;
  }
  for (var i = 0; i < rows.length; i++) {
    if (normalizeText_(rows[i].data[key]) === needle) {
      return rows[i];
    }
  }
  return null;
}

function getRowLink_(sheetName, rowNumber) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    return '';
  }
  return 'https://docs.google.com/spreadsheets/d/' + ss.getId() +
    '/edit#gid=' + sheet.getSheetId() + '&range=A' + rowNumber;
}

function appendNoteLine_(existingValue, noteLine) {
  var current = normalizeText_(existingValue);
  return current ? current + '\n' + noteLine : noteLine;
}

function compactAuditLogLine_(value) {
  var line = normalizeText_(value);
  if (!line) {
    return '';
  }

  line = line.replace(': discovered in Katana API without matching request ', ': Katana no request ');
  line = line.replace(': verified in Katana API ', ': Katana verified ');
  line = line.replace(': not found in latest Katana API inventory', ': Katana missing');
  line = line.replace(': synced from ', ': synced ');
  line = line.replace(' (setup in progress)', ' setup');
  line = line.replace(': revoked via ', ': revoked ');
  line = line.replace(': refreshed from account activity', ': review refreshed');
  line = line.replace(': review action Reduce', ': reduce access');
  line = line.replace(': Katana API verified account ', ': Katana verified ');
  line = line.replace('Initial Slack notification failed: ', 'Slack failed: ');

  return line;
}

function compactAuditLog_(value) {
  var current = normalizeText_(value);
  if (!current) {
    return '';
  }

  var lines = current.split(/\r?\n+/);
  var compacted = [];
  var seen = {};
  for (var i = 0; i < lines.length; i++) {
    var line = compactAuditLogLine_(lines[i]);
    if (!line) {
      continue;
    }
    if (seen[line]) {
      continue;
    }
    seen[line] = true;
    compacted.push(line);
  }

  if (compacted.length > 2) {
    compacted = compacted.slice(compacted.length - 2);
  }

  return compacted.join(' | ');
}

function compactIncidentNoteLine_(value) {
  var line = normalizeText_(value);
  if (!line) {
    return '';
  }

  line = line.replace(/^Created from /, 'Source: ');
  line = line.replace(/^Still happening: /, 'Ongoing: ');
  line = line.replace(/^First noticed: /, 'First: ');
  line = line.replace(/^Details: /, 'Details: ');

  if (line.length > 120) {
    line = line.substring(0, 117) + '...';
  }

  return line;
}

function compactIncidentNotes_(value) {
  var current = normalizeText_(value);
  if (!current) {
    return '';
  }

  var parts = current.split(/\r?\n+/);
  var compacted = [];
  var seen = {};
  for (var i = 0; i < parts.length; i++) {
    var line = compactIncidentNoteLine_(parts[i]);
    if (!line || seen[line]) {
      continue;
    }
    seen[line] = true;
    compacted.push(line);
  }

  return compacted.join(' | ');
}

function applyOperationalRowFormatting_(range) {
  if (!range) {
    return;
  }

  range
    .setBackground('#ffffff')
    .setFontColor('#0f172a')
    .setFontWeight('normal')
    .setWrap(false)
    .setHorizontalAlignment('left')
    .setVerticalAlignment('middle');

  var sheet = range.getSheet();
  sheet.setRowHeights(range.getRow(), range.getNumRows(), 20);
}

function sortRequestSection_() {
  var ctx = getAccessControlContext_();
  var startRow = ctx.requestHeaderRow + 1;
  var endRow = ctx.accountSectionRow - 1;
  sortTableSection_(
    ctx.sheet,
    startRow,
    endRow,
    GMP_SCHEMA.requests.keys,
    GMP_SCHEMA.requests.columns,
    compareRequestRows_
  );
  if (endRow >= startRow) {
    var range = ctx.sheet.getRange(startRow, 1, endRow - startRow + 1, GMP_SCHEMA.requests.keys.length);
    applyOperationalRowFormatting_(range);
    applyDropdowns_(ctx.sheet, startRow, endRow - startRow + 1, GMP_SCHEMA.requests);
    applyDateFormatsByKeys_(ctx.sheet, GMP_SCHEMA.requests.columns, ['SUBMITTED_AT', 'SLA_DUE', 'DECISION_DATE'], startRow, endRow - startRow + 1, 'yyyy-mm-dd hh:mm');
  }
}

function sortArchivedRequestSection_() {
  var ctx = getAccessControlContext_();
  if (ctx.archiveSectionRow === -1) {
    return;
  }

  var startRow = ctx.archiveHeaderRow + 1;
  var endRow = ctx.sheet.getLastRow();
  sortTableSection_(
    ctx.sheet,
    startRow,
    endRow,
    GMP_SCHEMA.requests.keys,
    GMP_SCHEMA.requests.columns,
    compareRequestRows_
  );
  if (endRow >= startRow) {
    var range = ctx.sheet.getRange(startRow, 1, endRow - startRow + 1, GMP_SCHEMA.requests.keys.length);
    applyOperationalRowFormatting_(range);
    applyDropdowns_(ctx.sheet, startRow, endRow - startRow + 1, GMP_SCHEMA.requests);
    applyDateFormatsByKeys_(ctx.sheet, GMP_SCHEMA.requests.columns, ['SUBMITTED_AT', 'SLA_DUE', 'DECISION_DATE'], startRow, endRow - startRow + 1, 'yyyy-mm-dd hh:mm');
  }
}

function sortAccountSection_() {
  var ctx = getAccessControlContext_();
  var startRow = ctx.accountHeaderRow + 1;
  var endRow = ctx.archiveSectionRow === -1 ? Math.max(ctx.sheet.getLastRow(), startRow - 1) : Math.max(ctx.archiveSectionRow - 1, startRow - 1);
  if (endRow < startRow) {
    return;
  }

  var values = ctx.sheet.getRange(startRow, 1, endRow - startRow + 1, GMP_SCHEMA.accounts.keys.length).getValues();
  var accounts = [];
  for (var i = 0; i < values.length; i++) {
    if (isBlankRow_(values[i]) || isDisplayGroupRow_(values[i])) {
      continue;
    }
    accounts.push(rowToObject_(values[i], GMP_SCHEMA.accounts.keys, GMP_SCHEMA.accounts.columns));
  }

  accounts.sort(compareAccountRows_);
  var renderedRows = buildRenderedAccountRows_(accounts);
  var requiredRows = Math.max(renderedRows.length, endRow - startRow + 1);
  ensureSheetHasRows_(ctx.sheet, startRow + requiredRows - 1);

  var output = renderedRows.slice();
  while (output.length < requiredRows) {
    output.push(buildBlankRow_(GMP_SCHEMA.accounts.keys.length));
  }

  var targetRange = ctx.sheet.getRange(startRow, 1, requiredRows, GMP_SCHEMA.accounts.keys.length);
  replaceRangeValues_(targetRange, output);
  applyOperationalRowFormatting_(targetRange);
  applyDropdowns_(ctx.sheet, startRow, requiredRows, GMP_SCHEMA.accounts);
  applyDateFormatsByKeys_(ctx.sheet, GMP_SCHEMA.accounts.columns, ['LAST_VERIFIED_AT', 'GRANTED_ON', 'LAST_LOGIN'], startRow, requiredRows, 'yyyy-mm-dd hh:mm');
  applyDateFormatsByKeys_(ctx.sheet, GMP_SCHEMA.accounts.columns, ['NEXT_REVIEW_DUE', 'LAST_REVIEW_DATE'], startRow, requiredRows, 'yyyy-mm-dd');
  formatAccountGroupRows_(ctx.sheet, startRow, renderedRows.length);
}

function sortReviewSection_() {
  var sheet = getRequiredSheet_(GMP_SCHEMA.tabs.REVIEW_CHECKS);
  var startRow = GMP_SCHEMA.layout.REVIEW_CONTROL_ROW + 1;
  var endRow = Math.max(sheet.getLastRow(), startRow - 1);
  if (endRow < startRow) {
    return;
  }

  var values = sheet.getRange(startRow, 1, endRow - startRow + 1, GMP_SCHEMA.reviews.keys.length).getValues();
  var reviews = [];
  for (var i = 0; i < values.length; i++) {
    if (isBlankRow_(values[i]) || isDisplayGroupRow_(values[i])) {
      continue;
    }
    reviews.push(rowToObject_(values[i], GMP_SCHEMA.reviews.keys, GMP_SCHEMA.reviews.columns));
  }

  reviews = dedupeReviewRecordsForOutput_(reviews);
  reviews.sort(compareReviewRows_);
  var renderedRows = buildRenderedReviewRows_(reviews);
  var requiredRows = Math.max(renderedRows.length, endRow - startRow + 1);
  ensureSheetHasRows_(sheet, startRow + requiredRows - 1);

  var output = renderedRows.slice();
  while (output.length < requiredRows) {
    output.push(buildBlankRow_(GMP_SCHEMA.reviews.keys.length));
  }

  var targetRange = sheet.getRange(startRow, 1, requiredRows, GMP_SCHEMA.reviews.keys.length);
  replaceRangeValues_(targetRange, output);
  applyOperationalRowFormatting_(targetRange);
  applyDropdowns_(sheet, startRow, requiredRows, GMP_SCHEMA.reviews);
  applyDateFormatsByKeys_(sheet, GMP_SCHEMA.reviews.columns, ['LAST_LOGIN', 'DECISION_DATE'], startRow, requiredRows, 'yyyy-mm-dd hh:mm');
  applyDateFormatsByKeys_(sheet, GMP_SCHEMA.reviews.columns, ['NEXT_REVIEW_DUE'], startRow, requiredRows, 'yyyy-mm-dd');
  formatReviewGroupRows_(sheet, startRow, renderedRows.length);
  setReviewChecksColumnVisibility_(sheet);
}

function dedupeReviewRecordsForOutput_(reviews) {
  var byId = {};
  for (var i = 0; i < reviews.length; i++) {
    var reviewId = normalizeText_(reviews[i].REVIEW_ID);
    if (!reviewId) {
      continue;
    }
    if (!byId[reviewId]) {
      byId[reviewId] = reviews[i];
    } else {
      byId[reviewId] = chooseReviewRecordToKeep_(byId[reviewId], reviews[i]);
    }
  }

  var deduped = [];
  var seenIds = {};
  for (var j = 0; j < reviews.length; j++) {
    var currentId = normalizeText_(reviews[j].REVIEW_ID);
    if (currentId && byId[currentId] === reviews[j] && !seenIds[currentId]) {
      deduped.push(reviews[j]);
      seenIds[currentId] = true;
    } else if (!currentId) {
      deduped.push(reviews[j]);
    }
  }

  var byIdentity = {};
  var result = [];
  for (var k = 0; k < deduped.length; k++) {
    var identityKey = reviewReusableIdentityKey_(deduped[k]);
    if (!identityKey) {
      result.push(deduped[k]);
      continue;
    }
    if (!byIdentity[identityKey]) {
      byIdentity[identityKey] = [];
    }
    byIdentity[identityKey].push(deduped[k]);
  }

  var seenIdentity = {};
  for (var m = 0; m < deduped.length; m++) {
    var rowIdentity = reviewReusableIdentityKey_(deduped[m]);
    if (!rowIdentity || seenIdentity[rowIdentity]) {
      continue;
    }
    result.push(chooseReusableReviewRecordForIdentity_(byIdentity[rowIdentity]));
    seenIdentity[rowIdentity] = true;
  }

  return result;
}

function chooseReusableReviewRecordForIdentity_(records) {
  if (!records || !records.length) {
    return null;
  }

  var current = null;
  var latestDone = null;
  for (var i = 0; i < records.length; i++) {
    var status = normalizeText_(records[i].DECISION_STATUS);
    if (status !== 'Done') {
      if (!current) {
        current = records[i];
        continue;
      }
      current = chooseReviewRecordToKeep_(current, records[i]);
      continue;
    }

    if (!latestDone) {
      latestDone = records[i];
      continue;
    }

    latestDone = chooseReviewRecordToKeep_(latestDone, records[i]);
  }

  return current || latestDone || records[0];
}

function chooseReviewRecordToKeep_(left, right) {
  var leftScore = reviewRecordKeepScore_(left);
  var rightScore = reviewRecordKeepScore_(right);
  if (rightScore > leftScore) {
    return right;
  }
  if (rightScore < leftScore) {
    return left;
  }

  if (compareByDateDesc_(right.DECISION_DATE, left.DECISION_DATE) < 0) {
    return right;
  }
  if (compareByDateDesc_(right.NEXT_REVIEW_DUE, left.NEXT_REVIEW_DUE) < 0) {
    return right;
  }
  return left;
}

function reviewRecordKeepScore_(review) {
  var status = normalizeText_(review.DECISION_STATUS);
  var score = 0;
  if (status === 'Done') {
    score += 1000;
  } else if (status === 'Waiting') {
    score += 700;
  } else if (status === 'Open') {
    score += 500;
  }

  if (review.DECISION_DATE instanceof Date && !isNaN(review.DECISION_DATE)) {
    score += 40;
  }
  if (review.NEXT_REVIEW_DUE instanceof Date && !isNaN(review.NEXT_REVIEW_DUE)) {
    score += 20;
  }
  if (normalizeText_(review.REVIEWER_ACTION)) {
    score += 10;
  }
  if (normalizeText_(review.SLACK_THREAD)) {
    score += 5;
  }
  return score;
}

function sortIncidentSection_() {
  var sheet = getRequiredSheet_(GMP_SCHEMA.tabs.INCIDENTS);
  var startRow = GMP_SCHEMA.layout.INCIDENT_HEADER_ROW + 1;
  var endRow = sheet.getLastRow();
  sortTableSection_(
    sheet,
    startRow,
    endRow,
    GMP_SCHEMA.incidents.keys,
    GMP_SCHEMA.incidents.columns,
    compareIncidentRows_
  );
  if (endRow >= startRow) {
    var range = sheet.getRange(startRow, 1, endRow - startRow + 1, GMP_SCHEMA.incidents.keys.length);
    applyOperationalRowFormatting_(range);
    applyDropdowns_(sheet, startRow, endRow - startRow + 1, GMP_SCHEMA.incidents);
    applyDateFormatsByKeys_(sheet, GMP_SCHEMA.incidents.columns, ['REPORTED_AT', 'RESOLVED_AT'], startRow, endRow - startRow + 1, 'yyyy-mm-dd hh:mm');
  }
}

function sortTableSection_(sheet, startRow, endRow, keys, columnMap, comparator) {
  if (endRow < startRow) {
    return;
  }

  var values = sheet.getRange(startRow, 1, endRow - startRow + 1, keys.length).getValues();
  var populatedRows = [];
  var blankRows = 0;

  for (var i = 0; i < values.length; i++) {
    if (isBlankRow_(values[i])) {
      blankRows++;
      continue;
    }

    populatedRows.push(rowToObject_(values[i], keys, columnMap));
  }

  populatedRows.sort(comparator);

  var sortedValues = [];
  for (var j = 0; j < populatedRows.length; j++) {
    sortedValues.push(objectToRow_(populatedRows[j], keys));
  }

  for (var k = 0; k < blankRows; k++) {
    sortedValues.push(buildBlankRow_(keys.length));
  }

  replaceRangeValues_(sheet.getRange(startRow, 1, sortedValues.length, keys.length), sortedValues);
}

function buildBlankRow_(length) {
  var row = [];
  for (var i = 0; i < length; i++) {
    row.push('');
  }
  return row;
}

function buildRenderedAccountRows_(accounts) {
  var rows = [];
  var grouped = {};
  for (var i = 0; i < ACCOUNT_SYSTEM_GROUPS.length; i++) {
    grouped[ACCOUNT_SYSTEM_GROUPS[i]] = [];
  }

  for (var j = 0; j < accounts.length; j++) {
    var label = accountSystemGroupLabel_(accounts[j].GMP_SYSTEM);
    grouped[label].push(accounts[j]);
  }

  for (var k = 0; k < ACCOUNT_SYSTEM_GROUPS.length; k++) {
    var groupLabel = ACCOUNT_SYSTEM_GROUPS[k];
    rows.push(buildAccountGroupRow_(groupLabel, GMP_SCHEMA.accounts.keys.length));
    for (var m = 0; m < grouped[groupLabel].length; m++) {
      rows.push(objectToRow_(grouped[groupLabel][m], GMP_SCHEMA.accounts.keys));
    }
  }

  return rows;
}

function buildAccountGroupRow_(label, length) {
  var row = buildBlankRow_(length);
  row[0] = label;
  return row;
}

function buildRenderedReviewRows_(reviews) {
  var rows = [];
  var openRows = [];
  var waitingRows = [];
  var completedRows = [];
  var today = new Date();
  var todayStart = new Date(today.getFullYear(), today.getMonth(), today.getDate());

  for (var i = 0; i < reviews.length; i++) {
    var status = normalizeText_(reviews[i].DECISION_STATUS);
    if (status === 'Done') {
      completedRows.push(reviews[i]);
    } else if (isScheduledReviewRow_(reviews[i], todayStart)) {
      waitingRows.push(reviews[i]);
    } else {
      openRows.push(reviews[i]);
    }
  }

  if (openRows.length) {
    rows = rows.concat(buildRenderedReviewSectionRows_(GMP_SCHEMA.sectionLabels.OPEN_REVIEWS, openRows));
  }

  if (waitingRows.length) {
    rows = rows.concat(buildRenderedReviewSectionRows_(GMP_SCHEMA.sectionLabels.WAITING_REVIEWS, waitingRows));
  }

  if (completedRows.length) {
    rows = rows.concat(buildRenderedReviewSectionRows_(GMP_SCHEMA.sectionLabels.COMPLETED_REVIEWS, completedRows));
  }

  return rows;
}

function isScheduledReviewRow_(review, todayStart) {
  if (normalizeText_(review.DECISION_STATUS) === 'Waiting') {
    return true;
  }
  if (!(review.NEXT_REVIEW_DUE instanceof Date) || isNaN(review.NEXT_REVIEW_DUE)) {
    return false;
  }
  var nextCheckDay = new Date(review.NEXT_REVIEW_DUE.getFullYear(), review.NEXT_REVIEW_DUE.getMonth(), review.NEXT_REVIEW_DUE.getDate());
  return nextCheckDay.getTime() > todayStart.getTime();
}

function buildRenderedReviewSectionRows_(sectionLabel, reviews) {
  var rows = [buildReviewGroupRow_(sectionLabel, GMP_SCHEMA.reviews.keys.length)];
  var grouped = {};
  var extraLabels = [];
  for (var i = 0; i < ACCOUNT_SYSTEM_GROUPS.length; i++) {
    grouped[ACCOUNT_SYSTEM_GROUPS[i]] = [];
  }

  for (var j = 0; j < reviews.length; j++) {
    var groupLabel = accountSystemGroupLabel_(reviews[j].SYSTEM);
    if (!grouped[groupLabel]) {
      grouped[groupLabel] = [];
      extraLabels.push(groupLabel);
    }
    grouped[groupLabel].push(reviews[j]);
  }

  var orderedLabels = ACCOUNT_SYSTEM_GROUPS.slice();
  extraLabels.sort();
  for (var k = 0; k < extraLabels.length; k++) {
    orderedLabels.push(extraLabels[k]);
  }

  for (var m = 0; m < orderedLabels.length; m++) {
    var systemLabel = orderedLabels[m];
    if (!grouped[systemLabel] || !grouped[systemLabel].length) {
      continue;
    }
    rows.push(buildReviewGroupRow_(systemLabel + ' (' + grouped[systemLabel].length + ')', GMP_SCHEMA.reviews.keys.length));
    for (var n = 0; n < grouped[systemLabel].length; n++) {
      rows.push(objectToRow_(grouped[systemLabel][n], GMP_SCHEMA.reviews.keys));
    }
  }

  return rows;
}

function buildReviewGroupRow_(label, length) {
  var row = buildBlankRow_(length);
  row[0] = label;
  return row;
}

function formatAccountGroupRows_(sheet, startRow, rowCount) {
  for (var i = 0; i < rowCount; i++) {
    var rowNumber = startRow + i;
    var values = sheet.getRange(rowNumber, 1, 1, GMP_SCHEMA.accounts.keys.length).getValues()[0];
    var rowRange = sheet.getRange(rowNumber, 1, 1, GMP_SCHEMA.accounts.keys.length);
    if (!isDisplayGroupRow_(values)) {
      continue;
    }

    rowRange
      .setBackground('#e8f1fb')
      .setFontColor('#163b65')
      .setFontWeight('bold')
      .setHorizontalAlignment('left')
      .setVerticalAlignment('middle')
      .clearDataValidations();
    sheet.setRowHeight(rowNumber, 20);
  }
}

function formatReviewGroupRows_(sheet, startRow, rowCount) {
  for (var i = 0; i < rowCount; i++) {
    var rowNumber = startRow + i;
    var values = sheet.getRange(rowNumber, 1, 1, GMP_SCHEMA.reviews.keys.length).getValues()[0];
    var rowRange = sheet.getRange(rowNumber, 1, 1, GMP_SCHEMA.reviews.keys.length);
    if (!isDisplayGroupRow_(values)) {
      continue;
    }

    var label = normalizeText_(values[0]);
    if (label === GMP_SCHEMA.sectionLabels.OPEN_REVIEWS ||
        label === GMP_SCHEMA.sectionLabels.WAITING_REVIEWS ||
        label === GMP_SCHEMA.sectionLabels.COMPLETED_REVIEWS) {
      rowRange
        .setBackground(
          label === GMP_SCHEMA.sectionLabels.OPEN_REVIEWS ? '#dbeafe' :
          (label === GMP_SCHEMA.sectionLabels.WAITING_REVIEWS ? '#fef3c7' : '#e2e8f0')
        )
        .setFontColor(
          label === GMP_SCHEMA.sectionLabels.OPEN_REVIEWS ? '#163b65' :
          (label === GMP_SCHEMA.sectionLabels.WAITING_REVIEWS ? '#92400e' : '#334155')
        )
        .setFontWeight('bold')
        .setHorizontalAlignment('left')
        .setVerticalAlignment('middle')
        .clearDataValidations();
    } else {
      rowRange
        .setBackground('#eff6ff')
        .setFontColor('#1e3a5f')
        .setFontWeight('bold')
        .setHorizontalAlignment('left')
        .setVerticalAlignment('middle')
        .clearDataValidations();
    }
    sheet.setRowHeight(rowNumber, 20);
  }
}

function ensureSheetHasRows_(sheet, requiredLastRow) {
  if (requiredLastRow <= sheet.getMaxRows()) {
    return;
  }
  sheet.insertRowsAfter(sheet.getMaxRows(), requiredLastRow - sheet.getMaxRows());
}

function isDisplayGroupRow_(rowValues) {
  if (!rowValues || !rowValues.length) {
    return false;
  }

  if (!normalizeText_(rowValues[0])) {
    return false;
  }

  for (var i = 1; i < rowValues.length; i++) {
    if (normalizeText_(rowValues[i]) !== '') {
      return false;
    }
  }
  return true;
}

function compareRequestRows_(left, right) {
  return compareByNumber_(priorityRank_(left.QUEUE_PRIORITY), priorityRank_(right.QUEUE_PRIORITY)) ||
    compareByDateAsc_(left.SLA_DUE, right.SLA_DUE) ||
    compareByDateAsc_(left.SUBMITTED_AT, right.SUBMITTED_AT) ||
    compareTextAsc_(left.REQUEST_ID, right.REQUEST_ID);
}

function compareAccountRows_(left, right) {
  return compareByNumber_(accountSystemGroupRank_(left.GMP_SYSTEM), accountSystemGroupRank_(right.GMP_SYSTEM)) ||
    compareByNumber_(accountStateRank_(left.CURRENT_STATE), accountStateRank_(right.CURRENT_STATE)) ||
    compareByNumber_(activityBandRank_(left.ACTIVITY_BAND), activityBandRank_(right.ACTIVITY_BAND)) ||
    compareByDateAsc_(left.NEXT_REVIEW_DUE, right.NEXT_REVIEW_DUE) ||
    compareByDateDesc_(left.LAST_VERIFIED_AT, right.LAST_VERIFIED_AT) ||
    compareTextAsc_(left.PERSON, right.PERSON) ||
    compareTextAsc_(left.ACCESS_ID, right.ACCESS_ID);
}

function compareReviewRows_(left, right) {
  if (normalizeText_(left.DECISION_STATUS) === 'Done' && normalizeText_(right.DECISION_STATUS) === 'Done') {
    return compareByDateDesc_(left.DECISION_DATE, right.DECISION_DATE) ||
      compareTextAsc_(left.REVIEW_ID, right.REVIEW_ID);
  }
  return compareByNumber_(reviewStatusRank_(left.DECISION_STATUS), reviewStatusRank_(right.DECISION_STATUS)) ||
    compareByNumber_(priorityRank_(left.REVIEW_PRIORITY), priorityRank_(right.REVIEW_PRIORITY)) ||
    compareByDateAsc_(left.NEXT_REVIEW_DUE, right.NEXT_REVIEW_DUE) ||
    compareTextAsc_(left.REVIEW_ID, right.REVIEW_ID);
}

function compareIncidentRows_(left, right) {
  return compareByNumber_(incidentStatusRank_(left.STATUS), incidentStatusRank_(right.STATUS)) ||
    compareByNumber_(priorityRank_(left.SEVERITY), priorityRank_(right.SEVERITY)) ||
    compareByDateDesc_(left.REPORTED_AT, right.REPORTED_AT) ||
    compareTextAsc_(left.INCIDENT_ID, right.INCIDENT_ID);
}

function compareByNumber_(left, right) {
  var a = typeof left === 'number' ? left : 999999;
  var b = typeof right === 'number' ? right : 999999;
  if (a < b) {
    return -1;
  }
  if (a > b) {
    return 1;
  }
  return 0;
}

function compareByDateAsc_(left, right) {
  return compareByNumber_(dateSortValue_(left, true), dateSortValue_(right, true));
}

function compareByDateDesc_(left, right) {
  return compareByNumber_(dateSortValue_(right, false), dateSortValue_(left, false));
}

function dateSortValue_(value, blanksLast) {
  if (!(value instanceof Date) || isNaN(value)) {
    return blanksLast ? 8640000000000000 : -8640000000000000;
  }
  return value.getTime();
}

function compareTextAsc_(left, right) {
  var a = normalizeText_(left).toLowerCase();
  var b = normalizeText_(right).toLowerCase();
  if (a < b) {
    return -1;
  }
  if (a > b) {
    return 1;
  }
  return 0;
}

function accountStateRank_(state) {
  var value = normalizePresenceState_(state);
  if (value === 'Setting Up') {
    return 0;
  }
  if (value === 'Missing') {
    return 1;
  }
  if (value === 'Unmanaged') {
    return 2;
  }
  if (value === 'Present') {
    return 3;
  }
  if (value === 'Revoked') {
    return 4;
  }
  return 5;
}

function activityBandRank_(band) {
  var value = normalizeActivityStatus_(band);
  if (value === 'No Evidence') {
    return 0;
  }
  if (value === 'Unknown') {
    return 1;
  }
  if (value === 'Some Evidence') {
    return 2;
  }
  if (value === 'Verified Active') {
    return 3;
  }
  return 4;
}

function reviewStatusRank_(status) {
  var value = normalizeText_(status);
  if (value === 'Open') {
    return 0;
  }
  if (value === 'Waiting') {
    return 1;
  }
  if (value === 'Done') {
    return 2;
  }
  return 3;
}

function incidentStatusRank_(status) {
  var value = normalizeText_(status);
  if (value === 'Open') {
    return 0;
  }
  if (value === 'Investigating') {
    return 1;
  }
  if (value === 'Resolved') {
    return 2;
  }
  return 3;
}

function accountSystemGroupLabel_(systemName) {
  var canonical = canonicalSystemName_(systemName);
  if (canonical) {
    return canonical;
  }

  var value = normalizeText_(systemName).toLowerCase();
  if (value.indexOf('katana') !== -1) {
    return 'Katana';
  }
  if (value.indexOf('wasp') !== -1) {
    return 'Wasp';
  }
  if (value.indexOf('shopify') !== -1) {
    return 'Shopify';
  }
  if (value.indexOf('shipstation') !== -1 || value.indexOf('ship station') !== -1) {
    return 'ShipStation';
  }
  if (value.indexOf('1password') !== -1 || value.indexOf('one password') !== -1) {
    return '1Password';
  }
  return 'Other Systems';
}

function canonicalSystemName_(systemName) {
  var value = normalizeText_(systemName).toLowerCase();
  if (!value) {
    return '';
  }
  if (value.indexOf('katana') !== -1) {
    return 'Katana';
  }
  if (value.indexOf('wasp') !== -1) {
    return 'Wasp';
  }
  if (value.indexOf('shopify') !== -1) {
    return 'Shopify';
  }
  if (value.indexOf('shipstation') !== -1 || value.indexOf('ship station') !== -1) {
    return 'ShipStation';
  }
  if (value.indexOf('1password') !== -1 || value.indexOf('one password') !== -1) {
    return '1Password';
  }
  return normalizeText_(systemName);
}

function accountSystemGroupRank_(systemName) {
  var label = accountSystemGroupLabel_(systemName);
  for (var i = 0; i < ACCOUNT_SYSTEM_GROUPS.length; i++) {
    if (ACCOUNT_SYSTEM_GROUPS[i] === label) {
      return i;
    }
  }
  return ACCOUNT_SYSTEM_GROUPS.length;
}

function normalizePresenceState_(state) {
  var value = normalizeText_(state);
  if (value === 'Provisioning') {
    return 'Setting Up';
  }
  if (value === 'Active') {
    return 'Present';
  }
  if (value === 'Stale') {
    return 'Missing';
  }
  if (value === 'Review Due') {
    return 'Unmanaged';
  }
  return value;
}

function normalizeSetupStatus_(status) {
  var value = normalizeText_(status);
  if (value === 'Provisioning') {
    return 'In Progress';
  }
  if (value === 'Provisioned') {
    return 'Completed';
  }
  if (value === 'Removal Pending') {
    return 'Removing Access';
  }
  return value;
}

function normalizeActivityStatus_(status) {
  var value = normalizeText_(status);
  if (value === 'Healthy') {
    return 'Verified Active';
  }
  if (value === 'Watch' || value === 'Low Use') {
    return 'Some Evidence';
  }
  if (value === 'Critical') {
    return 'No Evidence';
  }
  return value;
}
