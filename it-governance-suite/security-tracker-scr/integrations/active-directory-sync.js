function onOpen() {
  var ui = SpreadsheetApp.getUi();
  var advancedMenu = ui.createMenu('Advanced')
    .addItem('Refresh Katana Accounts', 'refreshKatanaAccounts')
    .addItem('Refresh Activity Signals', 'refreshActivitySignals')
    .addItem('Refresh Shopify Activity', 'refreshShopifyActivity')
    .addSeparator()
    .addItem('Setup 3-Tab Workbook', 'setupSheet')
    .addItem('Setup Triggers', 'setupTriggers')
    .addSeparator()
    .addItem('Test Slack Connection', 'testSlackConnection');

  ui
    .createMenu('GMP Security')
    .addItem('Sync Approved Requests', 'syncAccessRequestsToAccounts')
    .addItem('Generate Review Checks', 'generateReviewChecks')
    .addItem('Run Daily Maintenance', 'runDailySecurityMaintenance')
    .addSeparator()
    .addSubMenu(advancedMenu)
    .addToUi();
}

function syncAccessRequestsToAccounts() {
  ensureThreeTabWorkbookReady_();
  var requestRows = getRequestRows_();
  var created = 0;
  var updated = 0;
  var revoked = 0;
  var denied = 0;
  var archived = 0;

  for (var i = 0; i < requestRows.length; i++) {
    var result = syncAccessRequestRowToAccount_(requestRows[i]);
    if (!result) {
      continue;
    }
    if (result.action === 'created') {
      created++;
    } else if (result.action === 'updated') {
      updated++;
    } else if (result.action === 'revoked') {
      revoked++;
    } else if (result.action === 'denied') {
      denied++;
    }
    if (archiveRequestRow_(requestRows[i])) {
      archived++;
    }
  }

  sortRequestSection_();
  sortAccountSection_();
  formatAccessControlSheet_(getRequiredSheet_(GMP_SCHEMA.tabs.ACCESS_CONTROL));

  sendHeartbeat('Request Sync', 'Created ' + created + ', updated ' + updated + ', revoked ' + revoked + ', denied ' + denied + ', archived ' + archived);

  var message = 'Request sync complete.\n\n' +
    'Created accounts: ' + created + '\n' +
    'Updated accounts: ' + updated + '\n' +
    'Revoked accounts: ' + revoked + '\n' +
    'Denied requests: ' + denied + '\n' +
    'Archived requests: ' + archived;

  try {
    SpreadsheetApp.getUi().alert('Sync Complete', message, SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (err) {
    Logger.log(message);
  }
}

function syncAccessRequestRowToAccount_(requestRow) {
  var request = requestRow.data;
  var approval = normalizeText_(request.APPROVAL);
  var provisioning = normalizeSetupStatus_(request.PROVISIONING);
  var requestType = normalizeText_(request.REQUEST_TYPE);

  if (approval === 'Denied') {
    return finalizeDeniedRequest_(requestRow);
  }

  if (approval !== 'Approved') {
    return null;
  }

  if (requestType === 'New Access' || requestType === 'Change Access') {
    return upsertActiveAccountFromRequest_(requestRow);
  }

  if (requestType === 'Remove Access' && (provisioning === 'Completed' || provisioning === 'Closed')) {
    return revokeActiveAccountFromRequest_(requestRow);
  }

  return null;
}

function finalizeDeniedRequest_(requestRow) {
  var request = requestRow.data;
  request.PROVISIONING = 'Closed';
  request.DECISION_DATE = request.DECISION_DATE || new Date();
  request.NEXT_SLACK_TRIGGER = 'Post denial update';
  updateRequestRow_(requestRow.rowNumber, request, requestRow.sheetName);

  return {
    action: 'denied',
    accessId: normalizeText_(request.ACCESS_ID)
  };
}

function upsertActiveAccountFromRequest_(requestRow) {
  var request = requestRow.data;
  var canonicalSystem = canonicalSystemName_(request.GMP_SYSTEM);
  var accessId = normalizeIdentifier_(request.ACCESS_ID);
  var existingAccountRow = findBestAccountRowForRequest_(request);
  var provisioning = normalizeSetupStatus_(request.PROVISIONING);

  if (!provisioning || provisioning === 'Open') {
    provisioning = 'In Progress';
    request.PROVISIONING = provisioning;
  }

  if (!accessId) {
    accessId = existingAccountRow ? normalizeIdentifier_(existingAccountRow.data.ACCESS_ID) : nextStableId_('ACC');
  }

  var account = existingAccountRow ? existingAccountRow.data : buildEmptyRecord_(GMP_SCHEMA.accounts.keys);
  var privileged = isPrivilegedAccessLevel_(request.ACCESS_LEVEL) ? 'Yes' : 'No';
  var isProvisioned = provisioning === 'Completed' || provisioning === 'Closed';
  var requestType = normalizeText_(request.REQUEST_TYPE);
  var currentState = normalizePresenceState_(account.CURRENT_STATE || '');
  var targetState = isProvisioned ? 'Present' : ((requestType === 'Change Access' && existingAccountRow) ? (currentState || 'Present') : 'Setting Up');
  var grantedOn = isProvisioned ? (account.GRANTED_ON || request.DECISION_DATE || new Date()) : (account.GRANTED_ON || '');

  account.ACCESS_ID = accessId;
  account.SOURCE_REQUEST_ID = request.REQUEST_ID;
  account.GMP_SYSTEM = canonicalSystem;
  account.PERSON = request.TARGET_USER;
  account.COMPANY_EMAIL = request.COMPANY_EMAIL;
  account.DEPT = request.DEPT;
  account.PLATFORM_ACCOUNT = request.COMPANY_EMAIL;
  account.ACCESS_LEVEL = request.ACCESS_LEVEL;
  account.PRIVILEGED = privileged;
  account.CURRENT_STATE = targetState;
  account.LAST_VERIFIED_AT = isProvisioned ? new Date() : (account.LAST_VERIFIED_AT || '');
  account.GRANTED_ON = grantedOn;
  account.ACTIVITY_BAND = account.ACTIVITY_BAND || 'Unknown';
  account.EXPECTED_USE = account.EXPECTED_USE || defaultExpectedUse_(canonicalSystem);
  account.MIN_LOGINS_30D = account.MIN_LOGINS_30D || defaultMinLogins30d_(account.EXPECTED_USE);
  account.STALE_AFTER_DAYS = account.STALE_AFTER_DAYS || defaultStaleAfterDays_(canonicalSystem, account.EXPECTED_USE);
  account.ACTIVITY_SOURCE = account.ACTIVITY_SOURCE || 'Manual review';
  account.ACTIVITY_CONFIDENCE = account.ACTIVITY_CONFIDENCE || 'Low';
  account.REVIEW_TRIGGER = isProvisioned && privileged === 'Yes' ? 'Privileged Access' : account.REVIEW_TRIGGER;
  account.REVIEW_STATUS = isProvisioned ? (account.REVIEW_STATUS || '') : account.REVIEW_STATUS;
  account.NEXT_REVIEW_DUE = isProvisioned ? (account.NEXT_REVIEW_DUE || defaultNextReviewDue_(canonicalSystem, request.ACCESS_LEVEL)) : account.NEXT_REVIEW_DUE;
  account.MFA = account.MFA || '';
  account.OWNER = request.IT_OWNER;
  account.SLACK_THREAD = request.SLACK_THREAD || account.SLACK_THREAD;
  account.NOTES = appendNoteLine_(
    account.NOTES,
    formatDateOnly_(new Date()) + ': synced ' + request.REQUEST_ID + (isProvisioned ? '' : ' setup')
  );

  if (existingAccountRow) {
    updateAccountRow_(existingAccountRow.rowNumber, account);
  } else {
    appendAccountRecord_(account);
  }

  request.ACCESS_ID = accessId;
  request.GMP_SYSTEM = canonicalSystem;
  request.DECISION_DATE = request.DECISION_DATE || new Date();
  request.NEXT_SLACK_TRIGGER = isProvisioned ? 'Post setup update' : 'Continue setup';
  updateRequestRow_(requestRow.rowNumber, request, requestRow.sheetName);

  return {
    action: existingAccountRow ? 'updated' : 'created',
    accessId: accessId
  };
}

function revokeActiveAccountFromRequest_(requestRow) {
  var request = requestRow.data;
  var canonicalSystem = canonicalSystemName_(request.GMP_SYSTEM);
  var accessRow = findBestAccountRowForRequest_(request);

  if (!accessRow) {
    return null;
  }

  var account = accessRow.data;
  account.CURRENT_STATE = 'Revoked';
  account.REVIEW_STATUS = 'Done';
  account.REVIEW_TRIGGER = 'Access Removed';
  account.NEXT_REVIEW_DUE = '';
  account.LAST_VERIFIED_AT = new Date();
  account.SLACK_THREAD = request.SLACK_THREAD || account.SLACK_THREAD;
  account.NOTES = appendNoteLine_(
    account.NOTES,
    formatDateOnly_(new Date()) + ': revoked ' + request.REQUEST_ID
  );
  updateAccountRow_(accessRow.rowNumber, account);

  request.ACCESS_ID = account.ACCESS_ID;
  request.GMP_SYSTEM = canonicalSystem;
  request.DECISION_DATE = request.DECISION_DATE || new Date();
  request.PROVISIONING = normalizeSetupStatus_(request.PROVISIONING);
  request.NEXT_SLACK_TRIGGER = 'Confirm removal in Slack';
  updateRequestRow_(requestRow.rowNumber, request, requestRow.sheetName);

  return {
    action: 'revoked',
    accessId: account.ACCESS_ID
  };
}

function buildEmptyRecord_(keys) {
  var record = {};
  for (var i = 0; i < keys.length; i++) {
    record[keys[i]] = '';
  }
  return record;
}

function isPrivilegedAccessLevel_(accessLevel) {
  var value = normalizeText_(accessLevel).toLowerCase();
  return value === 'admin' || value === 'full access' || value === 'read-write';
}

function defaultExpectedUse_(systemName) {
  var value = canonicalSystemName_(systemName).toLowerCase();
  if (value.indexOf('katana') !== -1) {
    return 'Daily';
  }
  if (value.indexOf('google drive') !== -1) {
    return 'Ad Hoc';
  }
  if (value.indexOf('shopify') !== -1) {
    return 'Weekly';
  }
  return 'Weekly';
}

function defaultMinLogins30d_(expectedUse) {
  var value = normalizeText_(expectedUse).toLowerCase();
  if (value === 'daily') {
    return 12;
  }
  if (value === 'weekly') {
    return 4;
  }
  if (value === 'ad hoc') {
    return 1;
  }
  return 2;
}

function defaultStaleAfterDays_(systemName, expectedUse) {
  var value = canonicalSystemName_(systemName).toLowerCase();
  if (value.indexOf('katana') !== -1) {
    return 14;
  }
  if (value.indexOf('shopify') !== -1) {
    return 30;
  }

  var useValue = normalizeText_(expectedUse).toLowerCase();
  if (useValue === 'daily') {
    return 14;
  }
  if (useValue === 'weekly') {
    return 30;
  }
  return 45;
}

function reviewIntervalDays_(systemName, accessLevel) {
  var system = canonicalSystemName_(systemName).toLowerCase();
  var level = normalizeText_(accessLevel).toLowerCase();

  if (system.indexOf('1password') !== -1) {
    if (level === 'admin') {
      return 30;
    }
    if (level === 'full access') {
      return 45;
    }
    if (level === 'read-write') {
      return 60;
    }
    return 90;
  }

  if (system.indexOf('katana') !== -1 || system.indexOf('wasp') !== -1) {
    if (level === 'admin') {
      return 45;
    }
    if (level === 'full access') {
      return 60;
    }
    if (level === 'read-write') {
      return 90;
    }
    return 120;
  }

  if (system.indexOf('shopify') !== -1 || system.indexOf('shipstation') !== -1) {
    if (level === 'admin') {
      return 60;
    }
    if (level === 'full access') {
      return 90;
    }
    if (level === 'read-write') {
      return 120;
    }
    return 120;
  }

  if (level === 'admin') {
    return 90;
  }
  if (level === 'full access') {
    return 120;
  }
  if (level === 'read-write') {
    return 120;
  }
  return 150;
}

function scheduledReviewTriggerLabel_(systemName, accessLevel) {
  return reviewIntervalDays_(systemName, accessLevel) + '-day review due';
}

function nextReviewDueFromDecision_(decisionDate, systemName, accessLevel, reviewerAction) {
  if (!(decisionDate instanceof Date) || isNaN(decisionDate)) {
    return '';
  }
  if (normalizeText_(reviewerAction) === 'Remove') {
    return addDays_(decisionDate, 14);
  }
  return addDays_(decisionDate, reviewIntervalDays_(systemName, accessLevel));
}

function defaultNextReviewDue_(systemName, accessLevel) {
  return addDays_(new Date(), reviewIntervalDays_(systemName, accessLevel));
}
