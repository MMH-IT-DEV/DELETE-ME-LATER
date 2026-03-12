function refreshShopifyActivity() {
  ensureThreeTabWorkbookReady_();

  var token = PropertiesService.getScriptProperties().getProperty(GMP_CONFIG.SHOPIFY_TOKEN_PROPERTY);
  var store = PropertiesService.getScriptProperties().getProperty(GMP_CONFIG.SHOPIFY_STORE_PROPERTY) || 'mymagichealer.myshopify.com';

  if (!token) {
    throw new Error('Missing SHOPIFY_TOKEN script property.');
  }

  var response = UrlFetchApp.fetch('https://' + store + '/admin/api/2024-01/users.json', {
    method: 'get',
    headers: { 'X-Shopify-Access-Token': token },
    muteHttpExceptions: true
  });

  if (response.getResponseCode() !== 200) {
    throw new Error('Shopify API returned HTTP ' + response.getResponseCode());
  }

  var payload = JSON.parse(response.getContentText());
  var users = payload.users || [];
  var loginMap = {};

  for (var i = 0; i < users.length; i++) {
    var user = users[i];
    if (!user.email) {
      continue;
    }
    loginMap[normalizeText_(user.email).toLowerCase()] = user.last_login_at ? new Date(user.last_login_at) : '';
  }

  var accountRows = getAccountRows_();
  var updated = 0;
  for (var j = 0; j < accountRows.length; j++) {
    var account = accountRows[j].data;
    if (normalizeText_(account.GMP_SYSTEM).toLowerCase().indexOf('shopify') === -1) {
      continue;
    }

    var email = normalizeText_(account.COMPANY_EMAIL).toLowerCase();
    if (!loginMap.hasOwnProperty(email) || !loginMap[email]) {
      continue;
    }

    account.LAST_LOGIN = loginMap[email];
    account.DAYS_SINCE_LOGIN = daysBetween_(account.LAST_LOGIN, new Date());
    account.LAST_VERIFIED_AT = new Date();
    if (normalizePresenceState_(account.CURRENT_STATE) !== 'Revoked') {
      account.CURRENT_STATE = 'Present';
    }
    account.ACTIVITY_SOURCE = 'Shopify Admin API';
    account.ACTIVITY_CONFIDENCE = 'High';
    updateAccountRow_(accountRows[j].rowNumber, account);
    updated++;
  }

  sortAccountSection_();
  refreshActivitySignals();
  sendHeartbeat('Shopify Activity Refresh', updated + ' account rows updated');
  return updated;
}

function refreshActivitySignals() {
  ensureThreeTabWorkbookReady_();
  var accountRows = getAccountRows_();
  var today = new Date();
  var updated = 0;

  for (var i = 0; i < accountRows.length; i++) {
    var account = accountRows[i].data;
    if (normalizeText_(account.ACCESS_ID) === '') {
      continue;
    }

    if (account.LAST_LOGIN instanceof Date && !isNaN(account.LAST_LOGIN)) {
      account.DAYS_SINCE_LOGIN = daysBetween_(account.LAST_LOGIN, today);
    } else if (normalizeText_(account.DAYS_SINCE_LOGIN) !== '') {
      account.DAYS_SINCE_LOGIN = toInteger_(account.DAYS_SINCE_LOGIN, '');
    } else {
      account.DAYS_SINCE_LOGIN = '';
    }

    account.ACTIVITY_SCORE = calculateActivityScore_(account, today);
    account.ACTIVITY_BAND = calculateActivityBand_(account, today);
    if (!(account.NEXT_REVIEW_DUE instanceof Date) || isNaN(account.NEXT_REVIEW_DUE)) {
      var state = normalizePresenceState_(account.CURRENT_STATE);
      if (state !== 'Revoked' && state !== 'Setting Up') {
        account.NEXT_REVIEW_DUE = defaultNextReviewDue_(account.GMP_SYSTEM, account.ACCESS_LEVEL);
      }
    }

    var trigger = determineReviewTrigger_(account, today);
    account.REVIEW_TRIGGER = trigger ? trigger.type : '';
    if (!trigger && normalizeText_(account.REVIEW_STATUS) === 'Open') {
      account.REVIEW_STATUS = '';
    }

    updateAccountRow_(accountRows[i].rowNumber, account);
    updated++;
  }

  sortAccountSection_();
  sendHeartbeat('Activity Refresh', updated + ' account rows recalculated');
  return updated;
}

function generateReviewChecks() {
  ensureThreeTabWorkbookReady_();
  var accountRows = getAccountRows_();
  var today = new Date();
  var created = 0;
  var updated = 0;

  for (var i = 0; i < accountRows.length; i++) {
    var accountRow = accountRows[i];
    var account = accountRow.data;
    var trigger = determineReviewTrigger_(account, today);

    if (!trigger || normalizePresenceState_(account.CURRENT_STATE) === 'Revoked' || normalizePresenceState_(account.CURRENT_STATE) === 'Setting Up') {
      continue;
    }

    var reviewRow = findOpenReviewRowForAccessAny_(account);
    var review = reviewRow ? reviewRow.data : buildEmptyRecord_(GMP_SCHEMA.reviews.keys);
    var previousStatus = reviewRow ? normalizeText_(review.DECISION_STATUS) : '';

    if (!review.REVIEW_ID) {
      review.REVIEW_ID = nextStableId_('REV');
    }

    review.REVIEW_CYCLE = reviewCycleLabel_(trigger.reviewDue || account.NEXT_REVIEW_DUE || today);
    review.REVIEW_PRIORITY = trigger.priority;
    review.PERSON = account.PERSON;
    review.SYSTEM = account.GMP_SYSTEM;
    review.ACCESS_LEVEL = account.ACCESS_LEVEL;
    review.PRESENCE_STATUS = normalizePresenceState_(account.CURRENT_STATE);
    review.ACTIVITY_STATUS = normalizeActivityStatus_(account.ACTIVITY_BAND);
    review.LAST_LOGIN = account.LAST_LOGIN;
    review.LOGINS_30D = account.LOGINS_30D;
    review.WHY_FLAGGED = describeReviewFlag_(account, trigger);
    if (previousStatus === 'Done') {
      review.REVIEWER_ACTION = '';
      review.NOTES = appendUniqueNoteLine_(review.NOTES, formatDateOnly_(today) + ': review cycle reopened');
    } else {
      review.REVIEWER_ACTION = review.REVIEWER_ACTION || '';
    }
    review.DECISION_STATUS = 'Open';
    if (reviewRow && previousStatus === 'Waiting') {
      review.DECISION_STATUS = 'Open';
      review.NEXT_REVIEW_DUE = '';
      review.NOTES = appendNoteLine_(review.NOTES, formatDateOnly_(today) + ': follow-up reopened');
    }
    review.NEXT_REVIEW_DUE = '';
    review.NOTES = appendUniqueNoteLine_(review.NOTES, formatDateOnly_(today) + ': review refreshed');
    review.TRIGGER_TYPE = trigger.type;
    review.ACCESS_ID = account.ACCESS_ID;
    review.DAYS_SINCE_LOGIN = account.DAYS_SINCE_LOGIN;
    review.ACTIVITY_SCORE = account.ACTIVITY_SCORE;
    review.DECISION_DATE = '';
    review.LINKED_REQUEST = account.SOURCE_REQUEST_ID;
    review.SLACK_THREAD = account.SLACK_THREAD;
    review.NEXT_SLACK_TRIGGER = 'Post review queue alert';

    if (reviewRow) {
      updateReviewRow_(reviewRow.rowNumber, review);
      updated++;
    } else {
      appendReviewRecord_(review);
      created++;
    }

    account.REVIEW_TRIGGER = trigger.type;
    account.REVIEW_STATUS = review.DECISION_STATUS;
    account.NEXT_REVIEW_DUE = review.DECISION_STATUS === 'Waiting'
      ? review.NEXT_REVIEW_DUE
      : (trigger.reviewDue || account.NEXT_REVIEW_DUE || today);
    updateAccountRow_(accountRow.rowNumber, account);

    if (normalizeText_(account.SOURCE_REQUEST_ID)) {
      linkReviewToSourceRequest_(account.SOURCE_REQUEST_ID, review.REVIEW_ID);
    }
  }

  sortAccountSection_();
  sortReviewSection_();
  formatReviewChecksSheet_(getRequiredSheet_(GMP_SCHEMA.tabs.REVIEW_CHECKS));
  sendHeartbeat('Review Queue Generation', 'Created ' + created + ', updated ' + updated);
  return { created: created, updated: updated };
}

function runDailySecurityMaintenance() {
  ensureThreeTabWorkbookReady_();

  syncAccessRequestsToAccounts();

  var katanaApiKey = PropertiesService.getScriptProperties().getProperty(GMP_CONFIG.KATANA_API_KEY_PROPERTY);
  if (katanaApiKey) {
    refreshKatanaAccountsInternal_(false);
  }

  var token = PropertiesService.getScriptProperties().getProperty(GMP_CONFIG.SHOPIFY_TOKEN_PROPERTY);
  if (token) {
    refreshShopifyActivity();
  } else {
    refreshActivitySignals();
  }

  generateReviewChecks();
  sendOpenReviewSummary();
}

function calculateActivityScore_(account, today) {
  var staleAfterDays = toInteger_(account.STALE_AFTER_DAYS, 30);
  var daysSinceLogin = normalizeText_(account.DAYS_SINCE_LOGIN) === '' ? '' : toInteger_(account.DAYS_SINCE_LOGIN, staleAfterDays * 2);
  var min30d = Math.max(toInteger_(account.MIN_LOGINS_30D, defaultMinLogins30d_(account.EXPECTED_USE)), 1);
  var logins30d = Math.max(toInteger_(account.LOGINS_30D, 0), 0);
  var nextReviewDue = account.NEXT_REVIEW_DUE instanceof Date ? account.NEXT_REVIEW_DUE : '';
  var hasUsageEvidence = hasUsageEvidence_(account);

  var freshnessScore = 0;
  if (!hasUsageEvidence) {
    freshnessScore = 30;
  } else if (daysSinceLogin === '') {
    freshnessScore = 8;
  } else {
    freshnessScore = Math.max(0, 40 - Math.round((daysSinceLogin / staleAfterDays) * 40));
  }

  var usageScore = 0;
  if (!hasUsageEvidence) {
    usageScore = 25;
  } else {
    var usageRatio = Math.min(logins30d / min30d, 1.5);
    usageScore = Math.round(Math.min(usageRatio, 1) * 30);
  }

  var controlScore = 0;
  if (isYesValue_(account.MFA)) {
    controlScore += 10;
  } else if (normalizeText_(account.MFA) === 'N/A') {
    controlScore += 5;
  }

  if (!nextReviewDue || !(nextReviewDue instanceof Date) || nextReviewDue >= today) {
    controlScore += 10;
  }

  var privilegedModifier = isYesValue_(account.PRIVILEGED) ? 5 : 10;
  return Math.max(0, Math.min(100, freshnessScore + usageScore + controlScore + privilegedModifier));
}

function calculateActivityBand_(account) {
  var hasUsageEvidence = hasUsageEvidence_(account);
  if (!hasUsageEvidence) {
    return 'Unknown';
  }

  var min30d = Math.max(toInteger_(account.MIN_LOGINS_30D, defaultMinLogins30d_(account.EXPECTED_USE)), 1);
  var logins30d = Math.max(toInteger_(account.LOGINS_30D, 0), 0);
  var daysSinceLogin = normalizeText_(account.DAYS_SINCE_LOGIN) === '' ? '' : toInteger_(account.DAYS_SINCE_LOGIN, 0);

  if ((daysSinceLogin !== '' && daysSinceLogin <= 30) || logins30d >= min30d) {
    return 'Verified Active';
  }
  if (daysSinceLogin !== '' || logins30d > 0 || toInteger_(account.LOGINS_90D, 0) > 0) {
    return 'Some Evidence';
  }
  return 'No Evidence';
}

function determineReviewTrigger_(account, today) {
  var state = normalizePresenceState_(account.CURRENT_STATE);
  if (state === 'Revoked' || state === 'Setting Up') {
    return null;
  }

  var staleAfterDays = Math.max(toInteger_(account.STALE_AFTER_DAYS, 30), 1);
  var daysSinceLogin = normalizeText_(account.DAYS_SINCE_LOGIN) === '' ? '' : toInteger_(account.DAYS_SINCE_LOGIN, 0);
  var nextReviewDue = account.NEXT_REVIEW_DUE instanceof Date ? account.NEXT_REVIEW_DUE : '';
  var activityBand = normalizeActivityStatus_(account.ACTIVITY_BAND);
  var hasUsageEvidence = hasUsageEvidence_(account);
  var reviewStatus = normalizeText_(account.REVIEW_STATUS);

  if ((reviewStatus === 'Waiting' || reviewStatus === 'Open') && nextReviewDue && nextReviewDue > today) {
    return null;
  }

  if (state === 'Missing') {
    return { type: 'Missing From System', priority: 'High', reviewDue: today };
  }
  if (state === 'Unmanaged') {
    return { type: 'Unmanaged Account', priority: 'High', reviewDue: today };
  }
  if (nextReviewDue && nextReviewDue <= today) {
    return { type: scheduledReviewTriggerLabel_(account.GMP_SYSTEM, account.ACCESS_LEVEL), priority: 'Medium', reviewDue: nextReviewDue };
  }
  if (hasUsageEvidence && (activityBand === 'No Evidence' || (daysSinceLogin !== '' && daysSinceLogin > staleAfterDays))) {
    return { type: 'No Activity Evidence', priority: 'High', reviewDue: today };
  }
  return null;
}

function hasUsageEvidence_(account) {
  if (account.LAST_LOGIN instanceof Date && !isNaN(account.LAST_LOGIN)) {
    return true;
  }
  if (toInteger_(account.LOGINS_30D, 0) > 0 || toInteger_(account.LOGINS_90D, 0) > 0) {
    return true;
  }

  var source = normalizeText_(account.ACTIVITY_SOURCE);
  if (!source) {
    return false;
  }
  return source !== 'Katana Users API' && source !== 'Manual review';
}

function describeReviewFlag_(account, trigger) {
  if (trigger.type === 'Missing From System') {
    return 'Missing in latest system check';
  }
  if (trigger.type === 'Unmanaged Account') {
    return 'No approved request on file';
  }
  if (normalizeText_(trigger.type).toLowerCase().indexOf('review due') !== -1) {
    return normalizeText_(trigger.type).toLowerCase();
  }
  if (trigger.type === 'No Activity Evidence') {
    return 'No recent activity evidence';
  }
  return 'Account review needed';
}

function findOpenReviewRowForAccessAny_(account) {
  var accountIdentity = accountReusableReviewIdentityKey_(account);
  var reviewRows = getReviewRows_();
  var relatedRows = [];
  for (var i = 0; i < reviewRows.length; i++) {
    if (reviewReusableIdentityKey_(reviewRows[i].data) !== accountIdentity) {
      continue;
    }
    relatedRows.push(reviewRows[i]);
  }
  return chooseCurrentReviewRowForAccount_(relatedRows) || chooseLatestDoneReviewRowForAccount_(relatedRows);
}

function linkReviewToSourceRequest_(requestId, reviewId) {
  var requestRow = findRequestRowById_(requestId);
  if (!requestRow) {
    return;
  }
  requestRow.data.LINKED_REVIEW_ID = reviewId;
  updateRequestRow_(requestRow.rowNumber, requestRow.data, requestRow.sheetName);
}

function reviewCycleLabel_(dateValue) {
  var date = dateValue instanceof Date && !isNaN(dateValue) ? dateValue : new Date();
  var year = Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy');
  var cycle = Math.floor(date.getMonth() / 4) + 1;
  return year + '-C' + cycle;
}
