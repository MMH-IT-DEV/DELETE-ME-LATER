/**
 * IT Security Tracker - Heartbeat & Activity Logging to Health Monitor
 *
 * This file runs inside the Systems Health Tracker script project.
 * It logs activity from the IT Security Tracker sheet into this health tracker.
 *
 * Webhook URL must be set in Script Properties:
 *   Key: SLACK_WEBHOOK_URL
 *   Value: your Slack webhook URL
 *
 * CRITICAL: GAS V8 — var only, no const/let, no arrow functions, no template literals
 */

var SEC_TRACKER = {
  SHEET_ID:       '1WBzed_RMtPC3kii-ybJ3lE4Fo1pAqWBt9OJyCSVR-mQ',
  SHEET_URL:      'https://docs.google.com/spreadsheets/d/1WBzed_RMtPC3kii-ybJ3lE4Fo1pAqWBt9OJyCSVR-mQ/edit',
  HEALTH_SHEET_ID: '1jnWtdBPzR7DreihCHQASiN7splmRS7HYJTR77GpfI5w',
  SCRIPT_NAME:    'IT Security Tracker',

  TABS: {
    INCIDENTS:       'Incidents',
    ACCESS_REQUESTS: 'Access Requests',
    ACCESS_REGISTER: 'Access Register',
    REVIEWS:         'Reviews'
  },

  COLUMNS: {
    INC_STATUS:      11,
    AR_STATUS:       9,
    REV_NEXT_REVIEW: 11
  },

  AUTO_IDS: {
    INCIDENT_OPEN:  'SEC-001',
    INCIDENT_CLOSE: 'SEC-002',
    ACCESS_PENDING: 'SEC-003',
    ACCESS_APPROVED: 'SEC-004',
    ACCESS_DENIED:  'SEC-005',
    OVERDUE_REVIEW: 'SEC-006',
    STALE_REQUEST:  'SEC-007',
    HEARTBEAT:      'SEC-HB'
  }
};

// ============================================
// LOGGING - Write to Health Monitor tabs
// ============================================

function logActivity(automationId, functionName, status, details) {
  try {
    var healthSheet = SpreadsheetApp.openById(SEC_TRACKER.HEALTH_SHEET_ID);

    // Log to Automation Log
    var logTab = healthSheet.getSheetByName('Automation Log');
    if (logTab) {
      logTab.insertRowAfter(1);
      logTab.getRange(2, 1, 1, 6).setValues([[
        new Date(), automationId, SEC_TRACKER.SCRIPT_NAME, functionName, status, details
      ]]);
    }

    // Update Last Run in Registered Automations
    var regTab = healthSheet.getSheetByName('Registered Automations');
    if (regTab) {
      var data = regTab.getDataRange().getValues();
      for (var i = 1; i < data.length; i++) {
        if (data[i][0] === automationId) {
          regTab.getRange(i + 1, 11).setValue(new Date()); // Last Run
          regTab.getRange(i + 1, 12).setValue(status);     // Last Status
          break;
        }
      }
    }
  } catch (e) {
    Logger.log('Log error: ' + e.message);
  }
}

// ============================================
// SLACK — uses SLACK_WEBHOOK_URL from Script Properties
// ============================================

function secSendNotification(emoji, title, fields, buttonText, buttonUrl) {
  var fieldText = '';
  for (var f = 0; f < fields.length; f++) {
    fieldText += '*' + fields[f].label + ':* ' + fields[f].value + '\n';
  }
  return secSendSlackMessage({
    blocks: [
      { type: 'section', text: { type: 'mrkdwn', text: emoji + ' *' + title + '*' } },
      { type: 'section', text: { type: 'mrkdwn', text: fieldText.trim() } },
      { type: 'actions', elements: [{ type: 'button',
          text: { type: 'plain_text', text: buttonText, emoji: true },
          url: buttonUrl }] }
    ]
  });
}

function secSendSlackMessage(payload) {
  try {
    var url = getWebhookUrl(); // defined in api-key-expiration-notifier.js
    if (!url) { Logger.log('SEC: SLACK_WEBHOOK_URL not set'); return false; }
    var response = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
    return response.getResponseCode() === 200;
  } catch (e) {
    Logger.log('SEC Slack error: ' + e.message);
    return false;
  }
}

// ============================================
// TRIGGERS
// ============================================

function secSetupTriggers() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    var fn = triggers[i].getHandlerFunction();
    if (fn === 'onSheetEdit' || fn === 'dailyChecks') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }

  ScriptApp.newTrigger('onSheetEdit').forSpreadsheet(SEC_TRACKER.SHEET_ID).onEdit().create();
  ScriptApp.newTrigger('dailyChecks').timeBased().atHour(9).everyDays(1).create();

  logActivity(SEC_TRACKER.AUTO_IDS.HEARTBEAT, 'secSetupTriggers', 'Success', 'Triggers configured');
  secSendNotification('', 'IT Security Tracker Configured',
    [{ label: 'onEdit', value: 'Active' }, { label: 'Daily', value: '9 AM' }],
    'Open Sheet', SEC_TRACKER.SHEET_URL);
  try { SpreadsheetApp.getActive().toast('Triggers set!', 'Done', 5); } catch (e) {}
}

// ============================================
// MAIN TRIGGER
// ============================================

function onSheetEdit(e) {
  if (!e || !e.range || e.range.getRow() === 1) return;

  var sheet     = e.source.getActiveSheet();
  var sheetName = sheet.getName();
  var row       = e.range.getRow();
  var col       = e.range.getColumn();

  try {
    if (sheetName === SEC_TRACKER.TABS.INCIDENTS && col === SEC_TRACKER.COLUMNS.INC_STATUS) {
      handleIncidentStatus(sheet, row, e.value, e.oldValue);
    }
    if (sheetName === SEC_TRACKER.TABS.ACCESS_REQUESTS && col === SEC_TRACKER.COLUMNS.AR_STATUS) {
      handleAccessRequestStatus(sheet, row, e.value, e.oldValue);
    }
  } catch (error) {
    logActivity('SEC-ERR', 'onSheetEdit', 'Failed', error.message);
  }
}

// ============================================
// HANDLERS
// ============================================

function handleIncidentStatus(sheet, row, newValue, oldValue) {
  var rowData  = sheet.getRange(row, 1, 1, 12).getValues()[0];
  var id       = rowData[0];
  var reportedBy = rowData[2];
  var type     = rowData[3];
  var severity = rowData[4];

  if (newValue === 'Open' && oldValue !== 'Open') {
    var emoji = severity === 'CRITICAL' ? ':rotating_light:' : (severity === 'HIGH' ? ':warning:' : ':memo:');
    var success = secSendNotification(emoji, 'Security Incident Opened',
      [{ label: 'ID',       value: id || 'New' },
       { label: 'Type',     value: type },
       { label: 'Severity', value: severity },
       { label: 'By',       value: reportedBy }],
      'View', SEC_TRACKER.SHEET_URL + '#gid=832516455');
    logActivity(SEC_TRACKER.AUTO_IDS.INCIDENT_OPEN, 'incidentOpened', success ? 'Success' : 'Failed', id);
  }

  if (newValue === 'Closed' && oldValue !== 'Closed') {
    var success2 = secSendNotification(':white_check_mark:', 'Incident Closed',
      [{ label: 'ID', value: id }, { label: 'Type', value: type }],
      'View', SEC_TRACKER.SHEET_URL + '#gid=832516455');
    logActivity(SEC_TRACKER.AUTO_IDS.INCIDENT_CLOSE, 'incidentClosed', success2 ? 'Success' : 'Failed', id);
  }
}

function handleAccessRequestStatus(sheet, row, newValue, oldValue) {
  var rowData    = sheet.getRange(row, 1, 1, 13).getValues()[0];
  var requestedBy = rowData[2];
  var user       = rowData[3];
  var email      = rowData[4];
  var system     = rowData[5];
  var accessLevel = rowData[6];
  var reason     = rowData[7];
  var approvedBy = rowData[9];
  var approvalDate = rowData[10];

  if (newValue === 'Pending' && oldValue !== 'Pending') {
    var success = secSendNotification(':clipboard:', 'New Access Request',
      [{ label: 'User',   value: user },
       { label: 'System', value: system },
       { label: 'Level',  value: accessLevel },
       { label: 'By',     value: requestedBy },
       { label: 'Reason', value: reason || '-' }],
      'Review', SEC_TRACKER.SHEET_URL);
    logActivity(SEC_TRACKER.AUTO_IDS.ACCESS_PENDING, 'accessPending',
      success ? 'Success' : 'Failed', user + ' to ' + system);
  }

  if (newValue === 'Approved' && oldValue !== 'Approved') {
    var success2 = secSendNotification(':white_check_mark:', 'Access Approved',
      [{ label: 'User',   value: user },
       { label: 'System', value: system },
       { label: 'Level',  value: accessLevel }],
      'Register', SEC_TRACKER.SHEET_URL);
    syncToAccessRegister(rowData, sheet, row);
    logActivity(SEC_TRACKER.AUTO_IDS.ACCESS_APPROVED, 'accessApproved',
      success2 ? 'Success' : 'Failed', user + ' to ' + system);
  }

  if (newValue === 'Denied' && oldValue !== 'Denied') {
    var success3 = secSendNotification(':x:', 'Access Denied',
      [{ label: 'User',   value: user },
       { label: 'System', value: system }],
      'View', SEC_TRACKER.SHEET_URL);
    logActivity(SEC_TRACKER.AUTO_IDS.ACCESS_DENIED, 'accessDenied',
      success3 ? 'Success' : 'Failed', user + ' to ' + system);
  }
}

function syncToAccessRegister(requestData, requestSheet, requestRow) {
  try {
    var ss  = SpreadsheetApp.openById(SEC_TRACKER.SHEET_ID);
    var reg = ss.getSheetByName(SEC_TRACKER.TABS.ACCESS_REGISTER);
    if (!reg) return;
    var user        = requestData[3];
    var email       = requestData[4];
    var system      = requestData[5];
    var accessLevel = requestData[6];
    var reason      = requestData[7];
    var approvedBy  = requestData[9];
    var approvalDate = requestData[10];
    reg.appendRow([user, email, system, accessLevel, 'Active',
      approvalDate || new Date(), approvedBy || requestData[2], reason, '', '', '', '']);
    requestSheet.getRange(requestRow, 13).setValue('Yes');
  } catch (e) {
    Logger.log('Sync error: ' + e.message);
  }
}

// ============================================
// DAILY CHECKS
// ============================================

function dailyChecks() {
  var start = new Date();
  try {
    checkOverdueReviews();
    checkPendingRequests();
    logActivity(SEC_TRACKER.AUTO_IDS.HEARTBEAT, 'dailyChecks',
      'Success', 'Duration: ' + (new Date() - start) + 'ms');
  } catch (e) {
    logActivity(SEC_TRACKER.AUTO_IDS.HEARTBEAT, 'dailyChecks', 'Failed', e.message);
    secSendNotification(':x:', 'Daily Checks Failed',
      [{ label: 'Error', value: e.message }],
      'Logs', 'https://docs.google.com/spreadsheets/d/' + SEC_TRACKER.HEALTH_SHEET_ID);
  }
}

function checkOverdueReviews() {
  var ss    = SpreadsheetApp.openById(SEC_TRACKER.SHEET_ID);
  var sheet = ss.getSheetByName(SEC_TRACKER.TABS.REVIEWS);
  if (!sheet) return;

  var data  = sheet.getDataRange().getValues();
  var today = new Date();
  today.setHours(0, 0, 0, 0);

  for (var i = 1; i < data.length; i++) {
    if (data[i][10] && new Date(data[i][10]) < today) {
      var days = Math.floor((today - new Date(data[i][10])) / 86400000);
      secSendNotification(':warning:', 'Review Overdue',
        [{ label: 'ID',   value: data[i][0] },
         { label: 'Days', value: days.toString() }],
        'Reviews', SEC_TRACKER.SHEET_URL);
      logActivity(SEC_TRACKER.AUTO_IDS.OVERDUE_REVIEW, 'overdueReview', 'Success', data[i][0]);
      break;
    }
  }
}

function checkPendingRequests() {
  var ss    = SpreadsheetApp.openById(SEC_TRACKER.SHEET_ID);
  var sheet = ss.getSheetByName(SEC_TRACKER.TABS.ACCESS_REQUESTS);
  if (!sheet) return;

  var data   = sheet.getDataRange().getValues();
  var cutoff = new Date(Date.now() - 86400000);
  var stale  = [];

  for (var i = 1; i < data.length; i++) {
    if (data[i][8] === 'Pending' && data[i][1] && new Date(data[i][1]) < cutoff) {
      stale.push({ user: data[i][3], system: data[i][5] });
    }
  }

  if (stale.length > 0) {
    var fields = [];
    for (var s = 0; s < stale.length; s++) {
      fields.push({ label: stale[s].user, value: stale[s].system });
    }
    secSendNotification(':hourglass:', 'Pending > 24h', fields, 'Review', SEC_TRACKER.SHEET_URL);
    logActivity(SEC_TRACKER.AUTO_IDS.STALE_REQUEST, 'staleRequests',
      'Success', stale.length + ' requests');
  }
}

// ============================================
// TESTS
// ============================================

function secTestSlack() {
  secSendNotification(':white_check_mark:', 'Connection Test',
    [{ label: 'Status', value: 'OK' },
     { label: 'Time',   value: new Date().toLocaleString() }],
    'Open Sheet', SEC_TRACKER.SHEET_URL);
  try { SpreadsheetApp.getActive().toast('Sent!', 'Test', 3); } catch (e) {}
}

function secTestHeartbeat() {
  logActivity(SEC_TRACKER.AUTO_IDS.HEARTBEAT, 'manualTest', 'Success', 'Manual heartbeat');
  try { SpreadsheetApp.getActive().toast('Heartbeat logged!', 'Test', 3); } catch (e) {}
}

function secTestFailure() {
  logActivity(SEC_TRACKER.AUTO_IDS.HEARTBEAT, 'manualTest', 'Failed', 'Simulated failure');
  try { SpreadsheetApp.getActive().toast('Failure logged!', 'Test', 3); } catch (e) {}
}
