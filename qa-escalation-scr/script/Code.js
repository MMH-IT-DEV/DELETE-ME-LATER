/**
 * QA Escalation Automation v2.4
 * Sheet: Product Quality & Safety Log
 * Trigger: onQAEscalationEdit (installable onEdit)
 * Heartbeat: System Health Monitor
 *
 * Slack notifications → slack/qa-notifications.js
 * Self-backup        → backup/backup.js
 */

// ============ CONFIGURATION ============
var CONFIG = {
  QA_SUBTEAM_ID: 'S0ABE9GKD7T',
  // Header names — columns are found dynamically so moving columns won't break the trigger
  COLUMN_HEADERS: {
    ORDER_NUMBER:   'Order #',
    LOT_NUMBER:     'Lot #',
    COMPLAINT:      'Complaint Description',
    PHOTO:          'Photos Received',
    RESOLUTION:     'Resolution',
    QA_ESCALATION:  'QA Escalation',
    QA_SENT:        'QA Sent'
  },
  HEADER_ROW: 10,
  TRIGGER_VALUE: 'Escalate to QA',
  HEALTH_SHEET_ID: '1jnWtdBPzR7DreihCHQASiN7spImRS7HYjTR77Gpfi5w',
  SYSTEM_NAME: 'QA Escalation Bot'
};

function getSlackWebhookUrl_() {
  return PropertiesService.getScriptProperties().getProperty('SLACK_WEBHOOK_URL');
}

// ============ BACKUP CONFIG ============
var BACKUP = {
  NAME: 'QA Escalation Bot',
  TYPE: 'Notifications',
  VERSION: '1.0',
  SYSTEM_NAME: 'QA Escalation Bot - Backup'
};

// ============ MENU ============
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('QA Escalation')
    .addItem('Setup Trigger', 'setupTrigger')
    .addItem('Test Slack', 'testSlackNotification')
    .addSeparator()
    .addItem('💾 Save Backup', 'saveBackup')
    .addItem('📊 Backup Info', 'backupInfo')
    .addToUi();
}

// ============ DYNAMIC COLUMN FINDER ============
/**
 * Reads header row and returns a map of { KEY: columnNumber (1-indexed) }.
 * Returns -1 for any header not found — trigger skips gracefully.
 */
function getColumns_(sheet) {
  var headers = sheet.getRange(CONFIG.HEADER_ROW, 1, 1, sheet.getLastColumn()).getValues()[0];
  var cols = {};
  for (var key in CONFIG.COLUMN_HEADERS) {
    var idx = headers.indexOf(CONFIG.COLUMN_HEADERS[key]);
    cols[key] = idx >= 0 ? idx + 1 : -1;
  }
  return cols;
}

// ============ MAIN TRIGGER ============
function onQAEscalationEdit(e) {
  try {
    if (!e || !e.range) return;
    var sheet = e.range.getSheet();
    var row = e.range.getRow();
    var col = e.range.getColumn();

    var COLS = getColumns_(sheet);

    if (COLS.QA_ESCALATION === -1 || col !== COLS.QA_ESCALATION) return;
    if (row <= CONFIG.HEADER_ROW) return;
    if (e.value !== CONFIG.TRIGGER_VALUE) return;

    if (COLS.QA_SENT === -1) return;
    var qaSentCell = sheet.getRange(row, COLS.QA_SENT);
    var currentValue = qaSentCell.getValue().toString();
    if (currentValue === '\u2713' || currentValue === 'TRUE') return;

    var rowData = {
      orderNumber: COLS.ORDER_NUMBER > 0 ? sheet.getRange(row, COLS.ORDER_NUMBER).getValue() || '-' : '-',
      lotNumber:   COLS.LOT_NUMBER   > 0 ? sheet.getRange(row, COLS.LOT_NUMBER).getValue()   || '-' : '-',
      complaint:   COLS.COMPLAINT    > 0 ? sheet.getRange(row, COLS.COMPLAINT).getValue()    || '-' : '-',
      photo:       COLS.PHOTO        > 0 ? sheet.getRange(row, COLS.PHOTO).getValue()        || '-' : '-',
      resolution:  COLS.RESOLUTION   > 0 ? sheet.getRange(row, COLS.RESOLUTION).getValue()  || '-' : '-'
    };

    var success = sendSlackNotification(rowData, row);
    qaSentCell.setValue(success ? '\u2713' : '\u2717');
    if (success) {
      logHeartbeat(CONFIG.SYSTEM_NAME, '✓ Success', 'Row ' + row + ' escalated');
      try { sendHeartbeatToMonitor_('QA Escalation Automation', 'success', 'Row ' + row + ' escalated'); } catch (hbErr) { Logger.log('HB error: ' + hbErr.message); }
    } else {
      logError(CONFIG.SYSTEM_NAME, 'Slack notification failed for row ' + row);
      try { sendHeartbeatToMonitor_('QA Escalation Automation', 'error', 'Slack notification failed for row ' + row); } catch (hbErr) { Logger.log('HB error: ' + hbErr.message); }
    }

  } catch (error) {
    Logger.log('Error: ' + error.message);
    logError(CONFIG.SYSTEM_NAME, error.message);
    try { sendHeartbeatToMonitor_('QA Escalation Automation', 'error', error.message); } catch (hbErr) { Logger.log('HB error: ' + hbErr.message); }
  }
}

// ============ HEALTH LOGGING ============
function logHeartbeat(systemName, status, details) {
  try {
    var ss = SpreadsheetApp.openById(CONFIG.HEALTH_SHEET_ID);
    var sheet = ss.getSheetByName('Heartbeat Log');
    if (sheet) {
      sheet.appendRow([new Date(), systemName, status, details, 'auto']);
    }
  } catch (error) {
    Logger.log('Heartbeat log error: ' + error.message);
  }
}

function logError(systemName, errorMessage) {
  try {
    var ss = SpreadsheetApp.openById(CONFIG.HEALTH_SHEET_ID);
    var sheet = ss.getSheetByName('Error Log');
    if (sheet) {
      sheet.appendRow([new Date(), systemName, '✗ ERROR', errorMessage, '', '', '', '', '', 'Open']);
    }
    sendErrorSlackAlert(systemName, errorMessage);
  } catch (error) {
    Logger.log('Error log failed: ' + error.message);
  }
}

// ============ SETUP ============
function setupTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'onQAEscalationEdit') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  ScriptApp.newTrigger('onQAEscalationEdit')
    .forSpreadsheet(SpreadsheetApp.getActive())
    .onEdit()
    .create();
  SpreadsheetApp.getUi().alert('✓ Trigger installed!');
}

// ============ TEST ============
function testSlackNotification() {
  var testData = {
    orderNumber: 'TEST-12345',
    lotNumber: 'MS301',
    complaint: 'Test complaint message',
    photo: 'https://example.com/photo',
    resolution: 'Test resolution'
  };
  var success = sendSlackNotification(testData, 99);
  if (success) logHeartbeat(CONFIG.SYSTEM_NAME, '✓ Test', 'Manual test successful');
  SpreadsheetApp.getUi().alert(success ? '✓ Test sent!' : '✗ Test failed');
}

// ============ HEALTH MONITOR HEARTBEAT ============
// Requires Script Property: HEALTH_MONITOR_URL
function sendHeartbeatToMonitor_(systemName, status, details) {
  var url = PropertiesService.getScriptProperties().getProperty('HEALTH_MONITOR_URL');
  if (!url) { Logger.log('sendHeartbeatToMonitor_: HEALTH_MONITOR_URL not set'); return; }
  try {
    UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({ system: systemName, status: status, details: details || '' }),
      muteHttpExceptions: true
    });
  } catch (e) {
    Logger.log('sendHeartbeatToMonitor_ error: ' + e.message);
  }
}
