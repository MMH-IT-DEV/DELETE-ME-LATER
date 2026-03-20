/**
 * Systems Health Tracker - Daily Health Check + Utilities
 * 2026_Systems-Health-Tracker
 *
 * Works alongside api-key-expiration-notifier.gs in the same project.
 * This file handles: menu, triggers, daily checks, cleanup, standardize, styling.
 * The other file handles: webhook endpoints, heartbeat/error receiving, standalone checks.
 *
 * SETUP:
 * 1. Open this sheet > Extensions > Apps Script
 * 2. Paste both files into the project
 * 3. Project Settings > Script Properties > add:
 *    SLACK_WEBHOOK_URL = your webhook
 * 4. Run setupTriggers() once from the menu
 */

var HB_CONFIG = {
  WEBHOOK_PROPERTY: 'SLACK_WEBHOOK_URL',
  CHANNEL: '#it-support',
  REGISTRY_TAB: 'System Registry',
  LOG_TAB: 'Heartbeat Log',
  EXPIRY_WARN_DAYS: 30,
  TIMEOUT_MS: 10000
};

var REG_COLS = {
  NAME:       1,   // A  System Name
  DESC:       2,   // B  Description
  TYPE:       3,   // C  Type (FLOW / SCRIPT / BOT / GOOGLE SHEET)
  PLATFORM:   4,   // D  Platform
  GMP:        5,   // E  GMP Critical
  FREQUENCY:  6,   // F  Run Frequency
  SOP_ID:     7,   // G  Maintenance Guide
  AUTH_TYPE:  8,   // H  Auth Type
  AUTH_LOC:   9,   // I  Auth Location
  EXPIRY:     10,  // J  Expiry Date
  DAYS_LEFT:  11,  // K  Days Left
  OWNER:      12,  // L  Owner
  VALIDATED:  13,  // M  Validated
  LAST_VAL:   14,  // N  Last Validated
  HB_METHOD:  15,  // O  Heartbeat Method
  STATUS:     16,  // P  Status
  LAST_HB:    17,  // Q  Last Heartbeat
  LAST_RUN_OK: 18, // R  Last Run OK
  NOTES:      19   // S  Notes
};

var NUM_REG_COLS = 19; // must match sheet-setup.js NUM_COLS

var LOG_COLS = {
  TIMESTAMP: 1,  // A
  SYSTEM: 2,     // B
  STATUS: 3,     // C
  ACTION: 4,     // D
  DETAILS: 5     // E
};

// ===================== UNIFIED MENU =====================

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Health Monitor')
    .addItem('Run Full Health Check', 'runHealthCheck')
    .addItem('Test Slack Connection', 'testHealthSlack')
    .addSeparator()
    .addItem('Check Missing Heartbeats', 'checkMissingHeartbeats')
    .addItem('Check Expiring Credentials', 'checkExpiringCredentials')
    .addItem('Send Test Error Alert', 'testSlackAlert')
    .addSeparator()
    .addItem('Format Registry Sheet', 'setupHealthSheet')
    .addItem('Cleanup Heartbeat Log', 'cleanupHeartbeatLog')
    .addSeparator()
    .addItem('Setup Triggers  (8:30 AM check + 7 PM cleanup)', 'setupTriggers')
    .addItem('Remove All Triggers', 'removeAllTriggers')
    .addToUi();
}

// ===================== UNIFIED TRIGGERS =====================

function setupTriggers() {
  removeAllTriggers();
  ScriptApp.newTrigger('runHealthCheck')
    .timeBased().atHour(8).nearMinute(30).everyDays(1).create();
  ScriptApp.newTrigger('cleanupHeartbeatLogAuto')
    .timeBased().atHour(19).nearMinute(0).everyDays(1).create();
  // Keepalive for Slack /it-alert endpoint — prevents cold-start timeouts
  ScriptApp.newTrigger('keepAliveWarmup_')
    .timeBased().everyMinutes(5).create();
  setupHeartbeatLogHeader_();
  Logger.log('Triggers installed: runHealthCheck (8:30 AM), cleanupHeartbeatLogAuto (7 PM), keepAliveWarmup_ (every 5 min)');
  try {
    SpreadsheetApp.getUi().alert('Success',
      'Triggers installed:\n• Daily health check: 8:30 AM\n• Heartbeat Log cleanup: 7:00 PM\n• Slack keepalive: every 5 min (prevents /it-alert timeout)\n\nHeader row added to Heartbeat Log.\n\nTest with: Health Monitor > Run Full Health Check',
      SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (e) { /* running from editor, no UI context */ }
}

function removeAllTriggers() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    ScriptApp.deleteTrigger(triggers[i]);
  }
  Logger.log('All triggers removed');
}

// ===================== MAIN HEALTH CHECK =====================

function runHealthCheck() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var regSheet = ss.getSheetByName(HB_CONFIG.REGISTRY_TAB);
  var logSheet = ss.getSheetByName(HB_CONFIG.LOG_TAB);

  if (!regSheet || !logSheet) {
    Logger.log('Missing sheets');
    return;
  }

  var lastRow = regSheet.getLastRow();
  if (lastRow <= 1) return;

  var data = regSheet.getRange(2, 1, lastRow - 1, NUM_REG_COLS).getValues();
  var now = new Date();
  var tz = Session.getScriptTimeZone();
  var timestamp = Utilities.formatDate(now, tz, 'M/d/yyyy HH:mm:ss');
  var standardTs = Utilities.formatDate(now, tz, 'yyyy-MM-dd HH:mm:ss');

  var results = [];
  var failures = [];
  var warnings = [];
  var passed = 0;

  // --- Check each system ---
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var name     = sv(row[REG_COLS.NAME - 1]);
    var type     = sv(row[REG_COLS.TYPE - 1]);
    var authType = sv(row[REG_COLS.AUTH_TYPE - 1]);
    var expiry   = row[REG_COLS.EXPIRY - 1];
    var owner    = sv(row[REG_COLS.OWNER - 1]);

    if (!name) continue;

    var startTime = new Date().getTime();

    // --- URL reachability check (Notes col if it starts with http) ---
    var urlStatus = 'N/A';
    var urlDetail = 'No HTTP URL';
    var notes = sv(row[REG_COLS.NOTES - 1]);
    if (notes && notes.indexOf('http') === 0) {
      var urlResult = checkUrl(notes);
      urlStatus = urlResult.ok ? 'OK' : 'FAIL';
      urlDetail = 'HTTP ' + urlResult.code;
      if (!urlResult.ok) {
        failures.push({name: name, issue: 'URL unreachable (' + urlResult.code + ')', owner: owner});
      }
    }

    // --- Token expiry check ---
    var expiryStatus = 'N/A';
    var expiryDetail = '';
    if (expiry && expiry instanceof Date && !isNaN(expiry)) {
      var daysLeft = Math.floor((expiry - now) / 86400000);
      if (daysLeft < 0) {
        expiryStatus = 'EXPIRED';
        expiryDetail = 'Expired ' + Math.abs(daysLeft) + ' days ago';
        failures.push({name: name, issue: 'Token/key EXPIRED (' + Math.abs(daysLeft) + ' days ago)', owner: owner});
      } else if (daysLeft <= HB_CONFIG.EXPIRY_WARN_DAYS) {
        expiryStatus = 'WARNING';
        expiryDetail = daysLeft + ' days remaining';
        warnings.push({name: name, issue: 'Token/key expires in ' + daysLeft + ' days', owner: owner});
      } else {
        expiryStatus = 'OK';
        expiryDetail = daysLeft + ' days remaining';
      }
    } else if (sv(row[REG_COLS.EXPIRY - 1]) === 'NO EXPIRY' || authType === 'None') {
      expiryStatus = 'OK';
      expiryDetail = 'No expiry';
    }

    // --- Update Days Until Expiry in registry ---
    if (expiry && expiry instanceof Date && !isNaN(expiry)) {
      var daysUntil = Math.floor((expiry - now) / 86400000);
      regSheet.getRange(i + 2, REG_COLS.DAYS_LEFT).setValue(daysUntil);
    }

    var duration = new Date().getTime() - startTime;
    var overallStatus = 'Success';
    if (expiryStatus === 'EXPIRED' || urlStatus === 'FAIL') overallStatus = 'Failed';
    else if (expiryStatus === 'WARNING') overallStatus = 'Warning';

    if (overallStatus === 'Success') passed++;

    // --- Update Status (col L) and Last Heartbeat (col M) in registry ---
    var regStatus = overallStatus === 'Success' ? 'Healthy' : (overallStatus === 'Warning' ? 'Degraded' : 'Down');
    regSheet.getRange(i + 2, REG_COLS.STATUS).setValue(regStatus);
    regSheet.getRange(i + 2, REG_COLS.LAST_HB).setValue(standardTs);
    styleRegStatusCell(regSheet.getRange(i + 2, REG_COLS.STATUS), regStatus);

    var detail = 'Duration: ' + duration + 'ms | URL: ' + urlStatus +
      ' (' + urlDetail + ') | Auth: ' + expiryStatus + ' (' + expiryDetail + ')';

    results.push({
      timestamp: timestamp,
      system: name,
      status: overallStatus,
      action: 'HB: Full Check',
      details: detail
    });
  }

  // --- Slack webhook self-check ---
  var slackOk = checkSlackWebhook();
  results.push({
    timestamp: timestamp,
    system: 'Slack Webhook',
    status: slackOk ? 'Success' : 'Failed',
    action: 'HB: Webhook Check',
    details: slackOk ? 'Webhook responding OK' : 'Webhook UNREACHABLE'
  });
  if (!slackOk) {
    failures.push({name: 'Slack Webhook', issue: 'Webhook not responding', owner: 'IT'});
  } else {
    passed++;
  }

  // --- Write results to Heartbeat Log ---
  var logRows = [];
  for (var r = 0; r < results.length; r++) {
    var res = results[r];
    logRows.push([
      res.timestamp,
      res.system,
      res.status,
      res.action,
      res.details
    ]);
  }

  if (logRows.length > 0) {
    var insertRow = logSheet.getLastRow() + 1;
    logSheet.getRange(insertRow, 1, logRows.length, 5).setValues(logRows);
  }

  // --- Also run missing heartbeat + expiry checks from the other file ---
  try {
    if (typeof checkMissingHeartbeats === 'function') checkMissingHeartbeats();
    if (typeof checkExpiringCredentials === 'function') checkExpiringCredentials();
  } catch (e) {
    Logger.log('Extended checks error: ' + e);
  }

  // --- Send Slack summary ---
  var totalSystems = results.length;
  sendHealthSummary(totalSystems, passed, failures, warnings, ss.getId());

  // --- Show UI summary if run manually ---
  try {
    SpreadsheetApp.getUi().alert('Health Check Complete',
      'Total: ' + totalSystems + '\nPassed: ' + passed +
      '\nWarnings: ' + warnings.length + '\nFailed: ' + failures.length +
      '\n\nCheck #it-support for details.',
      SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (e) {
    Logger.log('Health check complete. Passed: ' + passed + ', Warnings: ' + warnings.length + ', Failed: ' + failures.length);
  }
}

// ===================== URL CHECK =====================

function checkUrl(url) {
  try {
    var response = UrlFetchApp.fetch(url, {
      method: 'get',
      muteHttpExceptions: true,
      followRedirects: true,
      validateHttpsCertificates: false,
      headers: { 'User-Agent': 'IT-HealthCheck/1.0' }
    });
    var code = response.getResponseCode();
    return { ok: (code >= 200 && code < 400), code: code };
  } catch (e) {
    return { ok: false, code: 'ERROR: ' + e.message.substring(0, 50) };
  }
}

// ===================== SLACK WEBHOOK CHECK =====================

function checkSlackWebhook() {
  try {
    var url = PropertiesService.getScriptProperties().getProperty(HB_CONFIG.WEBHOOK_PROPERTY);
    if (!url) return false;

    var resp = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({text: ''}),
      muteHttpExceptions: true
    });
    var code = resp.getResponseCode();
    return (code === 200 || code === 400);
  } catch (e) {
    return false;
  }
}

// ===================== SLACK SUMMARY =====================

function sendHealthSummary(total, passed, failures, warnings, ssId) {
  var url = PropertiesService.getScriptProperties().getProperty(HB_CONFIG.WEBHOOK_PROPERTY);
  if (!url) return;

  var failed = failures.length;
  var warned = warnings.length;
  var link = 'https://docs.google.com/spreadsheets/d/' + ssId + '/edit';

  var emoji = ':white_check_mark:';
  var title = 'All Systems Healthy';
  if (failed > 0) { emoji = ':rotating_light:'; title = failed + ' System(s) FAILING'; }
  else if (warned > 0) { emoji = ':warning:'; title = warned + ' Warning(s)'; }

  var blocks = [
    mkSec(emoji + '  *Daily Health Check - ' + title + '*'),
    {type: 'divider'},
    mkSec('*Total Systems:*  `' + total + '`\n*Passed:*  `' + passed + '`\n*Warnings:*  `' + warned + '`\n*Failed:*  `' + failed + '`')
  ];

  if (failed > 0) {
    var failText = '';
    for (var f = 0; f < failures.length; f++) {
      failText += ':x:  `' + failures[f].name + '` - ' + failures[f].issue + ' (Owner: ' + failures[f].owner + ')\n';
    }
    blocks.push({type: 'divider'});
    blocks.push(mkSec('*Failures:*\n' + failText));
  }

  if (warned > 0) {
    var warnText = '';
    for (var w = 0; w < warnings.length; w++) {
      warnText += ':warning:  `' + warnings[w].name + '` - ' + warnings[w].issue + ' (Owner: ' + warnings[w].owner + ')\n';
    }
    if (failed === 0) blocks.push({type: 'divider'});
    blocks.push(mkSec('*Warnings:*\n' + warnText));
  }

  blocks.push(mkSec('*Full Report:*  <' + link + '|View Health Tracker>'));
  blocks.push({type: 'context', elements: [{type: 'mrkdwn',
    text: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm') + '  |  Daily Health Check'}]});

  try {
    UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({
        channel: HB_CONFIG.CHANNEL,
        username: 'IT Health Monitor',
        icon_emoji: ':heartpulse:',
        text: 'Daily Health Check - ' + title,
        blocks: blocks
      }),
      muteHttpExceptions: true
    });
  } catch (e) {
    Logger.log('Slack send error: ' + e);
  }
}

// ===================== CLEANUP HEARTBEAT LOG =====================

function cleanupHeartbeatLog() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var logSheet = ss.getSheetByName(HB_CONFIG.LOG_TAB);
  if (!logSheet) return;

  var lastRow = logSheet.getLastRow();
  if (lastRow <= 1) {
    SpreadsheetApp.getUi().alert('Info', 'Heartbeat Log is already empty.', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  var rowCount = lastRow - 1;
  logSheet.deleteRows(2, rowCount);
  SpreadsheetApp.flush();

  SpreadsheetApp.getUi().alert('Done',
    'Cleared ' + rowCount + ' rows from Heartbeat Log.\nLog is now empty and ready for fresh checks.',
    SpreadsheetApp.getUi().ButtonSet.OK);
}

// ===================== AUTO CLEANUP — runs daily at 7 PM =====================

/**
 * Deletes all Heartbeat Log rows from before today.
 * Keeps today's data intact. Header row (row 1) is never touched.
 * Triggered automatically at 7 PM — see setupTriggers().
 */
function cleanupHeartbeatLogAuto() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var logSheet = ss.getSheetByName(HB_CONFIG.LOG_TAB);
  if (!logSheet) return;

  var lastRow = logSheet.getLastRow();
  if (lastRow <= 1) { Logger.log('cleanupHeartbeatLogAuto: log already empty'); return; }

  var today = new Date();
  today.setHours(0, 0, 0, 0);

  // Timestamps are in col A (row 2 onward). Find the first row from today.
  var data = logSheet.getRange(2, 1, lastRow - 1, 1).getValues();
  var firstTodayRow = -1;

  for (var i = 0; i < data.length; i++) {
    var ts = data[i][0];
    if (!ts) continue;
    var rowDate = ts instanceof Date ? new Date(ts.getTime()) : new Date(ts.toString());
    if (!isNaN(rowDate) && rowDate >= today) {
      firstTodayRow = i + 2; // 1-based row, offset for header
      break;
    }
  }

  if (firstTodayRow === -1) {
    // No rows from today — delete everything
    logSheet.deleteRows(2, lastRow - 1);
    Logger.log('cleanupHeartbeatLogAuto: deleted all ' + (lastRow - 1) + ' rows (none from today)');
    return;
  }

  if (firstTodayRow > 2) {
    var deleteCount = firstTodayRow - 2;
    logSheet.deleteRows(2, deleteCount);
    Logger.log('cleanupHeartbeatLogAuto: deleted ' + deleteCount + ' rows older than today');
  } else {
    Logger.log('cleanupHeartbeatLogAuto: nothing to delete — all rows are from today');
  }
}

// ===================== HEARTBEAT LOG HEADER SETUP =====================

/**
 * Adds a formatted header row to the Heartbeat Log if one doesn't exist.
 * Safe to call multiple times — skips if header already present.
 * Called by setupTriggers().
 */
function setupHeartbeatLogHeader_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var logSheet = ss.getSheetByName(HB_CONFIG.LOG_TAB);
  if (!logSheet) return;

  // Check if header already set
  var firstCell = logSheet.getLastRow() >= 1 ? logSheet.getRange(1, 1).getValue().toString() : '';
  if (firstCell === 'Timestamp') return;

  // Insert header row at top (pushes existing data down)
  if (logSheet.getLastRow() >= 1) {
    logSheet.insertRowBefore(1);
  }

  var hRange = logSheet.getRange(1, 1, 1, 5);
  hRange.setValues([['Timestamp', 'System', 'Status', 'Action', 'Details']]);
  hRange.setBackground('#263238');
  hRange.setFontColor('#ffffff');
  hRange.setFontWeight('bold');
  hRange.setFontSize(10);

  logSheet.setFrozenRows(1);
  logSheet.setColumnWidth(1, 160);
  logSheet.setColumnWidth(2, 220);
  logSheet.setColumnWidth(3, 90);
  logSheet.setColumnWidth(4, 160);
  logSheet.setColumnWidth(5, 350);

  SpreadsheetApp.flush();
  Logger.log('setupHeartbeatLogHeader_: header row added to Heartbeat Log');
}

// ===================== STANDARDIZE REGISTRY =====================

/**
 * Standardizes Status (col L) and Last Heartbeat (col M) in System Registry.
 * Handles plain text AND emoji-prefixed values from the old V2 script.
 */
function standardizeRegistry() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var regSheet = ss.getSheetByName(HB_CONFIG.REGISTRY_TAB);
  if (!regSheet) return;

  var lastRow = regSheet.getLastRow();
  if (lastRow <= 1) return;

  var tz = Session.getScriptTimeZone();
  var updated = 0;

  for (var r = 2; r <= lastRow; r++) {
    var name = sv(regSheet.getRange(r, REG_COLS.NAME).getValue());
    if (!name) continue;

    // --- Standardize Status (col L) ---
    var statusCell = regSheet.getRange(r, REG_COLS.STATUS);
    var rawStatus = sv(statusCell.getValue()).toLowerCase();

    // Strip any emoji prefixes from V2 format
    rawStatus = rawStatus.replace(/[\u{1F7E2}\u{1F534}\u{1F7E1}\u{26A0}\u{2705}\u{274C}\u{1F6A8}]\s*/gu, '').trim();

    var newStatus = 'Unknown';
    if (rawStatus === 'healthy' || rawStatus === 'ok' || rawStatus === 'up' || rawStatus === 'active' || rawStatus === 'success') {
      newStatus = 'Healthy';
    } else if (rawStatus === 'degraded' || rawStatus === 'warning' || rawStatus === 'slow') {
      newStatus = 'Degraded';
    } else if (rawStatus === 'down' || rawStatus === 'fail' || rawStatus === 'failed' || rawStatus === 'error' || rawStatus === 'offline') {
      newStatus = 'Down';
    } else if (rawStatus === 'unknown' || rawStatus === '') {
      newStatus = 'Unknown';
    }

    statusCell.setValue(newStatus);
    styleRegStatusCell(statusCell, newStatus);

    // --- Standardize Last Heartbeat (col M) ---
    var hbCell = regSheet.getRange(r, REG_COLS.LAST_HB);
    var rawHB = hbCell.getValue();

    if (rawHB) {
      var parsedDate = null;

      if (rawHB instanceof Date && !isNaN(rawHB)) {
        parsedDate = rawHB;
      } else {
        var hbStr = rawHB.toString().trim();
        parsedDate = new Date(hbStr);
      }

      if (parsedDate && !isNaN(parsedDate)) {
        var formatted = Utilities.formatDate(parsedDate, tz, 'yyyy-MM-dd HH:mm:ss');
        hbCell.setValue(formatted);
      } else {
        hbCell.setValue('');
      }
    }

    updated++;
  }

  regSheet.getRange(2, REG_COLS.LAST_HB, lastRow - 1, 1).setNumberFormat('@');
  SpreadsheetApp.flush();

  SpreadsheetApp.getUi().alert('Done',
    'Standardized ' + updated + ' rows.\n\nStatus: Healthy / Degraded / Down / Unknown\nLast Heartbeat: yyyy-MM-dd HH:mm:ss\n\nOld emoji prefixes have been removed.',
    SpreadsheetApp.getUi().ButtonSet.OK);
}

// ===================== STATUS CELL STYLING =====================

// Delegates to applyStatusStyle_ in sheet-setup.js (same GAS project scope)
function styleRegStatusCell(cell, status) {
  applyStatusStyle_(cell, status);
}

// ===================== HELPERS =====================

function mkSec(text) {
  return {type: 'section', text: {type: 'mrkdwn', text: text}};
}

function sv(val) {
  if (val === null || val === undefined || val === '') return '';
  return val.toString().trim();
}

// ===================== TEST =====================

function testHealthSlack() {
  var url = PropertiesService.getScriptProperties().getProperty(HB_CONFIG.WEBHOOK_PROPERTY);
  if (!url) {
    SpreadsheetApp.getUi().alert('Error', 'SLACK_WEBHOOK_URL not set in Script Properties.\n\nGo to: Project Settings > Script Properties > Add:\nKey: SLACK_WEBHOOK_URL\nValue: your webhook URL', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  var blocks = [
    mkSec(':heartpulse:  *IT Health Monitor - Connection Test*'),
    {type: 'divider'},
    mkSec('*Status:*  `Connected`\n*Channel:*  `' + HB_CONFIG.CHANNEL + '`\n*Checks:*  URL reachability, token expiry, missing heartbeats, webhook health'),
    {type: 'context', elements: [{type: 'mrkdwn', text: 'Test  |  ' +
      Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm')}]}
  ];

  try {
    var resp = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({
        channel: HB_CONFIG.CHANNEL,
        username: 'IT Health Monitor',
        icon_emoji: ':heartpulse:',
        text: 'Health Monitor - Connection Test',
        blocks: blocks
      }),
      muteHttpExceptions: true
    });
    var ok = resp.getResponseCode() === 200;
    SpreadsheetApp.getUi().alert(ok ? 'Success' : 'Error',
      ok ? 'Test sent to ' + HB_CONFIG.CHANNEL : 'Failed: HTTP ' + resp.getResponseCode(),
      SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (e) {
    SpreadsheetApp.getUi().alert('Error', e.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}
