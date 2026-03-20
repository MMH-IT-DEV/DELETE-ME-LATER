// ============================================
// 08_Logging.gs - LOGGING
// ============================================
// Activity logging → Activity tab (all flows, exec IDs, tree lines)
// Flow detail logging → F1-F5 tabs (tree format, grouped by event)
// Debug logging → Logger.log only (no sheet)
// ============================================

// ============================================
// DEBUG LOGGING — Logger.log only
// ============================================

/**
 * Debug logging — Logger.log only (no sheet write).
 * Preserves same signature for all 85+ call sites.
 */
function logToSheet(eventType, data, result) {
  try {
    var message = buildDebugMessage(eventType, data, result);
    var context = buildDebugContext(data);
    var logLine = '[' + eventType + '] ' + message;
    if (context) logLine += ' | ' + context;
    Logger.log(logLine);
  } catch (e) {
    Logger.log('Log error: ' + e.message);
  }
}

/**
 * Build a clean, actionable debug message from event type and data
 */
function buildDebugMessage(eventType, data, result) {
  var type = (eventType || '').toUpperCase();
  data = data || {};
  result = result || {};

  // F2 errors
  if (type === 'F2_SKU_NOT_FOUND') return 'SKU not found in Katana: ' + (data.sku || '?');
  if (type === 'F2_SITE_NOT_MAPPED') return 'WASP site not mapped to Katana location: ' + (data.site || '?');
  if (type === 'F2_LOCATION_NOT_FOUND') return 'Katana location not found: ' + (data.location || '?');

  // F3 errors
  if (type === 'F3_SYNC_ERROR') return 'WASP sync failed: ' + (data.sku || '') + ' — ' + (data.error || '');
  if (type === 'F3_POLL_ERROR') return 'Polling failed: ' + (data.error || '');

  // Batch processing
  if (type === 'BATCH_PROCESS_ERROR') return 'Batch processing crashed: ' + (data.error || '');
  if (type === 'BATCH_LOCK_BUSY') return 'Batch lock contention (concurrent webhooks): ' + (data.batchId || '');
  if (type === 'BATCH_ALREADY_PROCESSED') return 'Duplicate batch attempt (already done): ' + (data.batchId || '');
  if (type === 'BATCH_NO_ITEMS') return 'Empty batch — items expired from cache: ' + (data.batchId || '');

  // Pick/shipping errors
  if (type === 'PICK_COMPLETE_FAILED') return 'ShipStation pick completion failed: ' + (data.pickOrderNumber || '');
  if (type === 'KATANA_DELIVER_FAILED') return 'Failed to mark Katana SO delivered: SO ' + (data.soId || '');

  // WASP API errors
  if (type === 'WASP_PICK_ORDER_ERROR') return 'WASP pick order creation failed: ' + (data.orderNumber || '');

  // Handler errors (missing IDs)
  if (type.indexOf('_ERROR') >= 0 && (data.error || result.error)) {
    return (data.error || result.error || 'Unknown error');
  }

  // Default: event type + any error message
  var msg = eventType;
  if (data.error) msg += ': ' + data.error;
  if (result.error) msg += ': ' + result.error;
  return msg;
}

/**
 * Build structured context string for agent troubleshooting
 * Format: "SKU:CAR-2OZ | Qty:5 | Loc:PRODUCTION | MO:12345"
 */
function buildDebugContext(data) {
  if (!data) return '';
  var parts = [];
  if (data.sku || data.AssetTag || data.itemNumber) parts.push('SKU:' + (data.sku || data.AssetTag || data.itemNumber));
  if (data.qty || data.Quantity || data.quantity) parts.push('Qty:' + (data.qty || data.Quantity || data.quantity));
  if (data.location || data.LocationCode) parts.push('Loc:' + (data.location || data.LocationCode));
  if (data.site || data.SiteName || data.siteName) parts.push('Site:' + (data.site || data.SiteName || data.siteName));
  if (data.soId) parts.push('SO:' + data.soId);
  if (data.moId) parts.push('MO:' + data.moId);
  if (data.poId) parts.push('PO:' + data.poId);
  if (data.stNumber) parts.push('ST:' + data.stNumber);
  if (data.orderNumber) parts.push('Order:' + data.orderNumber);
  if (data.pickOrderNumber) parts.push('Pick:' + data.pickOrderNumber);
  if (data.batchId) parts.push('Batch:' + data.batchId);
  if (data.statusCode) parts.push('HTTP:' + data.statusCode);
  return parts.join(' | ');
}

/**
 * Clear debug sheet — no-op (Debug sheet removed)
 */
function clearDebugSheet() {
  // No-op — Debug sheet removed, logging via Logger.log only
}

/**
 * Format a value for display (truncate if too long)
 */
function formatValue(value, maxLength) {
  maxLength = maxLength || 100;
  if (value === null || value === undefined) return '-';
  if (typeof value === 'object') {
    var str = JSON.stringify(value);
    return str.length > maxLength ? str.substring(0, maxLength) + '...' : str;
  }
  var str = String(value);
  return str.length > maxLength ? str.substring(0, maxLength) + '...' : str;
}

// formatLogEntry removed — logToSheet now uses buildDebugMessage/buildDebugContext directly

// ============================================
// BACKWARD COMPATIBILITY — F2 wrappers
// ============================================
// Old F2-specific functions called from 03_WaspCallouts.gs
// These wrap to the universal logActivity() below
// Remove after updating handlers to call logActivity() directly
// ============================================

function logF2Activity(sku, quantity, success, lot, expiry, batchId, transactionId) {
  var status = success ? 'success' : 'failed';
  var details = 'SA  ' + sku + ' x' + Math.abs(quantity);
  var direction = quantity >= 0 ? '→ Katana' : 'from Katana';
  return logActivity('F2', details, status, direction, [{
    sku: sku,
    qty: Math.abs(quantity),
    lot: lot || '',
    success: success
  }]);
}

function logF2BatchHeader(transactionId, itemCount, overallSuccess) {
  // No-op — logActivity handles header+sub-items together
}

function logF2ActivityDetail(sku, quantity, success, lot, expiry, batchId) {
  // No-op — logActivity handles sub-items inline
}

function logF2BatchActivity(itemDetails, overallSuccess, transactionId) {
  var status = overallSuccess ? 'success' : 'failed';
  var details = 'SA  ' + itemDetails.length + ' items';
  var subItems = [];
  for (var i = 0; i < itemDetails.length; i++) {
    var item = itemDetails[i];
    subItems.push({
      sku: item.sku,
      qty: Math.abs(item.quantity),
      lot: item.lot || '',
      success: item.success !== undefined ? item.success : overallSuccess
    });
  }
  return logActivity('F2', details, status, '→ Katana', subItems);
}

function generateTransactionId() {
  return 'WK-' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd-HHmmss');
}

// ============================================
// ACTIVITY TAB LOGGING — UNIVERSAL (ALL FLOWS)
// ============================================
// Activity tab: ID | Time | Flow | Details
// Rows 1-3 are CC headers (frozen), data starts row 4
// Exec IDs: WK-001, WK-002, etc.
// ============================================

/**
 * Flow labels for Activity tab column C
 */
var FLOW_LABELS = {
  'F1': 'F1 Receiving',
  'F2': 'F2 Adjustments',
  'F3': 'F3 Transfers',
  'F4': 'F4 Manufacturing',
  'F5': 'F5 Shipping',
  'F6': 'F6 Amazon FBA',
  'F7': 'F7 Health Check'
};

/**
 * Success status text per flow
 */
var FLOW_STATUS_TEXT = {
  'F1': 'Received',
  'F2': 'Synced',
  'F3': 'Synced',
  'F4': 'Complete',
  'F5': 'Shipped',
  'F6': 'Complete',
  'F7': 'Checked'
};

/**
 * Flow-based header row colors (matches tracker flow colors)
 */
var FLOW_COLORS = {
  'F1': '#cce5ff',   // blue — Receiving
  'F2': '#b2dfdb',   // teal — Adjustments
  'F3': '#fce4d6',   // peach — Transfers
  'F4': '#f0e6f6',   // purple (bright) — Manufacturing
  'F5': '#d1ecf1',   // cyan — Shipping
  'F6': '#ffe0b2',   // orange — Staging
  'F7': '#e8eaf6'    // indigo — Health Check
};

/**
 * Get Activity sheet (assumes CC 3-row header already exists)
 */
function getActivitySheet() {
  var ss = SpreadsheetApp.openById(CONFIG.DEBUG_SHEET_ID);
  var activitySheet = ss.getSheetByName('Activity');
  if (!activitySheet) {
    activitySheet = ss.insertSheet('Activity');
    activitySheet.appendRow(['ACTIVITY LOG', '', '', 'Katana-WASP Inventory Sync', '', '']);
    activitySheet.appendRow(['All Flows — Chronological', '', '', '', '', '']);
    activitySheet.appendRow(['ID', 'Time', 'Flow', 'Details', 'Status', 'Error', 'Retry']);
    activitySheet.setFrozenRows(3);
    activitySheet.setColumnWidth(7, 50);
  }
  return activitySheet;
}

function getActivityStatusText_(flow, status) {
  if (status === 'failed') return 'Failed';
  if (status === 'partial') return 'Partial';
  if (status === 'skipped') return 'Skipped';
  if (status === 'reverted') return 'Reverted';
  return FLOW_STATUS_TEXT[flow] || 'Complete';
}

function joinActivitySegments_(segments) {
  var parts = [];
  for (var i = 0; i < (segments || []).length; i++) {
    var value = String(segments[i] || '').trim();
    if (value) parts.push(value);
  }
  return parts.join('  ');
}

function getActivityDisplayLocation_(location) {
  var text = String(location || '').trim();
  if (text === 'AMAZON-FBA-USA') return 'AMAZON-FBA';
  return text;
}

function getActivityDisplaySite_(siteName) {
  return String(siteName || '').trim();
}

function normalizeActivityScopeToken_(value) {
  return String(value || '').trim().toLowerCase();
}

function isSameActivityScope_(left, right) {
  var a = normalizeActivityScopeToken_(left);
  var b = normalizeActivityScopeToken_(right);
  return !!(a && b && a === b);
}

function buildActivityCountSummary_(count, singular, plural, verb) {
  var qty = Number(count || 0);
  var noun = qty === 1 ? singular : plural;
  return joinActivitySegments_([qty + ' ' + noun, verb]);
}

function buildActivityBatchMeta_(lot, expiry) {
  var parts = [];
  if (lot) parts.push('lot:' + lot);
  if (expiry) parts.push('exp:' + normalizeBusinessDate_(expiry));
  return joinActivitySegments_(parts);
}

function buildActivityActionText_(action, lot, expiry) {
  return joinActivitySegments_([buildActivityBatchMeta_(lot, expiry), action]);
}

function buildActivityLocationMeta_(siteName, location) {
  var loc = getActivityDisplayLocation_(location);
  var site = getActivityDisplaySite_(siteName);
  if (!loc) return '';
  if (site && isSameActivityScope_(site, loc)) return '';
  return 'bin:' + loc;
}

function buildActivityCompactMeta_(siteName, location, lot, expiry, extraSegments) {
  var segments = [];
  var locationMeta = buildActivityLocationMeta_(siteName, location);
  var batchMeta = buildActivityBatchMeta_(lot, expiry);
  if (batchMeta) segments.push(batchMeta);
  if (locationMeta) segments.push(locationMeta);
  for (var i = 0; i < (extraSegments || []).length; i++) {
    var extra = String(extraSegments[i] || '').trim();
    if (!extra || /^batch_id:/i.test(extra)) continue;
    if (extra) segments.push(extra);
  }
  return joinActivitySegments_(segments);
}

function buildActivitySourceActionContext_(sourceLabel, actionLabel, siteName) {
  var parts = [];
  var source = String(sourceLabel || '').trim();
  var action = String(actionLabel || '').trim();
  var site = getActivityDisplaySite_(siteName);

  if (source) parts.push(source);
  if (action && site) action += ' @ ' + site;
  if (action) parts.push(action);

  return parts.join(' | ');
}

function buildActivityTransferContext_(sourceLabel, actionLabel, fromSite, toSite) {
  var fromLabel = getActivityDisplaySite_(fromSite) || getActivityDisplayLocation_(fromSite) || 'mixed-source';
  var toLabel = getActivityDisplaySite_(toSite) || getActivityDisplayLocation_(toSite) || 'mixed-dest';
  return buildActivitySourceActionContext_(sourceLabel, actionLabel + ' ' + fromLabel + ' -> ' + toLabel, '');
}

function buildActivitySiteSummary_(values) {
  var seen = {};
  for (var i = 0; i < (values || []).length; i++) {
    var site = getActivityDisplaySite_(values[i]);
    if (site) seen[site] = true;
  }
  var sites = Object.keys(seen);
  if (sites.length === 1) return sites[0];
  if (sites.length > 1) return 'multi-site';
  return '';
}

function extractCanonicalActivityRef_(rawRef, fallbackPrefix, fallbackId) {
  var text = String(rawRef || '').trim();
  var match = text.match(/(#\d+|PO-\d+|ST-\d+|MO-\d+|SA-\d+|SO-[A-Za-z0-9-]+)/);
  if (match) return match[1];
  if (fallbackPrefix && fallbackId !== undefined && fallbackId !== null && fallbackId !== '') {
    return fallbackPrefix + fallbackId;
  }
  return text;
}

function buildActivityHeaderLine_(details, context) {
  var left = String(details || '').trim();
  var right = String(context || '').trim();
  if (!left) return right;
  if (!right) return left;
  return left + ' | ' + right;
}

function deriveActivityHeaderError_(subItems, headerError) {
  return headerError || '';
}

function runWithScriptWriteLock_(label, fn, waitMs) {
  var lock = LockService.getScriptLock();
  lock.waitLock(waitMs || 20000);
  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
}

function appendActivityBlockToSheet_(activitySheet, flow, details, status, context, subItems, linkInfo, preExecId, headerError) {
  var execId = preExecId || getNextExecId(activitySheet);
  var now = new Date();
  var timeStr = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');

  var flowLabel = FLOW_LABELS[flow] || flow;
  var statusText = getActivityStatusText_(flow, status);
  var headerLine = buildActivityHeaderLine_(details, context);
  var headerErrorText = deriveActivityHeaderError_(subItems, headerError);

  var headerRow = activitySheet.getLastRow() + 1;
  activitySheet.getRange(headerRow, 1, 1, 6).setValues([
    [execId, timeStr, flowLabel, headerLine, statusText, headerErrorText || '']
  ]);

  var actHdrBg = FLOW_COLORS[flow] || null;
  if (actHdrBg) activitySheet.getRange(headerRow, 1, 1, 6).setBackground(actHdrBg);

  if (linkInfo && linkInfo.text && linkInfo.url) {
    var cell = activitySheet.getRange(headerRow, 4);
    var startIdx = headerLine.indexOf(linkInfo.text);
    if (startIdx >= 0) {
      var richText = SpreadsheetApp.newRichTextValue()
        .setText(headerLine)
        .setLinkUrl(startIdx, startIdx + linkInfo.text.length, linkInfo.url)
        .build();
      cell.setRichTextValue(richText);
    }
  }

  if (subItems && subItems.length > 0) {
    for (var i = 0; i < subItems.length; i++) {
      var item = subItems[i];
      var isLast = (i === subItems.length - 1);

      var subLine;
      if (item.nested) {
        var nextNested = (i + 1 < subItems.length) && subItems[i + 1].nested;
        var nestedLine = nextNested ? '├─' : '└─';
        subLine = '    │   ' + nestedLine + ' x' + item.qty + (item.uom ? ' ' + item.uom : '');
        if (item.action) subLine += '  ' + item.action;
      } else if (item.isParent) {
        var parentTreeLine = isLast ? '└─' : '├─';
        subLine = '    ' + parentTreeLine + ' ' + item.sku;
        if (item.qty) subLine += ' x' + item.qty + (item.uom ? ' ' + item.uom : '');
        if (item.action) subLine += '  ' + item.action;
        if (item.batchCount) subLine += ' (' + item.batchCount + ' batches)';
      } else {
        var treeLine = isLast ? '└─' : '├─';
        subLine = '    ' + treeLine + ' ' + item.sku;
        if (item.qty) subLine += ' x' + item.qty + (item.uom ? ' ' + item.uom : '');
        if (item.action) subLine += '  ' + item.action;
      }

      var subStatus = item.status || (item.success !== false ? '' : 'Failed');
      var subError = item.error ? cleanErrorMessage(item.error) : '';
      if (String(subStatus || '').toLowerCase() === 'skipped') subError = '';

      var subRow = activitySheet.getLastRow() + 1;
      activitySheet.getRange(subRow, 1, 1, 6).setValues([['', '', '', subLine, subStatus, subError]]);

      var actSubStat = String(subStatus || '').toLowerCase();
      if (actSubStat === 'skipped') {
        activitySheet.getRange(subRow, 4).setBackground('#fff8e1');
      } else if (subError) {
        activitySheet.getRange(subRow, 4).setBackground('#ffebee');
      } else {
        activitySheet.getRange(subRow, 4).setBackground('#e8f5e9');
      }

      if (actSubStat === 'failed') {
        activitySheet.getRange(subRow, 5).setBackground('#f8d7da');
      } else if (actSubStat === 'skipped') {
        activitySheet.getRange(subRow, 5).setBackground('#fff8e1');
      } else if (subStatus) {
        activitySheet.getRange(subRow, 5).setBackground('#d4edda');
      }

      if (subError) {
        activitySheet.getRange(subRow, 6).setBackground('#fff0f0');
      }

      if (item.qtyColor && item.qty) {
        var qtyStr = 'x' + item.qty;
        var qtyStart = subLine.indexOf(qtyStr);
        if (qtyStart >= 0) {
          var qtyEnd = qtyStart + qtyStr.length;
          var qtyHex = item.qtyColor === 'green' ? '#008000' : item.qtyColor === 'grey' ? '#999999' : '#cc0000';
          var qtyStyle = SpreadsheetApp.newTextStyle().setForegroundColor(qtyHex).build();
          var qtyRtv = SpreadsheetApp.newRichTextValue()
            .setText(subLine)
            .setTextStyle(qtyStart, qtyEnd, qtyStyle)
            .build();
          activitySheet.getRange(subRow, 4).setRichTextValue(qtyRtv);
        }
      }
    }
  }

  return execId;
}

/**
 * Get next execution ID (WK-XXX format, auto-incrementing)
 * Uses ScriptProperties as atomic counter with document lock to prevent duplicates.
 * IMPORTANT: Uses getDocumentLock() (not getScriptLock()) to avoid releasing the
 * caller's script lock when this function is called from within doPost() or
 * processWebhookQueue() which already hold the script lock.
 * Bootstraps from Activity sheet on first use.
 */
function getNextExecId(activitySheet) {
  var lock = LockService.getUserLock();
  try { lock.waitLock(10000); } catch (e) {
    // Lock timeout — use timestamp-based fallback to guarantee uniqueness
    return 'WK-T' + (new Date().getTime() % 100000);
  }

  try {
    var props = PropertiesService.getScriptProperties();
    var counter = parseInt(props.getProperty('EXEC_ID_COUNTER') || '0', 10);

    // Bootstrap: if counter is 0, seed from the sheet's last exec ID
    if (counter === 0 && activitySheet) {
      var lastRow = activitySheet.getLastRow();
      if (lastRow > 3) {
        var bRow = lastRow;
        while (bRow > 3) {
          var val = activitySheet.getRange(bRow, 1).getValue();
          if (val && String(val).indexOf('WK-') === 0) {
            var sheetNum = parseInt(String(val).replace('WK-', ''), 10);
            if (!isNaN(sheetNum) && sheetNum > counter) counter = sheetNum;
            break;
          }
          bRow--;
        }
      }
    }

    var next = counter + 1;
    props.setProperty('EXEC_ID_COUNTER', String(next));

    if (next < 10) return 'WK-00' + next;
    if (next < 100) return 'WK-0' + next;
    return 'WK-' + next;
  } finally {
    lock.releaseLock();
  }
}

/**
 * Universal Activity Logger — works for ALL flows (F1-F6)
 * Writes to Activity tab with 6 columns: ID | Time | Flow | Details | Status | Error
 * LOG_ICONS removed — status column + conditional formatting handle visual cues.
 *
 * @param {string} flow - 'F1', 'F2', 'F3', 'F4', 'F5', or 'F6'
 * @param {string} details - Summary line, e.g. "MO-7177  FGJ-MS-1 x1733"
 * @param {string} status - 'success', 'failed', 'partial', or 'skipped'
 * @param {string} context - Direction/location, e.g. "→ PROD-RECEIVING" or "WASP → Katana @ SHIPPING-DOCK"
 * @param {Array} subItems - Array of {sku, qty, success, error, action, status} or null
 * @param {Object} linkInfo - Optional {text: '#90590', url: 'https://...'} for clickable ref
 * @param {string} preExecId - Optional pre-allocated exec ID
 * @param {string} headerError - Optional error/warning summary for column F
 * @return {string} execId - The generated execution ID (WK-XXX)
 */
function logActivity(flow, details, status, context, subItems, linkInfo, preExecId, headerError) {
  try {
    return runWithScriptWriteLock_('logActivity', function() {
      var activitySheet = getActivitySheet();
      return appendActivityBlockToSheet_(activitySheet, flow, details, status, context, subItems, linkInfo, preExecId, headerError);
    });

  } catch (e) {
    Logger.log('logActivity error: ' + e.message);
    return 'WK-ERR';
  }
}

/**
 * Clean error messages — map WASP error codes + strip prefixes
 * Returns short, human-readable error text for the Error column.
 * AI reads Error + Details columns together for diagnosis.
 */
function cleanErrorMessage(msg) {
  if (!msg) return '';
  var str = String(msg);

  // Map WASP error codes to human-readable messages
  if (str.indexOf('-46002') >= 0) return 'Insufficient qty at location';
  if (str.indexOf('-57009') >= 0) return 'Location not found in WASP';
  if (str.indexOf('-57041') >= 0) return 'Lot/DateCode required by WASP';
  if (str.indexOf('-70010') >= 0) return 'Duplicate order number';

  // Pattern-match common WASP messages
  var lower = str.toLowerCase();
  if (lower.indexOf('date code is missing') >= 0) return 'Missing expiry date';
  if (lower.indexOf('insufficient') >= 0 || lower.indexOf('not enough') >= 0) return 'Insufficient qty at location';
  if (lower.indexOf('location') >= 0 && lower.indexOf('not found') >= 0) return 'Location not found in WASP';
  if (lower.indexOf('item') >= 0 && lower.indexOf('not found') >= 0) return 'Item not in WASP';
  if (lower.indexOf('lot') >= 0 && lower.indexOf('not found') >= 0) return 'Lot not found at location';

  // Strip "ItemNumber: XXX is fail, message: " prefix
  var match = str.match(/message:\s*(.+)/i);
  if (match) str = match[1].trim();

  // Truncate long error messages for readability
  return str.length > 200 ? str.substring(0, 200) + '...' : str;
}

// ============================================
// FLOW DETAIL TAB LOGGING — TREE FORMAT (ALL FLOWS)
// ============================================
// 6 columns: Exec ID | Time | Ref# | Details | Status | Error
// Header row + tree-line sub-items, grouped by event
// Rows 1-3 are CC headers (frozen), data starts row 4
// ============================================

/**
 * Flow tab names
 */
var FLOW_TAB_NAMES = {
  'F1': 'F1 Receiving',
  'F2': 'F2 Adjustments',
  'F3': 'F3 Transfers',
  'F4': 'F4 Manufacturing',
  'F5': 'F5 Shipping'
};

/**
 * Universal Flow Detail Logger — writes grouped tree-format rows to flow tab
 * 6 columns: Exec ID | Time | Ref# | Details | Status | Error
 *
 * @param {string} flow - 'F1', 'F2', 'F3', 'F4', or 'F5'
 * @param {string} execId - Execution ID from logActivity() (WK-XXX)
 * @param {Object} header - Header row: {ref, detail, status, error, linkText, linkUrl}
 * @param {Array} subItems - Sub-items: [{sku, qty, detail, status, error}] or null
 */
function logFlowDetail(flow, execId, header, subItems) {
  try {
    return runWithScriptWriteLock_('logFlowDetail', function() {
    var tabName = FLOW_TAB_NAMES[flow];
    if (!tabName) return;

    var ss = SpreadsheetApp.openById(CONFIG.DEBUG_SHEET_ID);
    var sheet = ss.getSheetByName(tabName);
    if (!sheet) return;

    var timeStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');

    // Write header row
    var headerError = header.error ? cleanErrorMessage(String(header.error)) : '';
    var refText = header.ref || '';
    var rowNum = sheet.getLastRow() + 1;
    sheet.getRange(rowNum, 1, 1, 6).setValues([[execId, timeStr, refText, header.detail || '', header.status || '', headerError]]);

    // Color header row by flow
    var fdFlowColors = { 'F1': '#cce5ff', 'F2': '#b2dfdb', 'F3': '#fce4d6', 'F4': '#f0e6f6', 'F5': '#d1ecf1', 'F6': '#ffe0b2' };
    var fdHdrBg = fdFlowColors[flow] || null;
    if (fdHdrBg) sheet.getRange(rowNum, 1, 1, 6).setBackground(fdHdrBg);

    // Apply clickable link on Ref# column
    if (header.linkText && header.linkUrl && refText) {
      var cell = sheet.getRange(rowNum, 3);
      var richText = SpreadsheetApp.newRichTextValue()
        .setText(refText)
        .setLinkUrl(0, refText.length, header.linkUrl)
        .build();
      cell.setRichTextValue(richText);
    }

    // Write sub-item rows with tree lines
    // Supports nested batch sub-rows: item.nested=true renders deeper indent
    if (subItems && subItems.length > 0) {
      for (var i = 0; i < subItems.length; i++) {
        var item = subItems[i];
        var isLast = (i === subItems.length - 1);

        var fdIcon = '';

        var subDetail;
        if (item.nested) {
          // Nested batch sub-row: "│   ├─ x{qty}  lot:{lot}  exp:{exp}"
          var nextNested = (i + 1 < subItems.length) && subItems[i + 1].nested;
          var nestedLine = nextNested ? '├─' : '└─';
          subDetail = '│   ' + nestedLine + ' x' + item.qty;
          if (item.detail) subDetail += '  ' + item.detail;
        } else if (item.isParent) {
          // Parent multi-batch row: "├─ SKU x{qty}  detail (N batches)"
          var fdTreeLineP = isLast ? '└─' : '├─';
          subDetail = fdTreeLineP + ' ' + fdIcon + (item.sku || '');
          if (item.qty) subDetail += ' x' + item.qty;
          if (item.detail) subDetail += '  ' + item.detail;
          if (item.batchCount) subDetail += ' (' + item.batchCount + ' batches)';
        } else {
          // Standard sub-item
          var fdTreeLine = isLast ? '└─' : '├─';
          subDetail = fdTreeLine + ' ' + fdIcon + (item.sku || '');
          if (item.qty) subDetail += ' x' + item.qty;
          if (item.detail) subDetail += '  ' + item.detail;
        }

        if (item.uom && item.qty) {
          subDetail = subDetail.replace('x' + item.qty, 'x' + item.qty + ' ' + item.uom);
        }

        var subError = item.error ? cleanErrorMessage(String(item.error)) : '';
        var subStatusText = item.status || '';

        var subRow = sheet.getLastRow() + 1;
        sheet.getRange(subRow, 1, 1, 6).setValues([['', '', '', subDetail, subStatusText, subError]]);

        // Color sub-item detail cell (col D) by outcome
        var fdSubStat = subStatusText.toLowerCase();
        if (fdSubStat === 'skipped') {
          sheet.getRange(subRow, 4).setBackground('#fff8e1');
        } else if (subError) {
          sheet.getRange(subRow, 4).setBackground('#ffebee');
        } else {
          sheet.getRange(subRow, 4).setBackground('#e8f5e9');
        }

        // Color sub-item status cell (col E)
        if (fdSubStat === 'failed') {
          sheet.getRange(subRow, 5).setBackground('#f8d7da');
        } else if (fdSubStat === 'skipped') {
          sheet.getRange(subRow, 5).setBackground('#fff8e1');
        } else if (subStatusText) {
          sheet.getRange(subRow, 5).setBackground('#d4edda');
        }

        // Color sub-item error cell (col F)
        if (subError) {
          sheet.getRange(subRow, 6).setBackground('#fff0f0');
        }

        // Color qty text green (additions) or red (removals) via RichTextValue
        if (item.qtyColor && item.qty) {
          var fdQtyStr = 'x' + item.qty;
          var fdQtyStart = subDetail.indexOf(fdQtyStr);
          if (fdQtyStart >= 0) {
            var fdQtyEnd = fdQtyStart + fdQtyStr.length;
            var fdHex = item.qtyColor === 'green' ? '#008000' : item.qtyColor === 'grey' ? '#999999' : '#cc0000';
            var fdStyle = SpreadsheetApp.newTextStyle().setForegroundColor(fdHex).build();
            var fdRtv = SpreadsheetApp.newRichTextValue()
              .setText(subDetail)
              .setTextStyle(fdQtyStart, fdQtyEnd, fdStyle)
              .build();
            sheet.getRange(subRow, 4).setRichTextValue(fdRtv);
          }
        }
      }
    }
    });

  } catch (e) {
    Logger.log('logFlowDetail error: ' + e.message);
  }
}

// ============================================
// CONDITIONAL FORMATTING SETUP — ACTIVITY TAB
// ============================================
// Run ONCE from Apps Script editor after deploy.
// Replaces setBackground() calls — colors auto-clear when rows are deleted.
// ============================================

/**
 * Set up conditional formatting rules on the Activity tab.
 * Run once from Apps Script editor: setupActivityConditionalFormatting()
 *
 * Header rows (column A has WK-xxx):
 *   F1 Receiving → blue, F2 Adjustments → teal, F3 Transfers → peach,
 *   F4 Manufacturing → purple, F5 Shipping → cyan, F6 Staging → orange
 *
 * Status column formatting (applies to all rows):
 *   Success states → green, Failed → red, Partial → yellow, Skipped → amber
 */
function setupActivityConditionalFormatting() {
  var ss = SpreadsheetApp.openById(CONFIG.DEBUG_SHEET_ID);
  var sheet = ss.getSheetByName('Activity');
  if (!sheet) return;

  var fullRange = sheet.getRange('A4:F5000');
  var statusRange = sheet.getRange('E4:E5000');
  var errorRange = sheet.getRange('F4:F5000');

  // Clear existing conditional formatting on Activity tab
  sheet.clearConditionalFormatRules();

  var rules = [];

  // Header row rules — column A not empty + column C matches flow label
  var flowRules = [
    { label: 'F1 Receiving', color: '#cce5ff' },
    { label: 'F2 Adjustments', color: '#b2dfdb' },
    { label: 'F3 Transfers', color: '#fce4d6' },
    { label: 'F4 Manufacturing', color: '#f0e6f6' },
    { label: 'F5 Shipping', color: '#d1ecf1' },
    { label: 'F6 Staging', color: '#ffe0b2' }
  ];

  for (var i = 0; i < flowRules.length; i++) {
    var fr = flowRules[i];
    rules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied('=AND($A4<>"", $C4="' + fr.label + '")')
        .setBackground(fr.color)
        .setRanges([fullRange])
        .build()
    );
  }

  // Status column (E) coloring — applies to both headers and sub-items
  var statusColors = [
    { text: 'Complete', color: '#d4edda' },
    { text: 'Received', color: '#d4edda' },
    { text: 'Shipped', color: '#d4edda' },
    { text: 'Synced', color: '#d4edda' },
    { text: 'Consumed', color: '#d4edda' },
    { text: 'Produced', color: '#d4edda' },
    { text: 'Added', color: '#d4edda' },
    { text: 'Deducted', color: '#d4edda' },
    { text: 'Returned', color: '#d4edda' },
    { text: 'Staged', color: '#d4edda' },
    { text: 'Picked', color: '#d4edda' },
    { text: 'Edited', color: '#d4edda' },
    { text: 'Failed', color: '#f8d7da' },
    { text: 'Partial', color: '#fff3cd' },
    { text: 'Pending', color: '#fff8e1' },
    { text: 'Skipped', color: '#fff8e1' },
    { text: 'Voided', color: '#e0e0e0' },
    { text: 'Cancelled', color: '#e0e0e0' }
  ];

  for (var s = 0; s < statusColors.length; s++) {
    var sc = statusColors[s];
    rules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo(sc.text)
        .setBackground(sc.color)
        .setRanges([statusRange])
        .build()
    );
  }

  // Error column (F) — light red when non-empty
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND($F4<>"")')
      .setBackground('#fff0f0')
      .setRanges([errorRange])
      .build()
  );

  sheet.setConditionalFormatRules(rules);
}

/**
 * Clear all existing background colors from Activity tab data rows.
 * Run after setupActivityConditionalFormatting() to remove old setBackground colors.
 */
function clearActivityBackgrounds() {
  var ss = SpreadsheetApp.openById(CONFIG.DEBUG_SHEET_ID);
  var sheet = ss.getSheetByName('Activity');
  if (!sheet) return;

  var lastRow = sheet.getLastRow();
  if (lastRow <= 3) return;

  // Clear backgrounds on all data rows (keep header rows 1-3)
  sheet.getRange(4, 1, lastRow - 3, 6).setBackground(null);
}

// ============================================
// FLOW TAB FORMATTING SETUP — ALL TABS
// ============================================
// Run ONCE from Apps Script editor after deploy.
// Sets column widths and freezes header rows on F1-F5 + Activity.
// ============================================

/**
 * Set up column widths and frozen rows on all flow tabs + Activity.
 * Run once from Apps Script editor: setupFlowTabFormatting()
 */
function setupFlowTabFormatting() {
  var ss = SpreadsheetApp.openById(CONFIG.DEBUG_SHEET_ID);

  // F1-F5: 6 columns (Exec ID | Time | Ref# | Details | Status | Error)
  var flowWidths = [70, 50, 100, 400, 80, 300];
  var flowTabs = ['F1 Receiving', 'F2 Adjustments', 'F3 Transfers', 'F4 Manufacturing', 'F5 Shipping'];
  for (var t = 0; t < flowTabs.length; t++) {
    var sheet = ss.getSheetByName(flowTabs[t]);
    if (!sheet) continue;
    for (var c = 0; c < flowWidths.length; c++) {
      sheet.setColumnWidth(c + 1, flowWidths[c]);
    }
    // Update column headers (row 3)
    sheet.getRange(3, 1, 1, 6).setValues([['Exec ID', 'Time', 'Ref#', 'Details', 'Status', 'Error']]);
    sheet.setFrozenRows(3);
  }

  // Clean up flow tabs: delete excess columns G-J, clear old data
  for (var t = 0; t < flowTabs.length; t++) {
    var cleanSheet = ss.getSheetByName(flowTabs[t]);
    if (!cleanSheet) continue;
    var maxCol = cleanSheet.getMaxColumns();
    if (maxCol > 6) {
      cleanSheet.deleteColumns(7, maxCol - 6);
    }
    var lastRow = cleanSheet.getLastRow();
    if (lastRow > 3) {
      cleanSheet.getRange(4, 1, lastRow - 3, 6).clear();
    }
  }

  // Activity: 6 columns (do NOT clear data — master log)
  var actWidths = [70, 50, 120, 400, 80, 200];
  var actSheet = ss.getSheetByName('Activity');
  if (actSheet) {
    for (var c = 0; c < actWidths.length; c++) {
      actSheet.setColumnWidth(c + 1, actWidths[c]);
    }
    // Update headers row 3 to 6 columns
    actSheet.getRange(3, 1, 1, 6).setValues([['ID', 'Time', 'Flow', 'Details', 'Status', 'Error']]);
    actSheet.setFrozenRows(3);
  }
}

// ============================================
// CONDITIONAL FORMATTING SETUP — F1-F5 FLOW TABS
// ============================================
// Run ONCE from Apps Script editor after deploy.
// Replaces setBackground() calls in logFlowDetail() — colors auto-clear
// when rows are deleted or content is cleared.
// ============================================

/**
 * Set up conditional formatting rules on all F1-F5 flow tabs.
 * Run once from Apps Script editor: setupFlowTabsConditionalFormatting()
 *
 * Header rows (col A not empty) → flow color on full row
 * Sub-item detail (col D, col A empty):
 *   Skipped → amber, error present → light red, otherwise → light green
 * Status column (E): same status colors as Activity tab
 * Error column (F): light red when non-empty
 */
function setupFlowTabsConditionalFormatting() {
  var ss = SpreadsheetApp.openById(CONFIG.DEBUG_SHEET_ID);

  // NOTE: F2 Adjustments uses a 13-column flat format — handled by setupF2AdjustmentsFormatting()
  var flowEntries = [
    { tab: 'F1 Receiving',     color: '#cce5ff' },
    { tab: 'F3 Transfers',     color: '#fce4d6' },
    { tab: 'F4 Manufacturing', color: '#f0e6f6' },
    { tab: 'F5 Shipping',      color: '#d1ecf1' }
  ];

  var statusColors = [
    { text: 'Complete',   color: '#d4edda' },
    { text: 'Received',   color: '#d4edda' },
    { text: 'Shipped',    color: '#d4edda' },
    { text: 'Synced',     color: '#d4edda' },
    { text: 'Consumed',   color: '#d4edda' },
    { text: 'Produced',   color: '#d4edda' },
    { text: 'Added',      color: '#d4edda' },
    { text: 'Deducted',   color: '#d4edda' },
    { text: 'Returned',   color: '#d4edda' },
    { text: 'Staged',     color: '#d4edda' },
    { text: 'Picked',     color: '#d4edda' },
    { text: 'Edited',     color: '#d4edda' },
    { text: 'Failed',     color: '#f8d7da' },
    { text: 'Partial',    color: '#fff3cd' },
    { text: 'Pending',    color: '#fff8e1' },
    { text: 'Skipped',    color: '#fff8e1' },
    { text: 'Voided',     color: '#e0e0e0' },
    { text: 'Cancelled',  color: '#e0e0e0' }
  ];

  for (var t = 0; t < flowEntries.length; t++) {
    var entry = flowEntries[t];
    var sheet = ss.getSheetByName(entry.tab);
    if (!sheet) continue;

    sheet.clearConditionalFormatRules();
    var rules = [];

    var fullRange    = sheet.getRange('A4:F5000');
    var detailRange  = sheet.getRange('D4:D5000');
    var statusRange  = sheet.getRange('E4:E5000');
    var errorRange   = sheet.getRange('F4:F5000');

    // Header row (col A not empty) → flow color on all 6 columns
    rules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied('=$A4<>""')
        .setBackground(entry.color)
        .setRanges([fullRange])
        .build()
    );

    // Sub-item detail (col D): skipped → amber (checked before error/success)
    rules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied('=AND($A4="",$E4="Skipped")')
        .setBackground('#fff8e1')
        .setRanges([detailRange])
        .build()
    );

    // Sub-item detail (col D): error present → light red
    rules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied('=AND($A4="",$F4<>"")')
        .setBackground('#ffebee')
        .setRanges([detailRange])
        .build()
    );

    // Sub-item detail (col D): success → light green (fallback)
    rules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied('=AND($A4="",$D4<>"")')
        .setBackground('#e8f5e9')
        .setRanges([detailRange])
        .build()
    );

    // Status column (E) coloring
    for (var s = 0; s < statusColors.length; s++) {
      var sc = statusColors[s];
      rules.push(
        SpreadsheetApp.newConditionalFormatRule()
          .whenTextEqualTo(sc.text)
          .setBackground(sc.color)
          .setRanges([statusRange])
          .build()
      );
    }

    // Error column (F) — light red when non-empty
    rules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied('=$F4<>""')
        .setBackground('#fff0f0')
        .setRanges([errorRange])
        .build()
    );

    sheet.setConditionalFormatRules(rules);
  }
}

/**
 * Clear all existing direct background colors from F1-F5 flow tab data rows.
 * Run after setupFlowTabsConditionalFormatting() to remove old setBackground() colors.
 */
function clearFlowTabBackgrounds() {
  var ss = SpreadsheetApp.openById(CONFIG.DEBUG_SHEET_ID);
  var flowTabs = ['F1 Receiving', 'F2 Adjustments', 'F3 Transfers', 'F4 Manufacturing', 'F5 Shipping'];

  for (var t = 0; t < flowTabs.length; t++) {
    var sheet = ss.getSheetByName(flowTabs[t]);
    if (!sheet) continue;
    var lastRow = sheet.getLastRow();
    if (lastRow <= 3) continue;
    sheet.getRange(4, 1, lastRow - 3, 6).setBackground(null);
  }
}

/**
 * One-shot setup: apply all conditional formatting rules and clear old direct backgrounds.
 * Run ONCE from Apps Script editor after deploying this file.
 * After this runs, deleting or clearing rows will leave no background color behind.
 */
function setupAllConditionalFormatting() {
  setupActivityConditionalFormatting();
  setupFlowTabsConditionalFormatting();
  setupF2AdjustmentsFormatting();
  clearActivityBackgrounds();
  clearFlowTabBackgrounds();
  Logger.log('setupAllConditionalFormatting complete.');
}

// ============================================
// ACTIVITY ROW IN-PLACE UPDATES (ShipStation voids/reprints)
// ============================================

/**
 * Update an existing Activity row when a ShipStation label is voided or re-printed.
 * Finds the original "shipped" Activity row and updates it in-place to show "Voided" or "Shipped" status.
 *
 * @param {string} orderNumber - Shopify order number (e.g. "91500" or "#91500")
 * @param {Array} voidedItems - Array of {sku, qty} items that were voided (null if re-shipping)
 * @return {Object} {success, orderNumber, action, row} or {success: false, error}
 */
function updateActivityRow(orderNumber, voidedItems) {
  try {
    var activitySheet = getActivitySheet();
    var lastRow = activitySheet.getLastRow();

    // Clean order number — remove # prefix if present
    var cleanOrder = String(orderNumber).replace('#', '');
    var searchPattern = '#' + cleanOrder;

    // Search from bottom up for header row with this order number
    var headerRow = -1;
    for (var row = lastRow; row > 3; row--) {
      var colA = activitySheet.getRange(row, 1).getValue();
      if (!colA) continue; // Skip sub-item rows (empty column A)

      var details = String(activitySheet.getRange(row, 4).getValue());
      if (details.indexOf(searchPattern) >= 0) {
        headerRow = row;
        break;
      }
    }

    if (headerRow === -1) {
      return {success: false, error: 'Activity row not found for #' + cleanOrder};
    }

    // Get current status from column E
    var currentStatus = String(activitySheet.getRange(headerRow, 5).getValue());
    var isCurrentlyVoided = currentStatus === 'Voided';

    // Determine action
    var isVoidAction = (voidedItems && voidedItems.length > 0);

    // Only update if state change is needed
    if (isVoidAction && isCurrentlyVoided) {
      return {success: false, error: 'Order #' + cleanOrder + ' is already voided'};
    }
    if (!isVoidAction && !isCurrentlyVoided) {
      return {success: false, error: 'Order #' + cleanOrder + ' is already shipped'};
    }

    // Update header row status (column E)
    if (isVoidAction) {
      activitySheet.getRange(headerRow, 5).setValue('Voided');
    } else {
      activitySheet.getRange(headerRow, 5).setValue('Shipped');
    }

    // Update header row status background color
    var headerStatusCell = activitySheet.getRange(headerRow, 5);
    if (isVoidAction) {
      headerStatusCell.setBackground('#fff3cd'); // amber — voided
    } else {
      headerStatusCell.setBackground('#d4edda'); // green — shipped
    }

    // Update sub-item rows (rows immediately below header where column A is empty)
    var subRow = headerRow + 1;
    while (subRow <= lastRow) {
      var colA = activitySheet.getRange(subRow, 1).getValue();
      if (colA) break; // Hit next header row — stop

      var subDetails = String(activitySheet.getRange(subRow, 4).getValue());
      if (!subDetails) {
        subRow++;
        continue;
      }

      if (isVoidAction) {
        activitySheet.getRange(subRow, 5).setValue('Returned');
      } else {
        activitySheet.getRange(subRow, 5).setValue('Deducted');
      }

      // Flip qty color: red→green (void) or green→red (reship)
      var qtyMatch = subDetails.match(/x(\d+)/);
      if (qtyMatch) {
        var qtyStr = qtyMatch[0];
        var qtyStart = subDetails.indexOf(qtyStr);
        var qtyEnd = qtyStart + qtyStr.length;
        var newColor = isVoidAction ? '#008000' : '#cc0000';
        var qtyStyle = SpreadsheetApp.newTextStyle().setForegroundColor(newColor).build();
        var rtv = SpreadsheetApp.newRichTextValue()
          .setText(subDetails)
          .setTextStyle(qtyStart, qtyEnd, qtyStyle)
          .build();
        activitySheet.getRange(subRow, 4).setRichTextValue(rtv);
      }

      // Flip sub-item background: green→amber (void) or amber→green (reship)
      if (isVoidAction) {
        activitySheet.getRange(subRow, 4).setBackground('#fff8e1');
      } else {
        activitySheet.getRange(subRow, 4).setBackground('#e8f5e9');
      }

      subRow++;
    }

    return {
      success: true,
      orderNumber: '#' + cleanOrder,
      action: isVoidAction ? 'voided' : 'reshipped',
      row: headerRow
    };

  } catch (e) {
    return {success: false, error: 'updateActivityRow error: ' + e.message};
  }
}

/**
 * Update an existing Flow Detail row (for void/cancel — no new row)
 * Finds original header row by order number, updates status + sub-item colors.
 *
 * @param {string} flow - 'F5' (or any flow key)
 * @param {string} orderNumber - e.g. "92906" or "#92906"
 * @param {string} newStatus - New status text for header (e.g. "Voided")
 * @param {string} newSubStatus - New status for sub-items (e.g. "Returned")
 * @param {string} newQtyColor - 'green' or 'red' for sub-item qty
 * @return {Object} {success, row} or {success: false, error}
 */
function updateFlowDetailRow(flow, orderNumber, newStatus, newSubStatus, newQtyColor) {
  try {
    var tabName = FLOW_TAB_NAMES[flow];
    if (!tabName) return {success: false, error: 'Unknown flow: ' + flow};

    var ss = SpreadsheetApp.openById(CONFIG.DEBUG_SHEET_ID);
    var sheet = ss.getSheetByName(tabName);
    if (!sheet) return {success: false, error: 'Tab not found: ' + tabName};

    var cleanOrder = String(orderNumber).replace('#', '');
    var searchPattern = '#' + cleanOrder;

    var lastRow = sheet.getLastRow();
    var headerRow = -1;

    // Search from bottom up for header row with this order number (col C = Ref#)
    for (var row = lastRow; row > 3; row--) {
      var colA = sheet.getRange(row, 1).getValue();
      if (!colA) continue; // sub-item rows have empty col A
      var ref = String(sheet.getRange(row, 3).getValue());
      if (ref.indexOf(searchPattern) >= 0) {
        headerRow = row;
        break;
      }
    }

    if (headerRow === -1) {
      return {success: false, error: 'F5 row not found for #' + cleanOrder};
    }

    // Update header status (col E)
    sheet.getRange(headerRow, 5).setValue(newStatus);
    // Color header status cell
    var sLower = newStatus.toLowerCase();
    if (sLower === 'voided' || sLower === 'cancelled') {
      sheet.getRange(headerRow, 5).setBackground('#fff3cd'); // amber
    } else if (sLower === 'shipped' || sLower === 'returned') {
      sheet.getRange(headerRow, 5).setBackground('#d4edda'); // green
    }

    // Update sub-item rows below header
    var subRow = headerRow + 1;
    while (subRow <= lastRow) {
      var subColA = sheet.getRange(subRow, 1).getValue();
      if (subColA) break; // hit next header

      var subDetail = String(sheet.getRange(subRow, 4).getValue());
      if (!subDetail) { subRow++; continue; }

      // Update sub-item status (col E)
      if (newSubStatus) {
        sheet.getRange(subRow, 5).setValue(newSubStatus);
      }

      // Flip qty color
      if (newQtyColor) {
        var qtyMatch = subDetail.match(/x(\d+)/);
        if (qtyMatch) {
          var qtyStr = qtyMatch[0];
          var qtyStart = subDetail.indexOf(qtyStr);
          var qtyEnd = qtyStart + qtyStr.length;
          var qtyHex = newQtyColor === 'green' ? '#008000' : newQtyColor === 'grey' ? '#999999' : '#cc0000';
          var qtyStyle = SpreadsheetApp.newTextStyle().setForegroundColor(qtyHex).build();
          var rtv = SpreadsheetApp.newRichTextValue()
            .setText(subDetail)
            .setTextStyle(qtyStart, qtyEnd, qtyStyle)
            .build();
          sheet.getRange(subRow, 4).setRichTextValue(rtv);
        }
      }

      // Flip sub-item background
      if (newQtyColor === 'green') {
        sheet.getRange(subRow, 4).setBackground('#fff8e1'); // amber for returned
      } else {
        sheet.getRange(subRow, 4).setBackground('#e8f5e9'); // green for active
      }

      subRow++;
    }

    return {success: true, row: headerRow};

  } catch (e) {
    return {success: false, error: 'updateFlowDetailRow error: ' + e.message};
  }
}

function finalizeF5ManualRecoveryRows(orderNumber, recoveredLines) {
  var orderKey = normalizeF5RecoveryOrder_(orderNumber);
  var orderLineMap = {};
  orderLineMap[orderKey] = recoveredLines || [];
  return finalizeF5ManualRecoveryBatch(orderLineMap);
}

function normalizeF5ManualRecoveryBatchOrders(orderKeys) {
  orderKeys = orderKeys || [];
  if (!orderKeys.length) {
    return { success: false, error: 'No orders supplied', normalizedOrders: 0 };
  }

  var activitySheet = getActivitySheet();
  var ss = SpreadsheetApp.openById(CONFIG.DEBUG_SHEET_ID);
  var flowSheet = ss.getSheetByName(FLOW_TAB_NAMES.F5);
  var activityIndex = buildF5ManualRecoverySheetIndex_(activitySheet, 'activity');
  var flowIndex = buildF5ManualRecoverySheetIndex_(flowSheet, 'flow');
  var normalizedOrders = 0;
  var compactedRows = 0;
  var results = {};

  for (var i = 0; i < orderKeys.length; i++) {
    var orderKey = normalizeF5RecoveryOrder_(orderKeys[i]);
    if (!orderKey) continue;

    var activityEntries = activityIndex[orderKey] || [];
    var flowEntries = flowIndex[orderKey] || [];
    var orderResult = {};

    if (activityEntries.length > 0) {
      orderResult.activity = normalizeF5ManualRecoveryIndexedEntries_(activitySheet, activityEntries);
    }
    if (flowEntries.length > 0) {
      orderResult.flow = normalizeF5ManualRecoveryIndexedEntries_(flowSheet, flowEntries);
    }

    if ((orderResult.activity && orderResult.activity.success) ||
        (orderResult.flow && orderResult.flow.success)) {
      normalizedOrders++;
    }
    results[orderKey] = orderResult;
  }

  compactedRows += compactBlankF5ManualRecoveryRows_(activitySheet);
  compactedRows += compactBlankF5ManualRecoveryRows_(flowSheet);

  return {
    success: normalizedOrders > 0 || compactedRows > 0,
    normalizedOrders: normalizedOrders,
    compactedRows: compactedRows,
    results: results
  };
}

function finalizeF5ManualRecoveryBatch(orderLineMap) {
  orderLineMap = orderLineMap || {};
  var orderKeys = Object.keys(orderLineMap);
  if (!orderKeys.length) {
    return { success: false, error: 'No recovered orders supplied', repairedOrders: 0, loggedOrders: 0 };
  }

  var activitySheet = getActivitySheet();
  var ss = SpreadsheetApp.openById(CONFIG.DEBUG_SHEET_ID);
  var flowSheet = ss.getSheetByName(FLOW_TAB_NAMES.F5);
  var activityIndex = buildF5ManualRecoverySheetIndex_(activitySheet, 'activity');
  var flowIndex = buildF5ManualRecoverySheetIndex_(flowSheet, 'flow');
  var repairedOrders = 0;
  var loggedOrders = 0;
  var results = {};

  for (var i = 0; i < orderKeys.length; i++) {
    var orderKey = normalizeF5RecoveryOrder_(orderKeys[i]);
    var recoveredLines = orderLineMap[orderKeys[i]] || [];
    if (!orderKey || !recoveredLines.length) continue;

    var activityEntries = activityIndex[orderKey] || [];
    var flowEntries = flowIndex[orderKey] || [];
    var orderResult = {
      orderKey: orderKey,
      repaired: false,
      logged: false
    };

    if (activityEntries.length > 0) {
      orderResult.activity = repairF5ManualRecoveryIndexedEntries_(activitySheet, activityEntries, recoveredLines);
      orderResult.activityNormalized = normalizeF5ManualRecoveryIndexedEntries_(activitySheet, activityEntries);
    }
    if (flowEntries.length > 0) {
      orderResult.flow = repairF5ManualRecoveryIndexedEntries_(flowSheet, flowEntries, recoveredLines);
      orderResult.flowNormalized = normalizeF5ManualRecoveryIndexedEntries_(flowSheet, flowEntries);
    }

    if ((orderResult.activity && orderResult.activity.success) ||
        (orderResult.flow && orderResult.flow.success) ||
        (orderResult.activityNormalized && orderResult.activityNormalized.success) ||
        (orderResult.flowNormalized && orderResult.flowNormalized.success)) {
      orderResult.repaired = true;
      repairedOrders++;
    }

    if (activityEntries.length === 0 && flowEntries.length === 0) {
      orderResult.logged = appendF5ManualRecoveryVisiblePair_(orderKey, recoveredLines).success;
      if (orderResult.logged) loggedOrders++;
    } else {
      var sharedExecId = '';

      if (activityEntries.length === 0 && flowEntries.length > 0 && flowEntries[0].execId) {
        var addActivity = appendF5ManualRecoveryActivityRow_(orderKey, recoveredLines, flowEntries[0].execId);
        orderResult.activityAppended = addActivity;
        if (addActivity.success) {
          orderResult.logged = true;
          sharedExecId = addActivity.execId || flowEntries[0].execId;
        }
      } else if (activityEntries.length > 0) {
        sharedExecId = activityEntries[0].execId || '';
      }

      if (flowEntries.length === 0) {
        var addFlow = appendF5ManualRecoveryFlowRow_(orderKey, recoveredLines, sharedExecId);
        orderResult.flowAppended = addFlow;
        if (addFlow.success) {
          orderResult.logged = true;
        }
      }

      if (activityEntries.length === 0 && !sharedExecId) {
        var addActivityFallback = appendF5ManualRecoveryActivityRow_(orderKey, recoveredLines, '');
        orderResult.activityAppendedFallback = addActivityFallback;
        if (addActivityFallback.success) {
          orderResult.logged = true;
          sharedExecId = addActivityFallback.execId || sharedExecId;
        }
      }

      if (orderResult.logged) loggedOrders++;
    }

    results[orderKey] = orderResult;
  }

  return {
    success: true,
    repairedOrders: repairedOrders,
    loggedOrders: loggedOrders,
    results: results
  };
}

function buildF5ManualRecoverySheetIndex_(sheet, sheetType) {
  var map = {};
  if (!sheet || sheet.getLastRow() < 4) return map;

  var data = sheet.getRange(4, 1, sheet.getLastRow() - 3, 6).getDisplayValues();
  var current = null;

  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var rowNum = i + 4;
    var execId = String(row[0] || '').trim();
    var flowText = String(row[2] || '').trim();
    var refOrDetail = sheetType === 'activity' ? String(row[3] || '').trim() : String(row[2] || '').trim();
    var detailText = String(row[3] || '').trim();
    var statusText = String(row[4] || '').trim();
    var errorText = String(row[5] || '').trim();

    if (execId) {
      var orderKey = '';
      if (sheetType === 'activity') {
        if (flowText !== 'F5 Shipping') {
          current = null;
          continue;
        }
        orderKey = extractF5ManualRecoveryOrderKey_(refOrDetail);
      } else {
        orderKey = extractF5ManualRecoveryOrderKey_(refOrDetail);
      }

      if (!orderKey) {
        current = null;
        continue;
      }

      current = {
        orderKey: orderKey,
        headerRow: rowNum,
        execId: execId,
        headerDetail: detailText,
        headerStatus: statusText,
        headerError: errorText,
        subRows: []
      };
      if (!map[orderKey]) map[orderKey] = [];
      map[orderKey].push(current);
      continue;
    }

    if (current && detailText) {
      current.subRows.push({
        rowNum: rowNum,
        detail: detailText,
        status: statusText,
        error: errorText
      });
    }
  }

  return map;
}

function repairF5ManualRecoveryIndexedEntries_(sheet, entries, recoveredLines) {
  entries = entries || [];
  var repaired = 0;
  var lastError = '';
  var details = [];

  for (var i = 0; i < entries.length; i++) {
    var result = repairF5ManualRecoveryIndexedEntry_(sheet, entries[i], recoveredLines);
    details.push(result);
    if (result && result.success) {
      repaired++;
    } else if (result && result.error) {
      lastError = result.error;
    }
  }

  return {
    success: repaired > 0,
    repairedEntries: repaired,
    details: details,
    error: repaired > 0 ? '' : (lastError || 'No matching failed sub-lines found')
  };
}

function normalizeF5ManualRecoveryIndexedEntries_(sheet, entries) {
  entries = entries || [];
  var normalized = 0;
  var changedRows = 0;
  var lastError = '';
  var details = [];
  var collapseResult = null;

  for (var i = 0; i < entries.length; i++) {
    var result = normalizeF5ManualRecoveryIndexedEntry_(sheet, entries[i]);
    details.push(result);
    if (result && result.success) {
      normalized++;
      changedRows += Number(result.changedRows || 0);
    } else if (result && result.error) {
      lastError = result.error;
    }
  }

  if (entries.length > 1) {
    collapseResult = collapseF5ManualRecoveryIndexedEntries_(sheet, entries);
    details.push(collapseResult);
    if (collapseResult && collapseResult.success) {
      changedRows += Number(collapseResult.changedRows || 0);
    } else if (collapseResult && collapseResult.error) {
      lastError = collapseResult.error;
    }
  }

  return {
    success: normalized > 0 || !!(collapseResult && collapseResult.success),
    normalizedEntries: normalized,
    changedRows: changedRows,
    details: details,
    error: (normalized > 0 || (collapseResult && collapseResult.success)) ? '' : (lastError || 'No F5 entries normalized')
  };
}

function extractF5ManualRecoveryOrderKey_(text) {
  var match = String(text || '').match(/#\s*([A-Za-z0-9-]+)/);
  return normalizeF5RecoveryOrder_(match ? match[1] : '');
}

function repairF5ManualRecoveryIndexedEntry_(sheet, entry, recoveredLines) {
  try {
    if (!sheet || !entry || !entry.subRows || !entry.subRows.length) {
      return { success: false, error: 'No indexed F5 entry found' };
    }

    var needMap = buildF5ManualRecoveryNeedMap_(recoveredLines);
    var matched = 0;
    var remainingFailed = 0;
    var totalSubRows = entry.subRows.length;

    for (var i = 0; i < entry.subRows.length; i++) {
      var sub = entry.subRows[i];
      var detailText = String(sub.detail || '');
      var statusText = String(sub.status || '');
      var parsed = parseF5ManualRecoverySubLine_(detailText);

      if (statusText === 'Failed' && parsed && consumeF5ManualRecoveryNeed_(needMap, parsed.sku, parsed.qty)) {
        matched++;
        var detailCell = sheet.getRange(sub.rowNum, 4);
        var statusCell = sheet.getRange(sub.rowNum, 5);
        var errorCell = sheet.getRange(sub.rowNum, 6);
        statusCell.setValue('Deducted').setBackground('#d4edda');
        errorCell.clearContent().setBackground('#ffffff');
        detailCell.setBackground('#e8f5e9');
        rewriteF5ManualRecoveryQtyColor_(detailCell, detailText, 'red');
        statusText = 'Deducted';
      }

      if (statusText === 'Failed') remainingFailed++;
    }

    var headerDetailCell = sheet.getRange(entry.headerRow, 4);
    var headerStatusCell = sheet.getRange(entry.headerRow, 5);
    var headerErrorCell = sheet.getRange(entry.headerRow, 6);
    if (matched === 0) {
      if (remainingFailed === 0 && (entry.headerStatus === 'Partial' || entry.headerStatus === 'Failed')) {
        headerStatusCell.setValue('Shipped').setBackground('#d4edda');
        headerErrorCell.clearContent().setBackground('#ffffff');
        var rewrittenHeaderNoFailed = rewriteF5ManualRecoveryHeaderDetail_(entry.headerDetail, 0);
        if (rewrittenHeaderNoFailed !== entry.headerDetail) {
          headerDetailCell.setValue(rewrittenHeaderNoFailed);
        }
        return {
          success: true,
          row: entry.headerRow,
          matched: 0,
          remainingFailed: 0,
          headerOnly: true
        };
      }
      return { success: false, error: 'No matching failed sub-lines found for #' + entry.orderKey };
    }

    var headerStatus = remainingFailed === 0 ? 'Shipped' : (remainingFailed === totalSubRows ? 'Failed' : 'Partial');

    headerStatusCell.setValue(headerStatus);
    if (headerStatus === 'Shipped') {
      headerStatusCell.setBackground('#d4edda');
      headerErrorCell.clearContent().setBackground('#ffffff');
    } else if (headerStatus === 'Partial') {
      headerStatusCell.setBackground('#fff3cd');
      headerErrorCell.setValue(remainingFailed + ' failed').setBackground('#fff0f0');
    } else {
      headerStatusCell.setBackground('#f8d7da');
      headerErrorCell.setValue(remainingFailed + ' failed').setBackground('#fff0f0');
    }

    var rewrittenHeader = rewriteF5ManualRecoveryHeaderDetail_(entry.headerDetail, remainingFailed);
    if (rewrittenHeader !== entry.headerDetail) {
      headerDetailCell.setValue(rewrittenHeader);
    }

    return {
      success: true,
      row: entry.headerRow,
      matched: matched,
      remainingFailed: remainingFailed
    };
  } catch (e) {
    return { success: false, error: 'repairF5ManualRecoveryIndexedEntry error: ' + e.message };
  }
}

function normalizeF5ManualRecoveryIndexedEntry_(sheet, entry) {
  try {
    if (!sheet || !entry || !entry.headerRow) {
      return { success: false, error: 'No indexed F5 entry found' };
    }

    var changedRows = 0;
    var subRows = entry.subRows || [];

    for (var i = 0; i < subRows.length; i++) {
      var sub = subRows[i];
      var detailText = String(sub.detail || '');
      var statusText = String(sub.status || '');
      var errorText = String(sub.error || '');
      var detailCell = sheet.getRange(sub.rowNum, 4);
      var statusCell = sheet.getRange(sub.rowNum, 5);
      var errorCell = sheet.getRange(sub.rowNum, 6);

      if (statusText === 'Failed') {
        statusCell.setValue('Deducted').setBackground('#d4edda');
        errorCell.clearContent().setBackground('#ffffff');
        detailCell.setBackground('#e8f5e9');
        rewriteF5ManualRecoveryQtyColor_(detailCell, detailText, 'red');
        changedRows++;
        continue;
      }

      if (statusText === 'Deducted' && errorText) {
        errorCell.clearContent().setBackground('#ffffff');
        detailCell.setBackground('#e8f5e9');
        rewriteF5ManualRecoveryQtyColor_(detailCell, detailText, 'red');
        changedRows++;
      }
    }

    var headerDetailCell = sheet.getRange(entry.headerRow, 4);
    var headerStatusCell = sheet.getRange(entry.headerRow, 5);
    var headerErrorCell = sheet.getRange(entry.headerRow, 6);
    var headerStatus = String(entry.headerStatus || '');
    var headerError = String(entry.headerError || '');
    var rewrittenHeader = rewriteF5ManualRecoveryHeaderDetail_(entry.headerDetail, 0);

    if (headerStatus !== 'Shipped' || headerError || rewrittenHeader !== entry.headerDetail) {
      headerStatusCell.setValue('Shipped').setBackground('#d4edda');
      headerErrorCell.clearContent().setBackground('#ffffff');
      if (rewrittenHeader !== entry.headerDetail) {
        headerDetailCell.setValue(rewrittenHeader);
      }
      changedRows++;
    }

    return {
      success: changedRows > 0,
      row: entry.headerRow,
      changedRows: changedRows
    };
  } catch (e) {
    return { success: false, error: 'normalizeF5ManualRecoveryIndexedEntry error: ' + e.message };
  }
}

function collapseF5ManualRecoveryIndexedEntries_(sheet, entries) {
  try {
    entries = entries || [];
    if (!sheet || entries.length <= 1) {
      return { success: false, error: 'No duplicate F5 entries to collapse', changedRows: 0 };
    }

    var canonical = selectCanonicalF5ManualRecoveryEntry_(entries);
    if (!canonical) {
      return { success: false, error: 'No canonical F5 entry selected', changedRows: 0 };
    }

    var changedRows = 0;
    var blankedEntries = 0;

    for (var i = 0; i < entries.length; i++) {
      var entry = entries[i];
      if (!entry || entry.headerRow === canonical.headerRow) continue;
      var blankResult = blankF5ManualRecoveryIndexedEntry_(sheet, entry);
      if (blankResult && blankResult.success) {
        blankedEntries++;
        changedRows += Number(blankResult.changedRows || 0);
      }
    }

    return {
      success: blankedEntries > 0,
      blankedEntries: blankedEntries,
      changedRows: changedRows,
      canonicalRow: canonical.headerRow
    };
  } catch (e) {
    return { success: false, error: 'collapseF5ManualRecoveryIndexedEntries error: ' + e.message, changedRows: 0 };
  }
}

function selectCanonicalF5ManualRecoveryEntry_(entries) {
  entries = entries || [];
  if (!entries.length) return null;

  var signatureCounts = {};
  var profiles = [];

  for (var i = 0; i < entries.length; i++) {
    var profile = buildF5ManualRecoveryEntryProfile_(entries[i]);
    profiles.push(profile);
    signatureCounts[profile.signature] = (signatureCounts[profile.signature] || 0) + 1;
  }

  var best = null;
  for (var j = 0; j < profiles.length; j++) {
    var current = profiles[j];
    current.signatureCount = signatureCounts[current.signature] || 0;
    if (!best || compareF5ManualRecoveryEntryProfiles_(current, best) > 0) {
      best = current;
    }
  }

  return best ? best.entry : null;
}

function buildF5ManualRecoveryEntryProfile_(entry) {
  var subRows = (entry && entry.subRows) || [];
  var distinctMap = {};
  var distinctKeys = [];
  var failedCount = 0;
  var errorCount = 0;

  for (var i = 0; i < subRows.length; i++) {
    var sub = subRows[i] || {};
    var parsed = parseF5ManualRecoverySubLine_(sub.detail || '');
    var key = parsed ? (parsed.sku + '|' + parsed.qty) : ('RAW|' + String(sub.detail || '').trim());
    if (!distinctMap[key]) {
      distinctMap[key] = true;
      distinctKeys.push(key);
    }
    if (String(sub.status || '') === 'Failed') failedCount++;
    if (String(sub.error || '').trim()) errorCount++;
  }

  distinctKeys.sort();

  var headerSku = extractF5ManualRecoveryHeaderSku_(entry.headerDetail);
  var containsHeaderSku = false;
  if (headerSku) {
    for (var j = 0; j < distinctKeys.length; j++) {
      if (distinctKeys[j].indexOf(headerSku + '|') === 0) {
        containsHeaderSku = true;
        break;
      }
    }
  }

  return {
    entry: entry,
    signature: distinctKeys.join('||') || ('__EMPTY__|' + String(entry.headerDetail || '')),
    distinctCount: distinctKeys.length,
    totalSubRows: subRows.length,
    duplicateCount: Math.max(0, subRows.length - distinctKeys.length),
    failedCount: failedCount,
    errorCount: errorCount + (String(entry.headerError || '').trim() ? 1 : 0),
    statusRank: getF5ManualRecoveryHeaderStatusRank_(entry.headerStatus),
    headerSku: headerSku,
    containsHeaderSku: containsHeaderSku,
    extraDistinctBeyondHeader: headerSku && containsHeaderSku ? Math.max(0, distinctKeys.length - 1) : 0
  };
}

function compareF5ManualRecoveryEntryProfiles_(a, b) {
  if ((a.signatureCount || 0) !== (b.signatureCount || 0)) {
    return (a.signatureCount || 0) - (b.signatureCount || 0);
  }
  if ((a.statusRank || 0) !== (b.statusRank || 0)) {
    return (a.statusRank || 0) - (b.statusRank || 0);
  }
  if (!!a.headerSku !== !!b.headerSku) {
    return a.headerSku ? 1 : -1;
  }
  if ((a.containsHeaderSku ? 1 : 0) !== (b.containsHeaderSku ? 1 : 0)) {
    return a.containsHeaderSku ? 1 : -1;
  }
  if ((a.extraDistinctBeyondHeader || 0) !== (b.extraDistinctBeyondHeader || 0)) {
    return (b.extraDistinctBeyondHeader || 0) - (a.extraDistinctBeyondHeader || 0);
  }
  if ((a.failedCount || 0) !== (b.failedCount || 0)) {
    return (b.failedCount || 0) - (a.failedCount || 0);
  }
  if ((a.errorCount || 0) !== (b.errorCount || 0)) {
    return (b.errorCount || 0) - (a.errorCount || 0);
  }
  if ((a.duplicateCount || 0) !== (b.duplicateCount || 0)) {
    return (b.duplicateCount || 0) - (a.duplicateCount || 0);
  }
  if ((a.distinctCount || 0) !== (b.distinctCount || 0)) {
    return (a.distinctCount || 0) - (b.distinctCount || 0);
  }
  return Number(a.entry.headerRow || 0) - Number(b.entry.headerRow || 0);
}

function getF5ManualRecoveryHeaderStatusRank_(status) {
  status = String(status || '').trim();
  if (status === 'Shipped' || status === 'Voided' || status === 'Returned') return 3;
  if (status === 'Partial') return 2;
  if (status === 'Failed') return 1;
  return 0;
}

function extractF5ManualRecoveryHeaderSku_(detailText) {
  var text = String(detailText || '').replace(/[â†’].*$/, '').trim();
  if (/items?/i.test(text)) return '';
  var match = text.match(/#\s*[A-Za-z0-9-]+\s+([A-Za-z0-9._\/-]+)\s+x\s*[0-9]+(?:\.[0-9]+)?/i);
  return match ? String(match[1] || '').trim() : '';
}

function blankF5ManualRecoveryIndexedEntry_(sheet, entry) {
  try {
    if (!sheet || !entry || !entry.headerRow) {
      return { success: false, error: 'No indexed F5 entry to blank', changedRows: 0 };
    }

    var rowCount = 1 + ((entry.subRows && entry.subRows.length) || 0);
    var range = sheet.getRange(entry.headerRow, 1, rowCount, 6);
    range.clearContent();
    range.setBackground('#ffffff');
    return {
      success: true,
      row: entry.headerRow,
      changedRows: rowCount
    };
  } catch (e) {
    return { success: false, error: 'blankF5ManualRecoveryIndexedEntry error: ' + e.message, changedRows: 0 };
  }
}

function compactBlankF5ManualRecoveryRows_(sheet) {
  if (!sheet) return 0;
  var lastRow = sheet.getLastRow();
  if (lastRow < 4) return 0;

  var values = sheet.getRange(4, 1, lastRow - 3, 6).getDisplayValues();
  var runs = [];
  var runStart = -1;

  for (var i = 0; i < values.length; i++) {
    var isBlank = true;
    for (var c = 0; c < 6; c++) {
      if (String(values[i][c] || '').trim()) {
        isBlank = false;
        break;
      }
    }

    if (isBlank) {
      if (runStart === -1) runStart = i + 4;
    } else if (runStart !== -1) {
      runs.push({ startRow: runStart, count: (i + 4) - runStart });
      runStart = -1;
    }
  }

  if (runStart !== -1) {
    runs.push({ startRow: runStart, count: (lastRow + 1) - runStart });
  }

  var deleted = 0;
  for (var r = runs.length - 1; r >= 0; r--) {
    if (runs[r].count <= 0) continue;
    sheet.deleteRows(runs[r].startRow, runs[r].count);
    deleted += runs[r].count;
  }

  return deleted;
}

function compactBlankF5ManualRecoverySheets_() {
  var deleted = 0;
  var activitySheet = getActivitySheet();
  var ss = SpreadsheetApp.openById(CONFIG.DEBUG_SHEET_ID);
  var flowSheet = ss.getSheetByName(FLOW_TAB_NAMES.F5);

  deleted += compactBlankF5ManualRecoveryRows_(activitySheet);
  deleted += compactBlankF5ManualRecoveryRows_(flowSheet);

  return {
    success: deleted > 0,
    compactedRows: deleted
  };
}

function appendF5ManualRecoveryVisiblePair_(orderNumber, recoveredLines) {
  var activity = appendF5ManualRecoveryActivityRow_(orderNumber, recoveredLines, '');
  if (!activity.success) return { success: false, error: activity.error || 'Failed to append Activity row' };

  var flow = appendF5ManualRecoveryFlowRow_(orderNumber, recoveredLines, activity.execId || '');
  return {
    success: flow.success,
    execId: activity.execId || '',
    activity: activity,
    flow: flow
  };
}

function appendF5ManualRecoveryActivityRow_(orderNumber, recoveredLines, preExecId) {
  try {
    var cleanOrder = normalizeF5RecoveryOrder_(orderNumber);
    if (!cleanOrder) return { success: false, error: 'Order number missing' };
    recoveredLines = recoveredLines || [];
    if (!recoveredLines.length) return { success: false, error: 'No recovered lines supplied' };

    var headerDetails = buildF5ManualRecoveryHeaderText_(cleanOrder, recoveredLines);
    var shopifyLink = 'https://admin.shopify.com/store/mymagichealer/orders?query=' + encodeURIComponent(cleanOrder);
    var subItems = buildF5ManualRecoverySubItems_(recoveredLines);
    var execId = logActivity(
      'F5',
      headerDetails,
      'success',
      '→ SHOPIFY @ MMH Kelowna',
      subItems,
      { text: '#' + cleanOrder, url: shopifyLink },
      preExecId || null,
      ''
    );

    return { success: true, execId: execId };
  } catch (e) {
    return { success: false, error: 'appendF5ManualRecoveryActivityRow error: ' + e.message };
  }
}

function appendF5ManualRecoveryFlowRow_(orderNumber, recoveredLines, execId) {
  try {
    var cleanOrder = normalizeF5RecoveryOrder_(orderNumber);
    if (!cleanOrder) return { success: false, error: 'Order number missing' };
    recoveredLines = recoveredLines || [];
    if (!recoveredLines.length) return { success: false, error: 'No recovered lines supplied' };

    execId = String(execId || '').trim();
    if (!execId) execId = 'WK-MR-' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'HHmmss');

    var headerDetails = buildF5ManualRecoveryHeaderText_(cleanOrder, recoveredLines);
    var shopifyLink = 'https://admin.shopify.com/store/mymagichealer/orders?query=' + encodeURIComponent(cleanOrder);
    var f5FlowItems = [];

    for (var i = 0; i < recoveredLines.length; i++) {
      f5FlowItems.push({
        sku: recoveredLines[i].sku,
        qty: recoveredLines[i].qty,
        detail: 'SHOPIFY  manual recovery',
        status: 'Deducted',
        error: '',
        qtyColor: 'red'
      });
    }

    logFlowDetail('F5', execId, {
      ref: '#' + cleanOrder,
      detail: headerDetails + '  → SHOPIFY @ MMH Kelowna',
      status: 'Shipped',
      linkText: '#' + cleanOrder,
      linkUrl: shopifyLink
    }, f5FlowItems);

    return { success: true, execId: execId };
  } catch (e) {
    return { success: false, error: 'appendF5ManualRecoveryFlowRow error: ' + e.message };
  }
}

function buildF5ManualRecoveryHeaderText_(orderNumber, recoveredLines) {
  recoveredLines = recoveredLines || [];
  if (recoveredLines.length === 1) {
    return '#' + orderNumber + '  ' + recoveredLines[0].sku + ' x' + recoveredLines[0].qty + '  manual recovery';
  }
  return '#' + orderNumber + '  ' + recoveredLines.length + ' items  manual recovery';
}

function buildF5ManualRecoverySubItems_(recoveredLines) {
  var items = [];
  recoveredLines = recoveredLines || [];
  for (var i = 0; i < recoveredLines.length; i++) {
    items.push({
      sku: recoveredLines[i].sku,
      qty: recoveredLines[i].qty,
      success: true,
      action: 'SHOPIFY  manual recovery',
      status: 'Deducted',
      error: '',
      qtyColor: 'red'
    });
  }
  return items;
}

function finalizeF5ManualRecoveryOnSheet_(sheet, sheetType, orderNumber, recoveredLines) {
  try {
    if (!sheet) return { success: false, error: 'Sheet not found' };

    var cleanOrder = String(orderNumber || '').replace('#', '').trim();
    if (!cleanOrder) return { success: false, error: 'Order number missing' };

    var searchPattern = '#' + cleanOrder;
    var lastRow = sheet.getLastRow();
    var headerRow = -1;

    for (var row = lastRow; row > 3; row--) {
      var colA = sheet.getRange(row, 1).getValue();
      if (!colA) continue;

      if (sheetType === 'activity') {
        var flowLabel = String(sheet.getRange(row, 3).getValue() || '');
        if (flowLabel !== 'F5 Shipping') continue;
        var activityDetails = String(sheet.getRange(row, 4).getDisplayValue() || sheet.getRange(row, 4).getValue() || '');
        if (activityDetails.indexOf(searchPattern) >= 0) {
          headerRow = row;
          break;
        }
      } else {
        var ref = String(sheet.getRange(row, 3).getDisplayValue() || sheet.getRange(row, 3).getValue() || '');
        if (ref.indexOf(searchPattern) >= 0) {
          headerRow = row;
          break;
        }
      }
    }

    if (headerRow === -1) {
      return { success: false, error: 'Row not found for #' + cleanOrder };
    }

    var needMap = buildF5ManualRecoveryNeedMap_(recoveredLines);
    var subRow = headerRow + 1;
    var matched = 0;
    var remainingFailed = 0;
    var totalSubRows = 0;

    while (subRow <= lastRow) {
      var subColA = sheet.getRange(subRow, 1).getValue();
      if (subColA) break;

      totalSubRows++;
      var detailCell = sheet.getRange(subRow, 4);
      var statusCell = sheet.getRange(subRow, 5);
      var errorCell = sheet.getRange(subRow, 6);
      var detailText = String(detailCell.getDisplayValue() || detailCell.getValue() || '');
      var statusText = String(statusCell.getValue() || '');
      var parsed = parseF5ManualRecoverySubLine_(detailText);

      if (statusText === 'Failed' && parsed && consumeF5ManualRecoveryNeed_(needMap, parsed.sku, parsed.qty)) {
        matched++;
        statusCell.setValue('Deducted').setBackground('#d4edda');
        errorCell.clearContent().setBackground('#ffffff');
        detailCell.setBackground('#e8f5e9');
        rewriteF5ManualRecoveryQtyColor_(detailCell, detailText, 'red');
        statusText = 'Deducted';
      }

      if (statusText === 'Failed') remainingFailed++;
      subRow++;
    }

    var headerDetailCell = sheet.getRange(headerRow, 4);
    var headerStatusCell = sheet.getRange(headerRow, 5);
    var headerErrorCell = sheet.getRange(headerRow, 6);
    var headerDetail = String(headerDetailCell.getDisplayValue() || headerDetailCell.getValue() || '');
    if (matched === 0) {
      if (remainingFailed === 0 && (String(headerStatusCell.getValue() || '') === 'Partial' || String(headerStatusCell.getValue() || '') === 'Failed')) {
        headerStatusCell.setValue('Shipped').setBackground('#d4edda');
        headerErrorCell.clearContent().setBackground('#ffffff');
        var rewrittenHeaderNoFailed = rewriteF5ManualRecoveryHeaderDetail_(headerDetail, 0);
        if (rewrittenHeaderNoFailed !== headerDetail) {
          headerDetailCell.setValue(rewrittenHeaderNoFailed);
        }
        return {
          success: true,
          row: headerRow,
          matched: 0,
          remainingFailed: 0,
          headerOnly: true
        };
      }
      return { success: false, error: 'No matching failed sub-lines found for #' + cleanOrder };
    }

    var headerStatus = remainingFailed === 0 ? 'Shipped' : (remainingFailed === totalSubRows ? 'Failed' : 'Partial');

    headerStatusCell.setValue(headerStatus);
    if (headerStatus === 'Shipped') {
      headerStatusCell.setBackground('#d4edda');
      headerErrorCell.clearContent().setBackground('#ffffff');
    } else if (headerStatus === 'Partial') {
      headerStatusCell.setBackground('#fff3cd');
      headerErrorCell.setValue(remainingFailed + ' failed').setBackground('#fff0f0');
    } else {
      headerStatusCell.setBackground('#f8d7da');
      headerErrorCell.setValue(remainingFailed + ' failed').setBackground('#fff0f0');
    }

    var rewrittenHeader = rewriteF5ManualRecoveryHeaderDetail_(headerDetail, remainingFailed);
    if (rewrittenHeader !== headerDetail) {
      headerDetailCell.setValue(rewrittenHeader);
    }

    return {
      success: true,
      row: headerRow,
      matched: matched,
      remainingFailed: remainingFailed
    };
  } catch (e) {
    return { success: false, error: 'finalizeF5ManualRecoveryOnSheet error: ' + e.message };
  }
}

function buildF5ManualRecoveryNeedMap_(recoveredLines) {
  var map = {};
  recoveredLines = recoveredLines || [];
  for (var i = 0; i < recoveredLines.length; i++) {
    var sku = String(recoveredLines[i].sku || '').trim();
    var qty = Number(recoveredLines[i].qty || 0) || 0;
    if (!sku || !(qty > 0)) continue;
    var key = sku + '|' + qty;
    map[key] = (map[key] || 0) + 1;
  }
  return map;
}

function consumeF5ManualRecoveryNeed_(needMap, sku, qty) {
  var key = String(sku || '').trim() + '|' + (Number(qty || 0) || 0);
  if (!needMap[key]) return false;
  needMap[key]--;
  return true;
}

function parseF5ManualRecoverySubLine_(detailText) {
  var text = String(detailText || '').replace(/[├└─│]/g, ' ');
  var match = text.match(/([A-Za-z0-9._\/-]+)\s*x\s*([0-9]+(?:\.[0-9]+)?)/i);
  if (!match) return null;
  return {
    sku: match[1],
    qty: Number(match[2]) || 0
  };
}

function rewriteF5ManualRecoveryQtyColor_(detailCell, detailText, colorName) {
  var qtyMatch = String(detailText || '').match(/x([0-9]+(?:\.[0-9]+)?)/);
  if (!qtyMatch) return;

  var qtyStr = qtyMatch[0];
  var qtyStart = detailText.indexOf(qtyStr);
  if (qtyStart < 0) return;

  var qtyEnd = qtyStart + qtyStr.length;
  var qtyHex = colorName === 'green' ? '#008000' : colorName === 'grey' ? '#999999' : '#cc0000';
  var qtyStyle = SpreadsheetApp.newTextStyle().setForegroundColor(qtyHex).build();
  var richText = SpreadsheetApp.newRichTextValue()
    .setText(detailText)
    .setTextStyle(qtyStart, qtyEnd, qtyStyle)
    .build();
  detailCell.setRichTextValue(richText);
}

function rewriteF5ManualRecoveryHeaderDetail_(detailText, remainingFailed) {
  var text = String(detailText || '');
  text = text.replace(/\s+\d+\s+error(?:s)?(?=\s+→|\s*$)/i, '');
  if (remainingFailed > 0) {
    if (text.indexOf('→') >= 0) {
      text = text.replace(/\s+→/, '  ' + remainingFailed + ' error' + (remainingFailed > 1 ? 's' : '') + '  →');
    } else {
      text += '  ' + remainingFailed + ' error' + (remainingFailed > 1 ? 's' : '');
    }
  }
  return text.replace(/\s{3,}/g, '  ').trim();
}

// ============================================
// ADJUSTMENT AUDIT LOG
// ============================================

/**
 * Log a stock adjustment to the "Adjustments Log" tab on the Command Center sheet.
 * Called by F2 (WASP callouts), manual sync sheet adjustments, and WASP Move events.
 *
 * 13-column format:
 *   Timestamp | Source | Action | User | SKU | Item Name | Site | Location |
 *   Lot/Batch | Expiry | Diff | Katana SA# | Status
 *
 * Katana SA# is a clickable hyperlink when saId is provided:
 *   https://factory.katanamrp.com/stockadjustment/{saId}
 *
 * Row background color by Action:
 *   Add → light green, Remove → light red, Move → light blue, Sync → light amber
 *
 * @param {string} source      - 'WASP' or 'Sync Sheet'
 * @param {string} action      - 'Add', 'Remove', 'Move', or 'Sync'
 * @param {string} user        - Who made the adjustment (or empty string)
 * @param {string} sku         - Item SKU
 * @param {string} itemName    - Human-readable item name (or empty string)
 * @param {string} site        - Site name (e.g. 'MMH Kelowna')
 * @param {string} location    - WASP location. For moves: 'FROM→TO'
 * @param {string} lot         - Lot/batch number (or empty string)
 * @param {string} expiry      - Expiry date YYYY-MM-DD (or empty string)
 * @param {number} diff        - Change amount (positive for add, negative for remove)
 * @param {string} katanaSaNum - Katana SA# display text (e.g. 'B-WAX' or 'B-WAX +2')
 * @param {number|null} saId   - Katana numeric SA ID for hyperlink (or null)
 * @param {string} status      - 'OK', 'ERROR', or 'skipped'
 */
function logAdjustment(source, action, user, sku, itemName, site, location, lot, expiry, diff, katanaSaNum, saId, status) {

  try {
    var ss = SpreadsheetApp.openById(CONFIG.DEBUG_SHEET_ID);
    var sheet = ss.getSheetByName('F2 Adjustments');
    if (!sheet) {
      sheet = ss.insertSheet('F2 Adjustments');
      var headers = ['Timestamp', 'Source', 'Action', 'User', 'SKU', 'Item Name', 'Site', 'Location', 'Lot / Batch', 'Expiry', 'Diff', 'Katana SA#', 'Status'];
      var colWidths  = [155, 100, 80, 120, 120, 180, 120, 130, 120, 90, 70, 130, 80];
      sheet.getRange(1, 1, 1, 13).setValues([headers]);
      sheet.getRange(1, 1, 1, 13).setBackground('#1F2937').setFontColor('#ffffff').setFontWeight('bold');
      sheet.setFrozenRows(1);
      for (var c = 0; c < colWidths.length; c++) {
        sheet.setColumnWidth(c + 1, colWidths[c]);
      }
      // Conditional formatting
      var fullRange   = sheet.getRange(2, 1, 998, 13); // whole rows A:M
      var diffRange   = sheet.getRange(2, 11, 998, 1); // col K  Diff
      var statusRange = sheet.getRange(2, 13, 998, 1); // col M  Status
      var rules = [];
      // Row background by Source (col B): WASP=blue, Sync/Google Sheet=yellow
      rules.push(SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$B2="WASP"')         .setBackground('#dbeafe').setRanges([fullRange]).build());
      rules.push(SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$B2="Sync Sheet"')   .setBackground('#fef9c3').setRanges([fullRange]).build());
      rules.push(SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$B2="Google Sheet"') .setBackground('#fef9c3').setRanges([fullRange]).build());
      // Status cell
      rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('OK')      .setBackground('#c8e6c9').setRanges([statusRange]).build());
      rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('ERROR')   .setBackground('#ffcdd2').setRanges([statusRange]).build());
      rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('skipped') .setBackground('#fff3cd').setRanges([statusRange]).build());
      rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Skipped') .setBackground('#fff3cd').setRanges([statusRange]).build());
      rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('SKIPPED') .setBackground('#fff3cd').setRanges([statusRange]).build());
      // Diff cell
      rules.push(SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThan(0).setBackground('#c8e6c9').setFontColor('#2e7d32').setRanges([diffRange]).build());
      rules.push(SpreadsheetApp.newConditionalFormatRule().whenNumberLessThan(0)   .setBackground('#ffcdd2').setFontColor('#c62828').setRanges([diffRange]).build());
      sheet.setConditionalFormatRules(rules);
    }

    var ts = Utilities.formatDate(new Date(), 'America/Vancouver', 'yyyy-MM-dd HH:mm:ss');
    // Write all columns except SA# (col 12) — handled separately for hyperlink
    var row = [ts, source, action, user || '', sku, itemName || '', site || '', location, lot || '', expiry || '', diff, '', status];
    var newRow = sheet.getLastRow() + 1;
    sheet.getRange(newRow, 1, 1, 13).setValues([row]);

    // Col 12: SA# as hyperlink if numeric ID is available, otherwise plain text
    if (katanaSaNum) {
      if (saId) {
        var saUrl = 'https://factory.katanamrp.com/stockadjustment/' + saId;
        sheet.getRange(newRow, 12).setFormula('=HYPERLINK("' + saUrl + '","' + katanaSaNum.replace(/"/g, '') + '")');
        sheet.getRange(newRow, 12).setFontColor('#1565C0').setFontWeight('bold');
      } else {
        sheet.getRange(newRow, 12).setValue(katanaSaNum);
      }
    }
  } catch (e) {
    Logger.log('logAdjustment error: ' + e.message);
  }
}

/**
 * Re-apply header formatting and conditional formatting rules to the F2 Adjustments tab.
 * Run once from Apps Script editor: setupF2AdjustmentsFormatting()
 * Safe to re-run — clears old rules and reapplies fresh.
 */
function setupF2AdjustmentsFormatting() {
  var ss = SpreadsheetApp.openById(CONFIG.DEBUG_SHEET_ID);
  var sheet = ss.getSheetByName('F2 Adjustments');
  if (!sheet) { Logger.log('F2 Adjustments tab not found'); return; }

  // Header row — dark background, white bold text
  sheet.getRange(1, 1, 1, 13)
    .setBackground('#1F2937')
    .setFontColor('#ffffff')
    .setFontWeight('bold');
  sheet.setFrozenRows(1);

  // Column widths: Timestamp|Source|Action|User|SKU|Item Name|Site|Location|Lot/Batch|Expiry|Diff|Katana SA#|Status
  var colWidths = [155, 100, 80, 120, 120, 180, 120, 130, 120, 90, 70, 130, 80];
  for (var c = 0; c < colWidths.length; c++) {
    sheet.setColumnWidth(c + 1, colWidths[c]);
  }

  // Conditional formatting — clear old rules and reapply
  sheet.clearConditionalFormatRules();

  var lastRow = Math.max(sheet.getLastRow(), 100);
  var fullRange   = sheet.getRange(2, 1, lastRow, 13);  // whole rows A:M
  var diffRange   = sheet.getRange(2, 11, lastRow, 1);  // col K  Diff
  var statusRange = sheet.getRange(2, 13, lastRow, 1);  // col M  Status

  var rules = [];

  // Row background by Action (col C) — checked before Source rules
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$C2="Add"')    .setBackground('#e8f5e9').setRanges([fullRange]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$C2="Remove"') .setBackground('#fce4d6').setRanges([fullRange]).build());

  // Status cell (col M): OK=green, ERROR=red, skipped=amber
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('OK')   .setBackground('#c8e6c9').setRanges([statusRange]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('ERROR').setBackground('#ffcdd2').setRanges([statusRange]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('skipped').setBackground('#fff3cd').setRanges([statusRange]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Skipped').setBackground('#fff3cd').setRanges([statusRange]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('SKIPPED').setBackground('#fff3cd').setRanges([statusRange]).build());

  // Diff cell (col K): positive=green, negative=red
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThan(0).setBackground('#c8e6c9').setFontColor('#2e7d32').setRanges([diffRange]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberLessThan(0)   .setBackground('#ffcdd2').setFontColor('#c62828').setRanges([diffRange]).build());

  sheet.setConditionalFormatRules(rules);
  Logger.log('setupF2AdjustmentsFormatting complete. ' + (lastRow - 1) + ' data rows covered.');
}

// ============================================
// WEBHOOK QUEUE LOG
// ============================================

/**
 * Log every incoming webhook + its result to the "Webhook Queue" tab.
 * Called by doPost() after every request. Used for live debugging.
 *
 * Tab columns (5): Timestamp | Action | Status | Result | Payload
 * Tab is auto-created if missing (same structure as setupWebhookQueueTab).
 * Keeps newest entries only — trims to 500 data rows automatically.
 *
 * @param {Object} payload - Parsed POST body (has .action or .event or .resource_type)
 * @param {Object} result  - Return value from routeWebhook() (has .status)
 */
function logWebhookQueue(payload, result) {
  try {
    var ss    = SpreadsheetApp.openById(CONFIG.DEBUG_SHEET_ID);
    var sheet = ss.getSheetByName('Webhook Queue');

    // Auto-create tab if missing (matches setupWebhookQueueTab structure)
    if (!sheet) {
      sheet = ss.insertSheet('Webhook Queue');
      var headers   = ['Timestamp', 'Action', 'Status', 'Result', 'Payload'];
      var colWidths = [160, 220, 80, 160, 460];
      sheet.getRange(1, 1, 1, 5).setValues([headers]);
      sheet.getRange(1, 1, 1, 5)
        .setBackground('#1F2937').setFontColor('#ffffff').setFontWeight('bold');
      sheet.setFrozenRows(1);
      for (var c = 0; c < colWidths.length; c++) {
        sheet.setColumnWidth(c + 1, colWidths[c]);
      }
      var statusRange = sheet.getRange(2, 3, 998, 1);
      var rules = [];
      rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('ok')        .setBackground('#c8e6c9').setRanges([statusRange]).build());
      rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('processed') .setBackground('#c8e6c9').setRanges([statusRange]).build());
      rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('reverted')  .setBackground('#c8e6c9').setRanges([statusRange]).build());
      rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('skipped')   .setBackground('#fff3cd').setRanges([statusRange]).build());
      rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Skipped')   .setBackground('#fff3cd').setRanges([statusRange]).build());
      rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('SKIPPED')   .setBackground('#fff3cd').setRanges([statusRange]).build());
      rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('ignored')   .setBackground('#fff3cd').setRanges([statusRange]).build());
      rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('error')     .setBackground('#ffcdd2').setRanges([statusRange]).build());
      rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('diag')      .setBackground('#e3f2fd').setRanges([statusRange]).build());
      sheet.setConditionalFormatRules(rules);
    }

    var ts     = Utilities.formatDate(new Date(), 'America/Vancouver', 'yyyy-MM-dd HH:mm:ss');
    var action = (payload && (payload.action || payload.event || payload.resource_type)) || '(none)';
    var status = (result && result.status) || 'unknown';

    // Result column: status + message/error/reason if present
    var resultStr = status;
    if (result) {
      if (result.message) resultStr += ': ' + result.message;
      else if (result.error) resultStr += ': ' + result.error;
      else if (result.reason) resultStr += ': ' + result.reason;
    }
    if (resultStr.length > 200) resultStr = resultStr.substring(0, 200) + '\u2026';

    // Payload column: full JSON, truncated to 800 chars
    var payloadStr = '';
    try { payloadStr = JSON.stringify(payload); } catch (je) { payloadStr = String(payload); }
    if (payloadStr.length > 800) payloadStr = payloadStr.substring(0, 800) + '\u2026';

    var newRow = sheet.getLastRow() + 1;
    sheet.getRange(newRow, 1, 1, 5).setValues([[ts, action, status, resultStr, payloadStr]]);

    // Trim to 500 data rows (keep header row 1), delete oldest from top
    var MAX_ROWS = 500;
    var dataRows = newRow - 1; // newRow is the row we just wrote; -1 for header
    if (dataRows > MAX_ROWS) {
      sheet.deleteRows(2, dataRows - MAX_ROWS);
    }

  } catch (e) {
    Logger.log('logWebhookQueue error: ' + e.message);
  }
}

/**
 * Write a "pending" receipt to the Webhook Queue BEFORE processing starts.
 * Guarantees every arriving event is recorded even if processing crashes or
 * times out (the row stays as "pending" forever — visible as a gap).
 *
 * @param  {Object} payload - Parsed POST body
 * @return {number} Sheet row number written, or -1 on error
 */
function logWebhookReceipt(payload) {
  try {
    var ss    = SpreadsheetApp.openById(CONFIG.DEBUG_SHEET_ID);
    var sheet = ss.getSheetByName('Webhook Queue');
    if (!sheet) return -1;  // Tab not yet created — skip silently

    var ts     = Utilities.formatDate(new Date(), 'America/Vancouver', 'yyyy-MM-dd HH:mm:ss');
    var action = (payload && (payload.action || payload.event || payload.resource_type)) || '(none)';

    var payloadStr = '';
    try { payloadStr = JSON.stringify(payload); } catch (je) { payloadStr = String(payload); }
    if (payloadStr.length > 800) payloadStr = payloadStr.substring(0, 800) + '\u2026';

    var newRow = sheet.getLastRow() + 1;
    sheet.getRange(newRow, 1, 1, 5).setValues([[ts, action, 'pending', '', payloadStr]]);
    return newRow;

  } catch (e) {
    Logger.log('logWebhookReceipt error: ' + e.message);
    return -1;
  }
}

/**
 * Update Status + Result columns on a previously written receipt row.
 * Called after routeWebhook() completes (or fails) to fill in the final outcome.
 * Also trims the tab to 500 data rows (trim moved here from logWebhookQueue).
 *
 * @param {number} rowNum - Row number returned by logWebhookReceipt()
 * @param {Object} result - Return value from routeWebhook() (has .status)
 */
function updateWebhookQueueRow(rowNum, result) {
  try {
    if (!rowNum || rowNum < 2) return;

    var ss    = SpreadsheetApp.openById(CONFIG.DEBUG_SHEET_ID);
    var sheet = ss.getSheetByName('Webhook Queue');
    if (!sheet) return;

    var status = (result && result.status) || 'unknown';
    var resultStr = status;
    if (result) {
      if (result.message) resultStr += ': ' + result.message;
      else if (result.error) resultStr += ': ' + result.error;
      else if (result.reason) resultStr += ': ' + result.reason;
    }
    if (resultStr.length > 200) resultStr = resultStr.substring(0, 200) + '\u2026';

    sheet.getRange(rowNum, 3, 1, 2).setValues([[status, resultStr]]);

    // Trim to 500 data rows — delete oldest from top
    var dataRows = sheet.getLastRow() - 1;
    if (dataRows > 500) {
      sheet.deleteRows(2, dataRows - 500);
    }

  } catch (e) {
    Logger.log('updateWebhookQueueRow error: ' + e.message);
  }
}
