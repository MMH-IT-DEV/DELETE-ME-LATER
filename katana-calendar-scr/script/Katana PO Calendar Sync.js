/**
 * Katana PO → Google Calendar Sync
 * Optimized version with lean formatting
 * 
 * Setup:
 * 1. Script Properties → Add KATANA_API_KEY and CALENDAR_ID
 * 2. Run syncPurchaseOrdersToCalendar() manually or set daily trigger
 */

// ============================================
// CONFIGURATION
// ============================================

function getConfig() {
  const scriptProperties = PropertiesService.getScriptProperties();
  
  return {
    // API Settings
    KATANA_API_URL: 'https://api.katanamrp.com/v1',
    KATANA_API_KEY: scriptProperties.getProperty('KATANA_API_KEY'),
    
    // Calendar Settings
    CALENDAR_ID: scriptProperties.getProperty('CALENDAR_ID') || 'primary',
    
    // Filters
    DAYS_AHEAD: 90,
    PO_STATUSES: ['NOT_RECEIVED', 'PARTIALLY_RECEIVED'],
    
    // Display Settings
    MAX_ITEMS_SHOWN: 5,
    EVENT_COLOR: CalendarApp.EventColor.BLUE,
    
    // Supplier Mapping (fallback when API returns 404)
    SUPPLIER_MAP: {
      '1495698': 'SUMMIT LABELS (waiting for designs to be finalized)',
      '1481785': 'BULK BY CHO',
      '1480608': 'Qingdao Hainuo Biological Engineering Co., Ltd. (WOUND CARE)',
      '1478053': 'a',
      '1468286': 'Walmart',
      '1459442': 'Pres-On Corporation (HOLD)',
      '1438093': 'pre',
      '1438092': 'Pres-On Corporation',
      '1432127': 'Rona',
      '1405977': 'ECOIN',
      '1404470': 'Onyx Containers',
      '1401380': 'summit',
      '1398857': 'uli',
      '1375587': 'Chef Supplies',
      '1375584': 'Home Depot',
      '1375583': 'Vevor',
      '1370417': 'Grainger',
      '1364405': 'Sneed',
      '1349804': 'HARFINGTON',
      '1344148': 'TEST',
      '1341091': 'AMAZON',
      '1340129': 'Global Industrial',
      '1319571': 'NORTH BROADVIEW FARMS',
      '1317149': 'JEDWARDS',
      '1292446': 'SHENZHEN ANGELAPACK CO. LTD',
      '1233825': 'HYNAUT (WOUND CARE)',
      '1195114': 'BEAUTY PRIVATE LABELS',
      '1190994': 'SUMMIT LABELS',
      '1129975': 'SOURCE GRAPHICS',
      '1001568': 'HIDDEN ON HALL FARM AND FEED',
      '885328': 'GUANGZHOU YISON PRINTING (CUSTOM BOXES)',
      '883923': 'RHODES FARMS',
      '876314': 'SHENZEN HAIK PRINTING (LEAFLETS)',
      '853028': 'MIEDEMA HONEY FARM',
      '819187': 'TRADE TECHNOCRATS',
      '819185': 'NEW DIRECTION AROMATICS',
      '819181': 'COSTCO',
      '819170': 'ULINE'
    }
  };
}

// ============================================
// MAIN SYNC FUNCTION
// ============================================

function syncPurchaseOrdersToCalendar() {
  const CONFIG = getConfig();
  const startTime = new Date();

  log('🚀 Starting sync...', 'INFO');

  try {
    validateConfig(CONFIG);
    const calendar = getCalendar(CONFIG);
    const purchaseOrders = fetchOpenPurchaseOrders(CONFIG);

    log('📦 Found ' + purchaseOrders.length + ' open POs', 'INFO');

    // Collect open PO IDs for cleanup later
    var openPOIds = [];
    for (var j = 0; j < purchaseOrders.length; j++) {
      openPOIds.push(String(purchaseOrders[j].id));
    }

    var created = 0, updated = 0, skipped = 0;

    for (var i = 0; i < purchaseOrders.length; i++) {
      const result = processPurchaseOrder(purchaseOrders[i], calendar, CONFIG);
      if (result === 'CREATED') created++;
      else if (result === 'UPDATED') updated++;
      else skipped++;
    }

    // Remove events for POs that are no longer open (received)
    const removed = removeClosedPOEvents(calendar, openPOIds, CONFIG);

    const duration = ((new Date() - startTime) / 1000).toFixed(2);
    const summary = created + ' created, ' + updated + ' updated, ' + skipped + ' skipped, ' + removed + ' removed';
    log('✅ Done in ' + duration + 's: ' + summary, 'SUCCESS');

    // Heartbeat — report success to Health Monitor
    try { sendHeartbeatToMonitor_('Katana → Google Calendar PO Sync', 'success', summary); } catch (hbErr) { log('Heartbeat error: ' + hbErr.message, 'WARN'); }

  } catch (error) {
    log('❌ Failed: ' + error.message, 'ERROR');
    try { sendHeartbeatToMonitor_('Katana → Google Calendar PO Sync', 'error', error.message); } catch (hbErr) { log('Heartbeat error: ' + hbErr.message, 'WARN'); }
    throw error;
  }
}

// Remove calendar events for POs that are no longer open (fully received)
function removeClosedPOEvents(calendar, openPOIds, CONFIG) {
  var removed = 0;
  
  // Search calendar events in next 120 days
  const now = new Date();
  const futureDate = new Date();
  futureDate.setDate(futureDate.getDate() + 120);
  
  const events = calendar.getEvents(now, futureDate);
  
  for (var i = 0; i < events.length; i++) {
    const event = events[i];
    const desc = event.getDescription() || '';
    
    // Check if this is a PO event
    const poIdMatch = desc.match(/PO_ID:(\d+)/);
    if (poIdMatch) {
      const eventPOId = poIdMatch[1];
      
      // If PO ID is not in open list, delete the event
      if (openPOIds.indexOf(eventPOId) === -1) {
        const title = event.getTitle();
        event.deleteEvent();
        log('🗑️ Removed ' + title + ' (PO received)', 'INFO');
        removed++;
      }
    }
  }
  
  return removed;
}

// ============================================
// API FUNCTIONS
// ============================================

function fetchOpenPurchaseOrders(CONFIG) {
  const url = CONFIG.KATANA_API_URL + '/purchase_orders';
  const response = UrlFetchApp.fetch(url, {
    method: 'GET',
    headers: {
      'Accept': 'application/json',
      'Authorization': 'Bearer ' + CONFIG.KATANA_API_KEY
    },
    muteHttpExceptions: true
  });
  
  if (response.getResponseCode() !== 200) {
    throw new Error('Katana API error: ' + response.getResponseCode());
  }
  
  const data = JSON.parse(response.getContentText());
  const cutoffDate = new Date();
  cutoffDate.setDate(cutoffDate.getDate() + CONFIG.DAYS_AHEAD);
  
  return data.data.filter(function(po) {
    const arrivalDate = new Date(po.expected_arrival_date);
    return CONFIG.PO_STATUSES.indexOf(po.status) !== -1 && arrivalDate <= cutoffDate;
  });
}

function fetchSupplier(supplierId, CONFIG) {
  if (!supplierId) return { name: 'Unknown Supplier' };
  
  // Try API first
  const url = CONFIG.KATANA_API_URL + '/suppliers/' + supplierId;
  try {
    const response = UrlFetchApp.fetch(url, {
      method: 'GET',
      headers: {
        'Accept': 'application/json',
        'Authorization': 'Bearer ' + CONFIG.KATANA_API_KEY
      },
      muteHttpExceptions: true
    });
    
    if (response.getResponseCode() === 200) {
      return JSON.parse(response.getContentText());
    }
  } catch (e) {}
  
  // Fallback to mapping table
  const supplierIdStr = String(supplierId);
  if (CONFIG.SUPPLIER_MAP[supplierIdStr]) {
    return { name: CONFIG.SUPPLIER_MAP[supplierIdStr] };
  }
  
  return { name: 'Supplier #' + supplierId };
}

function fetchVariantDetails(variantId, CONFIG) {
  if (!variantId) return null;
  
  try {
    const url = CONFIG.KATANA_API_URL + '/variants/' + variantId;
    const response = UrlFetchApp.fetch(url, {
      method: 'GET',
      headers: {
        'Accept': 'application/json',
        'Authorization': 'Bearer ' + CONFIG.KATANA_API_KEY
      },
      muteHttpExceptions: true
    });
    
    if (response.getResponseCode() === 200) {
      return JSON.parse(response.getContentText());
    }
  } catch (e) {}
  
  return null;
}

function fetchProductName(productId, CONFIG) {
  if (!productId) return '';
  
  try {
    const url = CONFIG.KATANA_API_URL + '/products/' + productId;
    const response = UrlFetchApp.fetch(url, {
      method: 'GET',
      headers: {
        'Accept': 'application/json',
        'Authorization': 'Bearer ' + CONFIG.KATANA_API_KEY
      },
      muteHttpExceptions: true
    });
    
    if (response.getResponseCode() === 200) {
      const product = JSON.parse(response.getContentText());
      return product.name || '';
    }
  } catch (e) {}
  
  return '';
}

function fetchMaterialName(materialId, CONFIG) {
  if (!materialId) return '';
  
  try {
    const url = CONFIG.KATANA_API_URL + '/materials/' + materialId;
    const response = UrlFetchApp.fetch(url, {
      method: 'GET',
      headers: {
        'Accept': 'application/json',
        'Authorization': 'Bearer ' + CONFIG.KATANA_API_KEY
      },
      muteHttpExceptions: true
    });
    
    if (response.getResponseCode() === 200) {
      const material = JSON.parse(response.getContentText());
      return material.name || '';
    }
  } catch (e) {}
  
  return '';
}

// ============================================
// ITEM NAME RESOLUTION
// ============================================

function getItemName(variantId, CONFIG) {
  const variant = fetchVariantDetails(variantId, CONFIG);
  if (!variant) return 'Item #' + variantId;
  
  // Get base product/material name
  var baseName = '';
  if (variant.product_id) {
    baseName = fetchProductName(variant.product_id, CONFIG);
  } else if (variant.material_id) {
    baseName = fetchMaterialName(variant.material_id, CONFIG);
  }
  
  // Get variant type from config_attributes (e.g., "LAVENDER")
  var variantType = '';
  if (variant.config_attributes && variant.config_attributes.length > 0) {
    // Collect all config values
    var configValues = [];
    for (var i = 0; i < variant.config_attributes.length; i++) {
      var configValue = variant.config_attributes[i].config_value;
      if (configValue) {
        configValues.push(configValue);
      }
    }
    variantType = configValues.join(' / ');
  }
  
  // Build full name: "ESSENTIAL OIL / LAVENDER"
  if (baseName && variantType) {
    return baseName + ' / ' + variantType;
  } else if (baseName) {
    return baseName;
  } else if (variant.sku) {
    return variant.sku;
  }
  
  return 'Item #' + variantId;
}

// ============================================
// EVENT FORMATTING (LEAN VERSION)
// ============================================

function formatEventTitle(po, CONFIG) {
  // Use order_no (e.g., "PO-492"), fallback to ID only if missing
  const poNumber = po.order_no ? po.order_no : 'PO-' + po.id;
  const supplier = fetchSupplier(po.supplier_id, CONFIG);
  const supplierName = supplier.name || 'Unknown Supplier';
  
  return '🚚 ' + poNumber + ' • ' + supplierName;
}

function formatEventDescription(po, CONFIG) {
  // Use order_no (e.g., "PO-492"), fallback to ID only if missing
  const poNumber = po.order_no ? po.order_no : 'PO-' + po.id;
  const supplier = fetchSupplier(po.supplier_id, CONFIG);
  const supplierName = supplier.name || 'Unknown Supplier';
  
  // Format values
  const itemCount = po.purchase_order_rows ? po.purchase_order_rows.length : 0;
  const status = formatStatus(po.status);
  
  // Build lean description
  var desc = '';
  desc += '📦 ' + poNumber + ' | ' + supplierName + '\n';
  desc += '📍 ' + status + ' | ' + itemCount + ' items\n';
  desc += '\n';
  
  // Items section with quantities
  if (po.purchase_order_rows && po.purchase_order_rows.length > 0) {
    desc += '📋 Items:\n';
    
    const itemsToShow = Math.min(CONFIG.MAX_ITEMS_SHOWN, po.purchase_order_rows.length);
    for (var i = 0; i < itemsToShow; i++) {
      const row = po.purchase_order_rows[i];
      const itemName = getItemName(row.variant_id, CONFIG);
      const qty = formatQuantityWithUom(row.quantity, row.purchase_uom);
      desc += '• ' + itemName + ' • ' + qty + '\n';
    }
    
    if (itemCount > CONFIG.MAX_ITEMS_SHOWN) {
      desc += '  +' + (itemCount - CONFIG.MAX_ITEMS_SHOWN) + ' more\n';
    }
  }
  
  // Hidden metadata for event matching
  desc += '\n<!-- PO_ID:' + po.id + ' -->';
  
  return desc;
}

// ============================================
// CALENDAR OPERATIONS
// ============================================

function processPurchaseOrder(po, calendar, CONFIG) {
  const poNumber = po.purchase_order_no || ('PO-' + po.id);
  
  const existingEvent = findExistingEvent(calendar, po);
  
  if (existingEvent) {
    if (shouldUpdateEvent(existingEvent, po)) {
      updateCalendarEvent(existingEvent, po, CONFIG);
      log('✏️ Updated ' + poNumber, 'INFO');
      return 'UPDATED';
    } else {
      return 'SKIPPED';
    }
  } else {
    createCalendarEvent(calendar, po, CONFIG);
    log('✅ Created ' + poNumber, 'INFO');
    return 'CREATED';
  }
}

function findExistingEvent(calendar, po) {
  const arrivalDate = new Date(po.expected_arrival_date);
  
  const startOfDay = new Date(arrivalDate);
  startOfDay.setHours(0, 0, 0, 0);
  
  const endOfDay = new Date(arrivalDate);
  endOfDay.setHours(23, 59, 59, 999);
  
  const events = calendar.getEvents(startOfDay, endOfDay);
  
  for (var i = 0; i < events.length; i++) {
    const desc = events[i].getDescription() || '';
    if (desc.indexOf('PO_ID:' + po.id) !== -1) {
      return events[i];
    }
  }
  
  return null;
}

function shouldUpdateEvent(event, po) {
  // Check if date changed
  const currentDate = event.getAllDayStartDate();
  const newDate = new Date(po.expected_arrival_date);
  
  if (currentDate.toISOString().split('T')[0] !== newDate.toISOString().split('T')[0]) {
    return true;
  }
  
  // Check if old format (force update)
  const desc = event.getDescription() || '';
  if (desc.indexOf('═══') !== -1) {
    return true; // Old verbose format, update to lean
  }
  
  return false;
}

function createCalendarEvent(calendar, po, CONFIG) {
  const arrivalDate = new Date(po.expected_arrival_date);
  const title = formatEventTitle(po, CONFIG);
  const description = formatEventDescription(po, CONFIG);
  
  calendar.createAllDayEvent(title, arrivalDate, {
    description: description,
    color: CONFIG.EVENT_COLOR
  });
}

function updateCalendarEvent(event, po, CONFIG) {
  const arrivalDate = new Date(po.expected_arrival_date);
  const title = formatEventTitle(po, CONFIG);
  const description = formatEventDescription(po, CONFIG);
  
  event.setTitle(title);
  event.setDescription(description);
  event.setAllDayDate(arrivalDate);
}

// ============================================
// HELPER FUNCTIONS
// ============================================

function validateConfig(CONFIG) {
  if (!CONFIG.KATANA_API_KEY) {
    throw new Error('KATANA_API_KEY not set in Script Properties');
  }
}

function getCalendar(CONFIG) {
  const calendar = CalendarApp.getCalendarById(CONFIG.CALENDAR_ID);
  if (!calendar) {
    throw new Error('Calendar not found: ' + CONFIG.CALENDAR_ID);
  }
  return calendar;
}

function formatNumber(num) {
  if (!num) return '0';
  return num.toFixed(2).replace(/\d(?=(\d{3})+\.)/g, '$&,');
}

function formatQuantity(qty) {
  if (!qty) return '0 pcs';
  const num = Math.round(parseFloat(qty));
  return num.toLocaleString() + ' pcs';
}

function formatQuantityWithUom(qty, uom) {
  if (!qty) return '0';
  const num = Math.round(parseFloat(qty));
  const formattedNum = num.toLocaleString();
  
  // Use provided UoM or default to pcs
  const unit = uom ? uom : 'pcs';
  return formattedNum + ' ' + unit;
}

function formatStatus(status) {
  const statusMap = {
    'NOT_RECEIVED': 'Not Received',
    'PARTIALLY_RECEIVED': 'Partially Received',
    'RECEIVED': 'Received'
  };
  return statusMap[status] || status;
}

function log(message, level) {
  console.log('[' + level + '] ' + message);
}

// ============================================
// TEST & UTILITY FUNCTIONS
// ============================================

function testFetchPurchaseOrders() {
  const CONFIG = getConfig();
  validateConfig(CONFIG);
  const pos = fetchOpenPurchaseOrders(CONFIG);
  log('Found ' + pos.length + ' open POs within ' + CONFIG.DAYS_AHEAD + ' days', 'INFO');
  
  pos.forEach(function(po) {
    log('  ' + (po.purchase_order_no || po.id) + ' - ' + po.status + ' - ' + po.expected_arrival_date, 'INFO');
  });
}

function testCalendarAccess() {
  const CONFIG = getConfig();
  const calendar = getCalendar(CONFIG);
  log('✅ Calendar access OK: ' + calendar.getName(), 'SUCCESS');
}

function listSupplierIDs() {
  const CONFIG = getConfig();
  validateConfig(CONFIG);
  const pos = fetchOpenPurchaseOrders(CONFIG);
  
  const suppliers = {};
  pos.forEach(function(po) {
    if (po.supplier_id && !suppliers[po.supplier_id]) {
      suppliers[po.supplier_id] = [];
    }
    if (po.supplier_id) {
      suppliers[po.supplier_id].push(po.purchase_order_no || po.id);
    }
  });
  
  log('Supplier IDs found:', 'INFO');
  for (var id in suppliers) {
    log("  '" + id + "': '[NAME]',  // Used in: " + suppliers[id].join(', '), 'INFO');
  }
}

function deleteAllPOEvents() {
  // Use with caution - deletes all PO events from calendar
  const CONFIG = getConfig();
  const calendar = getCalendar(CONFIG);
  
  const now = new Date();
  const futureDate = new Date();
  futureDate.setDate(futureDate.getDate() + 120);
  
  const events = calendar.getEvents(now, futureDate);
  var deleted = 0;
  
  events.forEach(function(event) {
    const desc = event.getDescription() || '';
    if (desc.indexOf('PO_ID:') !== -1) {
      event.deleteEvent();
      deleted++;
    }
  });
  
  log('Deleted ' + deleted + ' PO events', 'INFO');
}

// Debug: Check what API returns for a single PO
function debugSinglePO() {
  const CONFIG = getConfig();
  validateConfig(CONFIG);
  
  const pos = fetchOpenPurchaseOrders(CONFIG);
  if (pos.length === 0) {
    log('No open POs found', 'INFO');
    return;
  }
  
  const po = pos[0];
  
  // Log ALL fields from PO object
  log('=== ALL PO FIELDS ===', 'INFO');
  for (var key in po) {
    if (key !== 'purchase_order_rows') {
      log(key + ': ' + JSON.stringify(po[key]), 'INFO');
    }
  }
  
  if (po.purchase_order_rows && po.purchase_order_rows.length > 0) {
    log('=== FIRST ROW FIELDS ===', 'INFO');
    const row = po.purchase_order_rows[0];
    for (var rowKey in row) {
      log('row.' + rowKey + ': ' + JSON.stringify(row[rowKey]), 'INFO');
    }
  }
}

// Debug: Fetch single PO by ID to see full details
function debugFetchSinglePO() {
  const CONFIG = getConfig();
  validateConfig(CONFIG);
  
  // Get first PO ID from list
  const pos = fetchOpenPurchaseOrders(CONFIG);
  if (pos.length === 0) {
    log('No open POs found', 'INFO');
    return;
  }
  
  const poId = pos[0].id;
  log('Fetching full details for PO ID: ' + poId, 'INFO');
  
  // Fetch individual PO
  const url = CONFIG.KATANA_API_URL + '/purchase_orders/' + poId;
  const response = UrlFetchApp.fetch(url, {
    method: 'GET',
    headers: {
      'Accept': 'application/json',
      'Authorization': 'Bearer ' + CONFIG.KATANA_API_KEY
    },
    muteHttpExceptions: true
  });
  
  if (response.getResponseCode() === 200) {
    const po = JSON.parse(response.getContentText());
    log('=== SINGLE PO ENDPOINT FIELDS ===', 'INFO');
    for (var key in po) {
      if (key !== 'purchase_order_rows') {
        log(key + ': ' + JSON.stringify(po[key]), 'INFO');
      }
    }
  } else {
    log('Error: ' + response.getResponseCode(), 'ERROR');
  }
}

// Debug: Check material full details
function debugMaterial() {
  const CONFIG = getConfig();
  validateConfig(CONFIG);
  
  const materialId = 7428686; // From previous debug
  const url = CONFIG.KATANA_API_URL + '/materials/' + materialId;
  
  const response = UrlFetchApp.fetch(url, {
    method: 'GET',
    headers: {
      'Accept': 'application/json',
      'Authorization': 'Bearer ' + CONFIG.KATANA_API_KEY
    },
    muteHttpExceptions: true
  });
  
  if (response.getResponseCode() === 200) {
    const material = JSON.parse(response.getContentText());
    log('=== ALL MATERIAL FIELDS ===', 'INFO');
    for (var key in material) {
      log(key + ': ' + JSON.stringify(material[key]), 'INFO');
    }
  } else {
    log('Error: ' + response.getResponseCode(), 'ERROR');
  }
}

// Fetch all suppliers from API and list their IDs
function testListAllSuppliers() {
  const CONFIG = getConfig();
  const url = CONFIG.KATANA_API_URL + '/suppliers';
  
  const response = UrlFetchApp.fetch(url, {
    method: 'GET',
    headers: {
      'Accept': 'application/json',
      'Authorization': 'Bearer ' + CONFIG.KATANA_API_KEY
    },
    muteHttpExceptions: true
  });
  
  log('Status: ' + response.getResponseCode(), 'INFO');
  
  if (response.getResponseCode() === 200) {
    const data = JSON.parse(response.getContentText());
    const suppliers = data.data || data;
    
    log('=== ALL SUPPLIERS ===', 'INFO');
    for (var i = 0; i < suppliers.length; i++) {
      log("'" + suppliers[i].id + "': '" + suppliers[i].name + "',", 'INFO');
    }
    log('Total: ' + suppliers.length + ' suppliers', 'INFO');
  } else {
    log('Error fetching suppliers - may need manual lookup', 'ERROR');
  }
}

// ============================================
// HEARTBEAT — Health Monitor integration
// Set HEALTH_MONITOR_URL in Script Properties
// ============================================

/**
 * @deprecated — internal config no longer used. See sendHeartbeatToMonitor_() below.
 */
function getHeartbeatConfig() {
  return {
    // System Health Monitor spreadsheet
    HEALTH_MONITOR_ID: '1jnWtdBPzR7DreihCHQASiN7spImRS7HYjTR77Gpfi5w',
    SYSTEM_ROW: 6,  // Row for "Katana → Google Calendar PO Sync" in System Registry
    
    // Column positions (1-indexed)
    COL_STATUS: 12,        // Column L - Status
    COL_HEARTBEAT: 13,     // Column M - Last Heartbeat
    COL_ERROR: 14,         // Column N - Last Error
    
    // Slack webhook for alerts (set in Script Properties)
    SLACK_WEBHOOK: PropertiesService.getScriptProperties().getProperty('SLACK_WEBHOOK_URL'),
    
    // Slack channel for IT alerts
    SLACK_CHANNEL: '#it-alerts',
    
    // Status indicators
    STATUS_HEALTHY: '🟢 Healthy',
    STATUS_WARNING: '🟡 Warning',
    STATUS_ERROR: '🔴 Error',
    
    // Stale threshold (hours) - alert if no heartbeat in this time
    STALE_THRESHOLD_HOURS: 6
  };
}

/**
 * @deprecated — replaced by sendHeartbeatToMonitor_() below.
 */
function sendHeartbeat(success, errorMessage) {
  const config = getHeartbeatConfig();
  const now = new Date().toISOString();
  
  try {
    const sheet = SpreadsheetApp.openById(config.HEALTH_MONITOR_ID)
                                .getSheetByName('System Registry');
    
    if (success) {
      // Update status to Healthy
      sheet.getRange(config.SYSTEM_ROW, config.COL_STATUS).setValue(config.STATUS_HEALTHY);
      sheet.getRange(config.SYSTEM_ROW, config.COL_HEARTBEAT).setValue(now);
      sheet.getRange(config.SYSTEM_ROW, config.COL_ERROR).setValue('');
      
      Logger.log('✅ Heartbeat sent: Healthy at ' + now);
    } else {
      // Update status to Error
      sheet.getRange(config.SYSTEM_ROW, config.COL_STATUS).setValue(config.STATUS_ERROR);
      sheet.getRange(config.SYSTEM_ROW, config.COL_HEARTBEAT).setValue(now);
      sheet.getRange(config.SYSTEM_ROW, config.COL_ERROR).setValue(errorMessage);
      
      // Send Slack alert
      sendSlackAlert('🔴 Katana PO Calendar Sync Failed', errorMessage);
      
      Logger.log('❌ Heartbeat sent: Error - ' + errorMessage);
    }
  } catch (e) {
    Logger.log('⚠️ Failed to send heartbeat: ' + e.message);
    // Try to send Slack alert even if spreadsheet update fails
    sendSlackAlert('⚠️ Heartbeat Update Failed', 'Could not update System Health Monitor: ' + e.message);
  }
}

/**
 * Send Slack alert for errors
 */
function sendSlackAlert(title, message) {
  const config = getHeartbeatConfig();
  
  if (!config.SLACK_WEBHOOK) {
    Logger.log('⚠️ Slack webhook not configured - skipping alert');
    return;
  }
  
  const payload = {
    blocks: [
      {
        type: "header",
        text: {
          type: "plain_text",
          text: title,
          emoji: true
        }
      },
      {
        type: "section",
        text: {
          type: "mrkdwn",
          text: "*System:* Katana → Google Calendar PO Sync\n*Time:* " + new Date().toLocaleString() + "\n*Error:* " + message
        }
      },
      {
        type: "section",
        text: {
          type: "mrkdwn",
          text: "📋 <https://script.google.com/home/projects/1mjraAHwRu5Nfel1yRzjuTK8I4uZPrl_QyUsYnZ67Bkmp-GTLFdUEMC1o/edit|Open Script> | <https://docs.google.com/spreadsheets/d/1jnWtdBPzR7DreihCHQASiN7spImRS7HYjTR77Gpfi5w/edit|System Health Monitor>"
        }
      }
    ]
  };
  
  try {
    UrlFetchApp.fetch(config.SLACK_WEBHOOK, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload)
    });
    Logger.log('📨 Slack alert sent');
  } catch (e) {
    Logger.log('⚠️ Failed to send Slack alert: ' + e.message);
  }
}

/**
 * Check if heartbeat is stale (no update in X hours)
 * Run this on a separate trigger (e.g., every 4 hours)
 */
function checkHeartbeatStale() {
  const config = getHeartbeatConfig();
  
  try {
    const sheet = SpreadsheetApp.openById(config.HEALTH_MONITOR_ID)
                                .getSheetByName('System Registry');
    
    const lastHeartbeat = sheet.getRange(config.SYSTEM_ROW, config.COL_HEARTBEAT).getValue();
    const currentStatus = sheet.getRange(config.SYSTEM_ROW, config.COL_STATUS).getValue();
    
    if (!lastHeartbeat) {
      Logger.log('⚠️ No heartbeat recorded yet');
      return;
    }
    
    const lastTime = new Date(lastHeartbeat);
    const now = new Date();
    const hoursSinceHeartbeat = (now - lastTime) / (1000 * 60 * 60);
    
    Logger.log('Last heartbeat: ' + hoursSinceHeartbeat.toFixed(1) + ' hours ago');
    
    if (hoursSinceHeartbeat > config.STALE_THRESHOLD_HOURS && currentStatus !== config.STATUS_ERROR) {
      // Mark as warning and alert
      sheet.getRange(config.SYSTEM_ROW, config.COL_STATUS).setValue(config.STATUS_WARNING);
      sheet.getRange(config.SYSTEM_ROW, config.COL_ERROR).setValue('Stale heartbeat - no sync in ' + Math.round(hoursSinceHeartbeat) + ' hours');
      
      sendSlackAlert(
        '🟡 Katana PO Sync - Stale Heartbeat',
        'No successful sync in ' + Math.round(hoursSinceHeartbeat) + ' hours. The sync may have stopped running.'
      );
    }
  } catch (e) {
    Logger.log('⚠️ Failed to check heartbeat: ' + e.message);
  }
}

/**
 * Test heartbeat functions
 */
function testHeartbeatSuccess() {
  sendHeartbeat(true);
}

function testHeartbeatError() {
  sendHeartbeat(false, 'Test error message - ignore this alert');
}

function testSlackAlert() {
  sendSlackAlert('🧪 Test Alert', 'This is a test alert from Katana PO Calendar Sync');
}


// ═══════════════════════════════════════════════════════════════════════════════
// UPDATED syncPurchaseOrdersToCalendar() with Heartbeat
// Replace your existing function with this version
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * Main sync function with heartbeat integration
 * This wraps the sync in try/catch and reports status
 */
function syncPurchaseOrdersToCalendarWithHeartbeat() {
  let success = true;
  let errorMessage = '';
  
  try {
    // Call the original sync function
    syncPurchaseOrdersToCalendar();
    
  } catch (e) {
    success = false;
    errorMessage = e.message;
    Logger.log('❌ Sync failed: ' + e.message);
  }
  
  // Send heartbeat regardless of success/failure
  sendHeartbeat(success, errorMessage);
}

// ═══════════════════════════════════════════════════════════════════════════════
// ALTERNATIVE: Modify existing syncPurchaseOrdersToCalendar()
// Add this code at the END of your existing syncPurchaseOrdersToCalendar function
// ═══════════════════════════════════════════════════════════════════════════════

/*
Add this to the END of syncPurchaseOrdersToCalendar(), just before the closing brace:

  // === HEARTBEAT ===
  try {
    sendHeartbeat(true);
  } catch (heartbeatError) {
    Logger.log('⚠️ Heartbeat failed: ' + heartbeatError.message);
  }

And wrap the entire function body in try/catch:

function syncPurchaseOrdersToCalendar() {
  try {
    // ... existing code ...
    
    // At the very end, send success heartbeat
    sendHeartbeat(true);
    
  } catch (e) {
    // Send error heartbeat
    sendHeartbeat(false, e.message);
    throw e;  // Re-throw to preserve original error behavior
  }
}
*/


// ═══════════════════════════════════════════════════════════════════════════════
// SETUP INSTRUCTIONS
// ═══════════════════════════════════════════════════════════════════════════════

/*
SETUP STEPS:

1. Add Slack Webhook to Script Properties:
   - Go to Project Settings → Script Properties
   - Add: SLACK_WEBHOOK_URL = https://hooks.slack.com/services/YOUR/WEBHOOK/URL
   
   To get a webhook URL:
   - Go to https://api.slack.com/apps
   - Create new app (or use existing)
   - Enable "Incoming Webhooks"
   - Add webhook to #it-alerts channel
   - Copy the webhook URL

2. Update your existing trigger:
   - Change trigger to run: syncPurchaseOrdersToCalendarWithHeartbeat
   - (instead of syncPurchaseOrdersToCalendar)

3. Add stale heartbeat check trigger:
   - Create new trigger for: checkHeartbeatStale
   - Run every 4 hours

SETUP (current approach):
1. Go to Project Settings → Script Properties
2. Add: HEALTH_MONITOR_URL = https://script.google.com/macros/s/AKfycbziw4_NeS0BcJDL4qGJrRqHKsTVxATPSRJzg3YpItetajxcgd3BLuUcxOZKykLSdr1UnA/exec
3. The daily trigger should call syncPurchaseOrdersToCalendar() — heartbeat fires automatically.
*/

// ============================================
// STANDARD HEARTBEAT HELPER
// POSTs to Health Monitor doPost() endpoint.
// Requires Script Property: HEALTH_MONITOR_URL
// ============================================

function sendHeartbeatToMonitor_(systemName, status, details) {
  var url = PropertiesService.getScriptProperties().getProperty('HEALTH_MONITOR_URL');
  if (!url) { Logger.log('sendHeartbeatToMonitor_: HEALTH_MONITOR_URL not set'); return; }

  var payload = JSON.stringify({
    system:  systemName,
    status:  status,
    details: details || ''
  });

  try {
    UrlFetchApp.fetch(url, {
      method:             'post',
      contentType:        'application/json',
      payload:            payload,
      muteHttpExceptions: true
    });
  } catch (e) {
    Logger.log('sendHeartbeatToMonitor_ fetch error: ' + e.message);
  }
}