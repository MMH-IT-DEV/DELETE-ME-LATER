// ============================================
// 07_PickMappings.gs - PICK ORDER TRACKING
// ============================================
// Stores mapping between WASP Pick Orders and Katana SOs
// Uses Google Sheets for persistence
// FIXED: Changed const/let to var for Google Apps Script compatibility
// ============================================

/**
 * Get or create the PickOrderMappings sheet
 */
function getPickMappingsSheet() {
  var ss = SpreadsheetApp.openById(CONFIG.DEBUG_SHEET_ID);
  var sheet = ss.getSheetByName(CONFIG.PICK_MAPPINGS_SHEET);

  // Create sheet if it doesn't exist
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.PICK_MAPPINGS_SHEET);
    sheet.getRange(1, 1, 1, 5).setValues([[
      'OrderNumber',
      'KatanaSoId',
      'CreatedAt',
      'Status',
      'CompletedAt'
    ]]);
    sheet.getRange(1, 1, 1, 5).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }

  return sheet;
}

/**
 * Store pick order to Katana SO mapping
 */
function storePickOrderMapping(pickOrderNumber, katanaSoId) {
  try {
    var sheet = getPickMappingsSheet();
    var cleanOrderNumber = String(pickOrderNumber).replace('#', '');
    var now = new Date().toISOString();

    // Check if already exists
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === cleanOrderNumber) {
        // Update existing row
        sheet.getRange(i + 1, 2).setValue(katanaSoId);
        sheet.getRange(i + 1, 4).setValue('pending');
        return;
      }
    }

    // Add new row
    sheet.appendRow([cleanOrderNumber, katanaSoId, now, 'pending', '']);

  } catch (error) {
    logToSheet('PICK_MAPPING_STORE_ERROR', {
      pickOrder: pickOrderNumber
    }, { error: error.message });
  }
}

/**
 * Get Katana SO ID from pick order number
 */
function getKatanaSoIdFromPickOrder(pickOrderNumber) {
  try {
    var sheet = getPickMappingsSheet();
    var cleanOrderNumber = String(pickOrderNumber).replace('#', '');

    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === cleanOrderNumber) {
        return data[i][1];
      }
    }

    return null;

  } catch (error) {
    logToSheet('PICK_MAPPING_LOOKUP_ERROR', {
      pickOrder: pickOrderNumber
    }, { error: error.message });
    return null;
  }
}

/**
 * Get pick order mapping by Katana SO ID (reverse lookup)
 * Returns {orderNumber, status, completedAt} or null
 */
function getPickOrderBySoId(soId) {
  try {
    var sheet = getPickMappingsSheet();
    var data = sheet.getDataRange().getValues();
    var soIdStr = String(soId);
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][1]) === soIdStr) {
        return {
          orderNumber: data[i][0],
          status: data[i][3],
          completedAt: data[i][4]
        };
      }
    }
    return null;
  } catch (error) {
    return null;
  }
}

/**
 * Mark pick order mapping as completed
 */
function clearPickOrderMapping(pickOrderNumber) {
  try {
    var sheet = getPickMappingsSheet();
    var cleanOrderNumber = String(pickOrderNumber).replace('#', '');
    var now = new Date().toISOString();

    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === cleanOrderNumber) {
        sheet.getRange(i + 1, 4).setValue('completed');
        sheet.getRange(i + 1, 5).setValue(now);
        return true;
      }
    }
    return false;

  } catch (error) {
    logToSheet('PICK_MAPPING_CLEAR_ERROR', {
      pickOrder: pickOrderNumber
    }, { error: error.message });
    return false;
  }
}

/**
 * Cleanup old completed mappings (run daily via trigger)
 */
function cleanupPickMappings() {
  try {
    var sheet = getPickMappingsSheet();
    var data = sheet.getDataRange().getValues();
    var cutoffDate = new Date();
    cutoffDate.setDate(cutoffDate.getDate() - 7); // 7 days ago

    var deletedCount = 0;

    // Go backwards to avoid row index issues
    for (var i = data.length - 1; i >= 1; i--) {
      var status = data[i][3];
      var completedAt = data[i][4];

      if (status === 'completed' && completedAt) {
        var completedDate = new Date(completedAt);
        if (completedDate < cutoffDate) {
          sheet.deleteRow(i + 1);
          deletedCount++;
        }
      }
    }

    return deletedCount;

  } catch (error) {
    logToSheet('PICK_MAPPING_CLEANUP_ERROR', {}, { error: error.message });
    return 0;
  }
}

/**
 * Migrate existing Script Properties to Sheet (run once)
 */
function migratePickMappingsToSheet() {
  var props = PropertiesService.getScriptProperties();
  var allProps = props.getProperties();
  var migratedCount = 0;

  for (var key in allProps) {
    if (key.indexOf('PICK_') === 0) {
      var orderNumber = key.replace('PICK_', '');
      var katanaSoId = allProps[key];

      storePickOrderMapping(orderNumber, katanaSoId);
      props.deleteProperty(key);
      migratedCount++;
      Logger.log('Migrated: ' + key + ' -> ' + katanaSoId);
    }
  }

  Logger.log('Migration complete. Migrated ' + migratedCount + ' mappings.');
  return migratedCount;
}
