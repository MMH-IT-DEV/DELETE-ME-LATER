/**
 * Sync Engine - Core sync logic, auto-create, logging
 * writeKatanaRaw / writeWaspRaw / writeBatchTab live in 04b_RawTabs.gs
 */

var SYNC_LOCATION_MAP = {
  'MMH Kelowna|Product': { site: 'MMH Kelowna', location: 'PRODUCTION' },
  'MMH Kelowna|Material': { site: 'MMH Kelowna', location: 'PRODUCTION' },
  'MMH Kelowna|Intermediate': { site: 'MMH Kelowna', location: 'PRODUCTION' },
  'MMH Kelowna|': { site: 'MMH Kelowna', location: 'PRODUCTION' },
  'MMH Mayfair|Product': { site: 'MMH Mayfair', location: 'PRODUCTION' },
  'MMH Mayfair|Material': { site: 'MMH Mayfair', location: 'PRODUCTION' },
  'MMH Mayfair|Intermediate': { site: 'MMH Mayfair', location: 'PRODUCTION' },
  'MMH Mayfair|': { site: 'MMH Mayfair', location: 'PRODUCTION' },
  'Storage Warehouse|Product': { site: 'Storage Warehouse', location: 'SW-STORAGE' },
  'Storage Warehouse|Material': { site: 'Storage Warehouse', location: 'SW-STORAGE' },
  'Storage Warehouse|Intermediate': { site: 'Storage Warehouse', location: 'SW-STORAGE' },
  'Storage Warehouse|': { site: 'Storage Warehouse', location: 'SW-STORAGE' }
};

var SKIP_SKU_PREFIXES = ['OP-'];

// Location separator bar colors (Katana side)
var LOCATION_COLORS = {
  'MMH Kelowna': '#1565c0',
  'MMH Mayfair': '#2e7d32',
  'Shopify': '#00897b',
  'Storage Warehouse': '#6a1b9a'
};

function isSkippedSku(sku) {
  if (!sku) return false;
  for (var i = 0; i < SKIP_SKU_PREFIXES.length; i++) {
    if (sku.indexOf(SKIP_SKU_PREFIXES[i]) === 0) return true;
  }
  return false;
}

function formatTimestamp() {
  return Utilities.formatDate(new Date(), 'America/Vancouver', 'yyyy-MM-dd HH:mm');
}

function clearDataRows(sheet) {
  if (!sheet) return;
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow > 1 && lastCol > 0) {
    var range = sheet.getRange(2, 1, lastRow - 1, lastCol);
    range.clearContent();
    range.clearFormat();
    range.clearDataValidations();
  }
}

// ============================================
// FIND NEW ITEMS - Detects Katana SKUs not in WASP
// ============================================

function findNewItems(katanaData, waspData) {
  var waspSkus = {};
  for (var i = 0; i < waspData.items.length; i++) {
    var item = waspData.items[i];
    if (item.itemNumber) waspSkus[item.itemNumber] = true;
  }

  var newItems = [];
  var seen = {};
  for (var j = 0; j < katanaData.items.length; j++) {
    var kItem = katanaData.items[j];
    if (!kItem.sku || isSkippedSku(kItem.sku)) continue;
    if (waspSkus[kItem.sku]) continue;
    if (seen[kItem.sku]) continue;
    newItems.push({ sku: kItem.sku, name: kItem.name, type: kItem.type });
    seen[kItem.sku] = true;
  }

  return newItems;
}

// ============================================
// AUTO-CREATE MISSING ITEMS
// ============================================

/**
 * Auto-creates Katana items that don't exist in WASP.
 * Called during every sync.
 *
 * @param {Object} katanaData - Full Katana data with items[]
 * @param {Object} waspData - Full WASP data with items[]
 * @param {Object} batchData - Batch data with lotTrackedSkus{}
 * @returns {Object} { created: [], failed: [] }
 */
function autoCreateMissingItems(katanaData, waspData, batchData) {
  var startTime = new Date();
  Logger.log('autoCreateMissingItems: starting...');

  // Build set of existing WASP SKUs
  var waspSkuSet = {};
  for (var w = 0; w < waspData.items.length; w++) {
    var wItem = waspData.items[w];
    if (wItem.itemNumber) waspSkuSet[wItem.itemNumber] = true;
  }

  // Build Katana SKU info: sku -> { type, name, batchTracked }
  var katanaSkuInfo = {};
  for (var k = 0; k < katanaData.items.length; k++) {
    var kItem = katanaData.items[k];
    if (!kItem.sku || isSkippedSku(kItem.sku)) continue;
    if (!katanaSkuInfo[kItem.sku]) {
      katanaSkuInfo[kItem.sku] = {
        type: kItem.type || '',
        name: kItem.name || kItem.sku,
        batchTracked: batchData.lotTrackedSkus[kItem.sku] ? true : false,
                uom: kItem.uom || 'ea'
      };
    }
  }

  var created = [];
  var failed = [];

  // Find and create missing SKUs
  var katanaSkus = Object.keys(katanaSkuInfo);
  for (var i = 0; i < katanaSkus.length; i++) {
    var sku = katanaSkus[i];
    if (waspSkuSet[sku]) continue; // Already exists in WASP

    var info = katanaSkuInfo[sku];
    var site = WASP_DEFAULT_SITE;
    var location = 'QA-Hold-1';

    Logger.log('Creating item: ' + sku + ' (' + info.type + ') at ' + site + '/' + location);

    var result = waspCreateItem(sku, info.name, site, location, info.batchTracked, info.uom);

    if (result.success) {
      created.push(sku);
      Logger.log('✓ Created: ' + sku);
    } else {
      failed.push({ sku: sku, error: result.response || 'Unknown error' });
      Logger.log('✗ Failed: ' + sku + ' - ' + (result.response || 'Unknown error'));
    }
  }

  var duration = ((new Date()) - startTime) / 1000;
  Logger.log('autoCreateMissingItems: created=' + created.length + ' failed=' + failed.length + ' duration=' + duration.toFixed(1) + 's');

  return { created: created, failed: failed };
}

/**
 * Helper that creates an item in WASP.
 *
 * @param {string} sku - Item number
 * @param {string} name - Item description
 * @param {string} site - Site name
 * @param {string} location - Location code
 * @param {boolean} lotTracking - Enable lot tracking
 * @param {string} [uom] - Unit of measure (e.g. 'ea', 'kg', 'lb')
 * @param {number} [salesPrice] - Sales price
 * @param {number} [cost] - Unit cost (purchase/avg cost)
 * @returns {Object} { success, code, response }
 */
function normalizeUomForWasp(uom) {
  if (!uom) return 'Each';
  var u = String(uom).trim().toLowerCase();
  if (u === 'each' || u === 'ea' || u === 'pc' || u === 'pcs' || u === '1' || u === 'g' || u === 'bucket' || u === 'unit') return 'Each';
  if (u === 'pk' || u === 'pack' || u === '6pack') return 'pack';
  if (u === 'box') return 'BOX';
  if (u === 'kg') return 'KG';
  if (u === 'rolls' || u === 'roll') return 'ROLLS';
  if (u === 'dozen' || u === 'dz') return 'Each';
  if (u === 'pound' || u === 'lb') return 'Each';
  return 'Each';
}
function waspCreateItem(sku, name, site, location, lotTracking, uom, salesPrice, cost) {
  var url = getWaspBase() + 'ic/item/createInventoryItems';
  var item = {
    ItemNumber: sku,
    ItemDescription: name || sku,
    StockingUnit: normalizeUomForWasp(uom)
  };
  if (salesPrice !== undefined && salesPrice !== null && salesPrice > 0) {
    item.SalesPrice = salesPrice;
  }
  if (cost !== undefined && cost !== null && cost > 0) {
    item.Cost = cost;
  }
  if (lotTracking) {
    item.TrackbyInfo = {
      TrackedByLot: true,
      TrackedByDateCode: true
    };
  }
  var targetSite = site || WASP_DEFAULT_SITE;
  var targetLoc = location || '';
  if (targetLoc) {
    item.ItemLocationSettings = [{
      SiteName: targetSite,
      LocationCode: targetLoc,
      PrimaryLocation: true
    }];
    item.ItemSiteSettings = [{
      SiteName: targetSite
    }];
  }
  var payload = [item];
  return waspApiCall(url, payload);
}

// ============================================
// AUTO-SYNC WASP ITEM METADATA (Cost, Description, UOM, Category)
// Runs before every Sync Now to keep WASP items in sync with Katana.
// ============================================

/**
 * Syncs item metadata (description, cost, UOM, category) from Katana to WASP.
 * Called automatically during every runFullSyncCore_().
 *
 * - Description: Katana item name → WASP ItemDescription
 * - Cost: Katana average_cost (highest non-zero per SKU) → WASP Cost
 *         If Katana cost is 0 (stock depleted), keeps current WASP cost
 * - UOM: Katana UOM normalized to WASP format
 * - Category: Katana category → WASP CategoryDescription
 *
 * WASP's updateInventoryItems resets omitted fields, so ALL fields are
 * sent in every payload to preserve values.
 *
 * @param {Object} katanaData - from fetchAllKatanaData()
 * @returns {Object} { updated, skipped, failed, duration }
 */
function syncWaspItemMetadata(katanaData) {
  var startTime = new Date();
  Logger.log('syncWaspItemMetadata: starting...');

  // Build Katana SKU map: sku -> { name, cost, uom, category }
  var katanaMap = {};
  for (var ki = 0; ki < katanaData.items.length; ki++) {
    var kItem = katanaData.items[ki];
    if (!kItem.sku) continue;
    var kCost = parseFloat(kItem.avgCost) || 0;

    if (!katanaMap[kItem.sku]) {
      katanaMap[kItem.sku] = {
        name: kItem.name || '',
        cost: kCost,
        uom: kItem.uom || 'ea',
        category: kItem.category || ''
      };
    } else {
      if (kCost > 0 && kCost > katanaMap[kItem.sku].cost) {
        katanaMap[kItem.sku].cost = kCost;
      }
      if (!katanaMap[kItem.sku].name && kItem.name) {
        katanaMap[kItem.sku].name = kItem.name;
      }
      if (!katanaMap[kItem.sku].category && kItem.category) {
        katanaMap[kItem.sku].category = kItem.category;
      }
    }
  }

  // Fetch WASP item catalog via advancedinfosearch (has Cost, Description, Category)
  var token = getStoredWaspToken_();
  var base = getWaspBase();
  var allWaspItems = [];
  var page = 1;
  var hasMore = true;

  while (hasMore && page <= 20) {
    var searchResp = UrlFetchApp.fetch(base + 'ic/item/advancedinfosearch', {
      method: 'POST',
      headers: { 'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json' },
      payload: JSON.stringify({ SearchPattern: '', PageSize: 100, PageNumber: page }),
      muteHttpExceptions: true
    });
    if (searchResp.getResponseCode() !== 200) break;
    var body = JSON.parse(searchResp.getContentText());
    var data = body.Data || [];
    for (var d = 0; d < data.length; d++) { allWaspItems.push(data[d]); }
    hasMore = body.HasSuccessWithMoreDataRemaining === true;
    page++;
    Utilities.sleep(300);
  }

  Logger.log('syncWaspItemMetadata: fetched ' + allWaspItems.length + ' WASP catalog items');

  // Compare and update
  var updated = 0;
  var skipped = 0;
  var failed = 0;

  for (var wi = 0; wi < allWaspItems.length; wi++) {
    var wItem = allWaspItems[wi];
    var sku = wItem.ItemNumber || '';
    if (!sku) continue;

    var kData = katanaMap[sku];
    if (!kData) { skipped++; continue; }

    var wDesc = wItem.ItemDescription || '';
    var wCost = parseFloat(wItem.Cost) || 0;
    var wUom = wItem.StockingUnit || 'Each';
    var wCat = wItem.CategoryDescription || '';

    var newDesc = kData.name || wDesc;
    var newUom = normalizeUomForWasp(kData.uom);
    var newCost = kData.cost > 0 ? kData.cost : wCost;
    var newCat = kData.category || wCat;

    // Skip if nothing changed
    if (newDesc === wDesc && Math.abs(newCost - wCost) < 0.01 && newUom === wUom && newCat === wCat) {
      skipped++;
      continue;
    }

    var updatePayload = [{
      ItemNumber: sku,
      ItemDescription: newDesc,
      StockingUnit: newUom,
      PurchaseUnit: newUom,
      SalesUnit: newUom,
      Cost: newCost,
      CategoryDescription: newCat,
      DimensionInfo: wItem.DimensionInfo || {
        DimensionUnit: 'Inch', Height: 0, Width: 0, Depth: 0,
        WeightUnit: 'Pound', Weight: 0, VolumeUnit: 'Cubic Inch', MaxVolume: 0
      }
    }];

    try {
      var resp = UrlFetchApp.fetch(base + 'ic/item/updateInventoryItems', {
        method: 'POST',
        headers: { 'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json' },
        payload: JSON.stringify(updatePayload),
        muteHttpExceptions: true
      });
      var respBody = resp.getContentText();
      if (resp.getResponseCode() === 200) {
        var parsed = JSON.parse(respBody);
        if (parsed.HasError === false) {
          updated++;
        } else if (respBody.indexOf('-57072') > -1 || respBody.indexOf('-57006') > -1) {
          // UOM or category not in WASP — retry with current values
          if (respBody.indexOf('-57072') > -1) {
            updatePayload[0].StockingUnit = wUom;
            updatePayload[0].PurchaseUnit = wUom;
            updatePayload[0].SalesUnit = wUom;
          }
          if (respBody.indexOf('-57006') > -1) {
            updatePayload[0].CategoryDescription = wCat;
          }
          var retry = UrlFetchApp.fetch(base + 'ic/item/updateInventoryItems', {
            method: 'POST',
            headers: { 'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json' },
            payload: JSON.stringify(updatePayload),
            muteHttpExceptions: true
          });
          var retryParsed = JSON.parse(retry.getContentText());
          if (retryParsed.HasError === false) { updated++; } else { failed++; }
        } else {
          failed++;
        }
      } else {
        failed++;
      }
    } catch (e) {
      failed++;
    }
    Utilities.sleep(300);
  }

  var duration = ((new Date()) - startTime) / 1000;
  Logger.log('syncWaspItemMetadata: updated=' + updated + ' skipped=' + skipped + ' failed=' + failed + ' duration=' + duration.toFixed(1) + 's');
  return { updated: updated, skipped: skipped, failed: failed, duration: duration };
}

// ============================================
// ENHANCED SYNC HISTORY LOGGING
// ============================================

/**
 * Sync history logger — 7 columns, color-coded by status and type.
 *
 * @param {Spreadsheet} ss
 * @param {string} type        - 'SYNC', 'PUSH', 'AUTO-CREATE'
 * @param {string} summary     - Brief human-readable summary (no error details)
 * @param {number} errCount    - Number of errors (0 = clean run)
 * @param {string} errDetails  - Semicolon-separated error strings ('' if none)
 * @param {number} duration    - Duration in seconds
 * @param {string} status      - 'SUCCESS', 'PARTIAL', 'ERROR'
 */
function logSyncHistory(ss, type, summary, errCount, errDetails, duration, status) {
  var sheet = ss.getSheetByName('Sync History');
  if (!sheet) {
    sheet = ss.insertSheet('Sync History');
    setupSyncHistoryHeaders_(sheet);
  } else if (sheet.getRange(1, 3).getValue() !== 'Status') {
    // Header is stale (old 5-col layout) — update in place
    setupSyncHistoryHeaders_(sheet);
  }

  var row = [
    formatTimestamp(),                           // A Timestamp
    type,                                        // B Type
    status,                                      // C Status
    parseFloat((duration || 0).toFixed(1)),      // D Duration (s)
    summary,                                     // E Summary
    errCount > 0 ? errCount : '',                // F Errors (blank if 0)
    errDetails || ''                             // G Error Details
  ];

  try {
    var nextRow = sheet.getLastRow() + 1;
    var rowRange = sheet.getRange(nextRow, 1, 1, 7);
    rowRange.setValues([row]);

    // Row background — subtle tint based on outcome
    var rowBg = status === 'SUCCESS' ? '#f1f8e9' :
                status === 'PARTIAL' ? '#fff8e1' : '#fce4ec';
    rowRange.setBackground(rowBg);

    // Type cell — blue for SYNC, purple for PUSH
    var typeCell = sheet.getRange(nextRow, 2);
    if (type === 'SYNC') {
      typeCell.setBackground('#e3f2fd').setFontColor('#0d47a1').setFontWeight('bold');
    } else if (type === 'PUSH') {
      typeCell.setBackground('#f3e5f5').setFontColor('#4a148c').setFontWeight('bold');
    } else {
      typeCell.setFontWeight('bold');
    }

    // Status cell — stronger color badge
    var statusCell = sheet.getRange(nextRow, 3);
    if (status === 'SUCCESS') {
      statusCell.setBackground('#c8e6c9').setFontColor('#1b5e20').setFontWeight('bold');
    } else if (status === 'PARTIAL') {
      statusCell.setBackground('#ffe0b2').setFontColor('#e65100').setFontWeight('bold');
    } else {
      statusCell.setBackground('#ef9a9a').setFontColor('#b71c1c').setFontWeight('bold');
    }

    // Error count cell — bold red when non-zero
    if (errCount > 0) {
      sheet.getRange(nextRow, 6).setFontColor('#c62828').setFontWeight('bold');
    }

  } catch (e) {
    Logger.log('Sync History log error: ' + e.message);
    try {
      sheet.clear();
      SpreadsheetApp.flush();
      setupSyncHistoryHeaders_(sheet);
      sheet.getRange(2, 1, 1, 7).setValues([row]);
    } catch (e2) {
      Logger.log('Sync History reset also failed: ' + e2.message);
    }
  }
}

function setupSyncHistoryHeaders_(sheet) {
  var headers = ['Timestamp', 'Type', 'Status', 'Sec', 'Summary', '#Err', 'Error Details'];
  var hRange = sheet.getRange(1, 1, 1, 7);
  hRange.setValues([headers]);
  hRange.setBackground('#263238').setFontColor('#ffffff').setFontWeight('bold').setFontSize(10);
  sheet.setFrozenRows(1);
  sheet.setColumnWidth(1, 130); // Timestamp
  sheet.setColumnWidth(2, 55);  // Type
  sheet.setColumnWidth(3, 80);  // Status
  sheet.setColumnWidth(4, 50);  // Sec
  sheet.setColumnWidth(5, 280); // Summary
  sheet.setColumnWidth(6, 45);  // #Err
  sheet.setColumnWidth(7, 500); // Error Details
}

// ============================================
// BATCH DETAIL REDIRECT (actual logic in 04b_RawTabs.gs)
// ============================================

/**
 * Redirect function for backward compatibility.
 * Actual writeBatchDetail logic is now writeBatchTab in 04b_RawTabs.gs
 */
function writeBatchDetail(ss, katanaData, waspData, batchData) {
  writeBatchTab(ss, katanaData, waspData, batchData);
}

// ============================================
// HELPER FUNCTIONS
// ============================================

/**
 * Updates an item's lot tracking setting in WASP.
 *
 * @param {string} sku - Item number
 * @param {boolean} lotTracking - Enable or disable lot tracking
 * @returns {Object} { success, code, response }
 */
function waspUpdateItemLotTracking(sku, lotTracking) {
  var url = getWaspBase() + 'ic/item/update';
  var payload = {
    ItemNumber: sku,
    LotTracking: lotTracking ? true : false
  };
  return waspApiCall(url, payload);
}

function getOrCreateSheet_(ss, name, setupFn) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    if (setupFn) setupFn(sheet);
  }
  return sheet;
}

// ============================================
// PROBE LOT TRACKING — Check BOTH Katana and WASP field names
// ============================================

/**
 * Quick diagnostic: fetches a known lot-tracked item from BOTH APIs
 * and logs every field, so we can see the exact field names.
 * Run from menu: "Probe Lot Fields"
 */
function probeLotFields() {
  Logger.log('=== PROBE LOT TRACKING FIELDS ===');
  Logger.log('');

  // ── KATANA: check products + materials for batch tracking fields ──
  Logger.log('═══ KATANA PRODUCTS ═══');
  try {
    var products = katanaFetchAllPages('/products');
    Logger.log('Fetched ' + products.length + ' products');
    if (products.length > 0) {
      Logger.log('ALL FIELDS on first product: ' + Object.keys(products[0]).join(', '));
      // Find fields containing "batch", "track", "lot", "serial"
      var pKeys = Object.keys(products[0]);
      var trackFields = [];
      for (var pk = 0; pk < pKeys.length; pk++) {
        var pLow = pKeys[pk].toLowerCase();
        if (pLow.indexOf('batch') > -1 || pLow.indexOf('track') > -1 || pLow.indexOf('lot') > -1 || pLow.indexOf('serial') > -1) {
          trackFields.push(pKeys[pk]);
        }
      }
      Logger.log('TRACKING-RELATED FIELDS: ' + (trackFields.length > 0 ? trackFields.join(', ') : 'NONE'));

      // Show tracking fields for first 5 products
      for (var p1 = 0; p1 < Math.min(products.length, 5); p1++) {
        var pVals = {};
        for (var tf = 0; tf < trackFields.length; tf++) {
          pVals[trackFields[tf]] = products[p1][trackFields[tf]];
        }
        Logger.log('  ' + (products[p1].name || products[p1].sku || 'product_' + p1) + ': ' + JSON.stringify(pVals));
      }

      // Find specifically lot-tracked products
      Logger.log('');
      Logger.log('Products with batch/lot tracking enabled:');
      var btCount = 0;
      for (var p2 = 0; p2 < products.length; p2++) {
        var anyTrue = false;
        for (var tf2 = 0; tf2 < trackFields.length; tf2++) {
          if (products[p2][trackFields[tf2]]) anyTrue = true;
        }
        if (anyTrue) {
          var pName = products[p2].name || products[p2].sku || '?';
          var pTrackVals = {};
          for (var tf3 = 0; tf3 < trackFields.length; tf3++) {
            pTrackVals[trackFields[tf3]] = products[p2][trackFields[tf3]];
          }
          Logger.log('  ' + pName + ': ' + JSON.stringify(pTrackVals));
          btCount++;
        }
      }
      Logger.log('Total lot-tracked products: ' + btCount + ' / ' + products.length);
    }
  } catch (e) { Logger.log('ERROR: ' + e.message); }

  Logger.log('');
  Logger.log('═══ KATANA MATERIALS ═══');
  try {
    var materials = katanaFetchAllPages('/materials');
    Logger.log('Fetched ' + materials.length + ' materials');
    if (materials.length > 0) {
      Logger.log('ALL FIELDS on first material: ' + Object.keys(materials[0]).join(', '));
      var mKeys = Object.keys(materials[0]);
      var mTrackFields = [];
      for (var mk = 0; mk < mKeys.length; mk++) {
        var mLow = mKeys[mk].toLowerCase();
        if (mLow.indexOf('batch') > -1 || mLow.indexOf('track') > -1 || mLow.indexOf('lot') > -1 || mLow.indexOf('serial') > -1) {
          mTrackFields.push(mKeys[mk]);
        }
      }
      Logger.log('TRACKING-RELATED FIELDS: ' + (mTrackFields.length > 0 ? mTrackFields.join(', ') : 'NONE'));

      // Show tracking fields for first 5 materials
      for (var m1 = 0; m1 < Math.min(materials.length, 5); m1++) {
        var mVals = {};
        for (var mtf = 0; mtf < mTrackFields.length; mtf++) {
          mVals[mTrackFields[mtf]] = materials[m1][mTrackFields[mtf]];
        }
        Logger.log('  ' + (materials[m1].name || 'material_' + m1) + ': ' + JSON.stringify(mVals));
      }

      // List lot-tracked materials
      Logger.log('');
      Logger.log('Materials with batch/lot tracking enabled:');
      var mBtCount = 0;
      for (var m2 = 0; m2 < materials.length; m2++) {
        var mAnyTrue = false;
        for (var mtf2 = 0; mtf2 < mTrackFields.length; mtf2++) {
          if (materials[m2][mTrackFields[mtf2]]) mAnyTrue = true;
        }
        if (mAnyTrue) {
          var mName = materials[m2].name || '?';
          var mtVals = {};
          for (var mtf3 = 0; mtf3 < mTrackFields.length; mtf3++) {
            mtVals[mTrackFields[mtf3]] = materials[m2][mTrackFields[mtf3]];
          }
          Logger.log('  ' + mName + ': ' + JSON.stringify(mtVals));
          mBtCount++;
        }
      }
      Logger.log('Total lot-tracked materials: ' + mBtCount + ' / ' + materials.length);
    }
  } catch (e) { Logger.log('ERROR: ' + e.message); }

  Logger.log('');

  // ── WASP: try infosearch and advancedinventorysearch for known items ──
  Logger.log('═══ WASP ic/item/infosearch ═══');
  try {
    var infoResult = waspFetch('ic/item/infosearch', { PageSize: 5, PageNumber: 1 });
    var infoData = infoResult.Data || [];
    if (infoResult.Data && infoResult.Data.ResultList) infoData = infoResult.Data.ResultList;
    if (infoData.length > 0) {
      Logger.log('Status: WORKING — ' + infoData.length + ' items');
      Logger.log('ALL FIELDS: ' + Object.keys(infoData[0]).join(', '));
      var iKeys = Object.keys(infoData[0]);
      var iTrackFields = [];
      for (var ik = 0; ik < iKeys.length; ik++) {
        var iLow = iKeys[ik].toLowerCase();
        if (iLow.indexOf('lot') > -1 || iLow.indexOf('track') > -1 || iLow.indexOf('batch') > -1 || iLow.indexOf('serial') > -1) {
          iTrackFields.push(iKeys[ik]);
        }
      }
      Logger.log('LOT-RELATED FIELDS: ' + (iTrackFields.length > 0 ? iTrackFields.join(', ') : 'NONE'));
      for (var i1 = 0; i1 < Math.min(infoData.length, 3); i1++) {
        var iVals = {};
        for (var itf = 0; itf < iTrackFields.length; itf++) {
          iVals[iTrackFields[itf]] = infoData[i1][iTrackFields[itf]];
        }
        Logger.log('  ' + (infoData[i1].ItemNumber || '?') + ': ' + JSON.stringify(iVals));
      }
    } else {
      Logger.log('Status: WORKING but empty');
    }
  } catch (e) {
    Logger.log('Status: FAILED — ' + e.message);
  }

  Logger.log('');
  Logger.log('═══ WASP advancedinventorysearch (B-PROP — known lot tracked) ═══');
  try {
    var bpResult = waspFetch('ic/item/advancedinventorysearch', {
      PageSize: 5, PageNumber: 1, ItemNumber: 'B-PROP'
    });
    var bpData = bpResult.Data || [];
    if (bpData.length > 0) {
      var bpKeys = Object.keys(bpData[0]);
      var bpTrackFields = [];
      for (var bk = 0; bk < bpKeys.length; bk++) {
        var bLow = bpKeys[bk].toLowerCase();
        if (bLow.indexOf('lot') > -1 || bLow.indexOf('track') > -1 || bLow.indexOf('batch') > -1 || bLow.indexOf('serial') > -1 || bLow.indexOf('date') > -1) {
          bpTrackFields.push(bpKeys[bk]);
        }
      }
      Logger.log('LOT-RELATED FIELDS: ' + (bpTrackFields.length > 0 ? bpTrackFields.join(', ') : 'NONE'));
      var bpVals = {};
      for (var btf = 0; btf < bpTrackFields.length; btf++) {
        bpVals[bpTrackFields[btf]] = bpData[0][bpTrackFields[btf]];
      }
      Logger.log('  B-PROP values: ' + JSON.stringify(bpVals));
    } else {
      Logger.log('  B-PROP: NOT FOUND');
    }
  } catch (e) { Logger.log('ERROR: ' + e.message); }

  Logger.log('');
  Logger.log('=== PROBE COMPLETE ===');
  Logger.log('Look for TRACKING-RELATED FIELDS above.');
  Logger.log('Katana: field should be "batch_tracked" (boolean)');
  Logger.log('WASP: field may be "TrackLot" or "LotTracking" if present');

  try {
    SpreadsheetApp.getUi().alert('Probe complete — check View → Logs for exact field names from both APIs.');
  } catch (e) {}
}

// ============================================
// PROBE WASP LOT TRACKING FIELDS — Find how WASP exposes lot tracking
// ============================================

/**
 * Probes WASP API endpoints to discover which ones return lot tracking info.
 * Tests known lot-tracked items (from Katana) and logs every field.
 * Run from menu: "Probe WASP Lot Fields"
 */
function probeWaspLotFields() {
  var token = getStoredWaspToken_();
  var base = getWaspBase();

  // Known lot-tracked SKUs from Katana — use these to spot the field
  var testSkus = ['B-PROP', 'EO-LAV', 'EGG-X', 'FGJ-US-1', 'LCP-1'];
  // Known NOT lot-tracked — for comparison
  var noLotSkus = ['NI-B221210', 'NI-SEAL-2OZ'];

  Logger.log('=== PROBE WASP LOT TRACKING FIELDS ===');
  Logger.log('Looking for: LotTracking, TrackLot, IsLotTracked, BatchTracking, or similar');
  Logger.log('Test lot-tracked SKUs: ' + testSkus.join(', '));
  Logger.log('Test non-lot SKUs: ' + noLotSkus.join(', '));
  Logger.log('');

  // ── Endpoint 1: advancedinventorysearch (known working) ──
  Logger.log('═══ ENDPOINT: ic/item/advancedinventorysearch ═══');
  try {
    var result1 = waspFetch('ic/item/advancedinventorysearch', { PageSize: 5, PageNumber: 1 });
    var data1 = result1.Data || [];
    if (data1.length > 0) {
      var keys1 = Object.keys(data1[0]);
      Logger.log('Status: WORKING — ' + data1.length + ' rows, ' + keys1.length + ' fields');
      Logger.log('ALL FIELDS: ' + keys1.join(', '));

      // Find lot-related fields
      var lotFields1 = [];
      for (var k1 = 0; k1 < keys1.length; k1++) {
        var lower = keys1[k1].toLowerCase();
        if (lower.indexOf('lot') > -1 || lower.indexOf('track') > -1 || lower.indexOf('batch') > -1 || lower.indexOf('serial') > -1) {
          lotFields1.push(keys1[k1]);
        }
      }
      Logger.log('LOT-RELATED FIELDS: ' + (lotFields1.length > 0 ? lotFields1.join(', ') : 'NONE FOUND'));

      // Show values for first 3 items
      for (var s1 = 0; s1 < Math.min(data1.length, 3); s1++) {
        var item1 = data1[s1];
        var sku1 = item1.ItemNumber || '?';
        var lotVals = {};
        for (var lf1 = 0; lf1 < lotFields1.length; lf1++) {
          lotVals[lotFields1[lf1]] = item1[lotFields1[lf1]];
        }
        Logger.log('  Sample ' + sku1 + ': ' + JSON.stringify(lotVals));
      }
    } else {
      Logger.log('Status: WORKING but empty');
    }
  } catch (e) {
    Logger.log('Status: FAILED — ' + e.message);
  }
  Logger.log('');

  // ── Now search for specific SKUs to compare lot vs non-lot ──
  Logger.log('═══ SEARCH SPECIFIC SKUs via advancedinventorysearch ═══');
  var allCheckSkus = testSkus.concat(noLotSkus);
  for (var cs = 0; cs < allCheckSkus.length; cs++) {
    try {
      var searchResult = waspFetch('ic/item/advancedinventorysearch', {
        PageSize: 5, PageNumber: 1, ItemNumber: allCheckSkus[cs]
      });
      var searchData = searchResult.Data || [];
      if (searchData.length > 0) {
        var sItem = searchData[0];
        // Log ALL fields for this item
        var sKeys = Object.keys(sItem);
        var lotRelated = {};
        for (var sk = 0; sk < sKeys.length; sk++) {
          var sLower = sKeys[sk].toLowerCase();
          if (sLower.indexOf('lot') > -1 || sLower.indexOf('track') > -1 || sLower.indexOf('batch') > -1 || sLower.indexOf('serial') > -1 || sLower.indexOf('date') > -1) {
            lotRelated[sKeys[sk]] = sItem[sKeys[sk]];
          }
        }
        Logger.log('  ' + allCheckSkus[cs] + ' → ' + JSON.stringify(lotRelated));
      } else {
        Logger.log('  ' + allCheckSkus[cs] + ' → NOT FOUND (0 results)');
      }
    } catch (e) {
      Logger.log('  ' + allCheckSkus[cs] + ' → ERROR: ' + e.message);
    }
  }
  Logger.log('');

  // ── Try other item-level endpoints ──
  var endpoints = [
    { name: 'ic/item/search', payload: { PageSize: 5, PageNumber: 1 } },
    { name: 'ic/item/itemsearch', payload: { PageSize: 5, PageNumber: 1 } },
    { name: 'ic/item/detailsearch', payload: { PageSize: 5, PageNumber: 1 } },
    { name: 'ic/item/get', payload: { ItemNumber: 'B-PROP' } },
    { name: 'ic/item/getbyitemnumber', payload: { ItemNumbers: ['B-PROP'] } },
    { name: 'ic/item/info', payload: { ItemNumber: 'B-PROP' } },
    { name: 'ic/item/detail', payload: { ItemNumber: 'B-PROP' } },
    { name: 'ic/item/lookup', payload: { ItemNumber: 'B-PROP' } },
    { name: 'ic/item/find', payload: { ItemNumber: 'B-PROP' } },
    { name: 'ic/item/list', payload: { PageSize: 5, PageNumber: 1 } },
    { name: 'ic/item/catalog', payload: { PageSize: 5, PageNumber: 1 } },
    { name: 'ic/item/inventorysearch', payload: { PageSize: 5, PageNumber: 1 } },
    { name: 'ic/item/getitem', payload: { ItemNumber: 'B-PROP' } },
    { name: 'ic/item/getitems', payload: { PageSize: 5, PageNumber: 1 } }
  ];

  for (var ei = 0; ei < endpoints.length; ei++) {
    var ep = endpoints[ei];
    Logger.log('═══ ENDPOINT: ' + ep.name + ' ═══');
    try {
      var options = {
        method: 'POST',
        headers: { 'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json' },
        payload: JSON.stringify(ep.payload),
        muteHttpExceptions: true
      };
      var response = UrlFetchApp.fetch(base + ep.name, options);
      var code = response.getResponseCode();
      var body = response.getContentText();
      var isHtml = body.indexOf('<!DOCTYPE') > -1 || body.indexOf('<html') > -1;

      if (code !== 200) {
        Logger.log('Status: ' + code + (code === 404 ? ' NOT FOUND' : ''));
        continue;
      }
      if (isHtml) {
        Logger.log('Status: 200 but HTML page (not API)');
        continue;
      }

      var parsed = JSON.parse(body);
      var epData = parsed.Data || parsed.data || [];
      if (epData && epData.ResultList) epData = epData.ResultList;
      if (!Array.isArray(epData)) {
        // Might be a single object
        Logger.log('Status: 200 JSON — non-array response');
        Logger.log('Top-level keys: ' + Object.keys(parsed).join(', '));
        Logger.log('Response (first 500): ' + body.substring(0, 500));
        // Check if it has lot-related fields at top level
        var topKeys = Object.keys(parsed);
        var topLot = [];
        for (var tk = 0; tk < topKeys.length; tk++) {
          var tLow = topKeys[tk].toLowerCase();
          if (tLow.indexOf('lot') > -1 || tLow.indexOf('track') > -1 || tLow.indexOf('batch') > -1) {
            topLot.push(topKeys[tk] + '=' + JSON.stringify(parsed[topKeys[tk]]));
          }
        }
        if (topLot.length > 0) Logger.log('LOT FIELDS IN RESPONSE: ' + topLot.join(', '));
        continue;
      }

      if (epData.length === 0) {
        Logger.log('Status: 200 JSON — empty Data array');
        Logger.log('Response keys: ' + Object.keys(parsed).join(', '));
        continue;
      }

      // We have data!
      var epKeys = Object.keys(epData[0]);
      Logger.log('Status: WORKING — ' + epData.length + ' rows, ' + epKeys.length + ' fields');
      Logger.log('ALL FIELDS: ' + epKeys.join(', '));

      var lotFields = [];
      for (var lk = 0; lk < epKeys.length; lk++) {
        var epLower = epKeys[lk].toLowerCase();
        if (epLower.indexOf('lot') > -1 || epLower.indexOf('track') > -1 || epLower.indexOf('batch') > -1 || epLower.indexOf('serial') > -1) {
          lotFields.push(epKeys[lk]);
        }
      }
      Logger.log('LOT-RELATED FIELDS: ' + (lotFields.length > 0 ? lotFields.join(', ') : 'NONE'));

      // Show first 2 items with lot fields
      for (var ei2 = 0; ei2 < Math.min(epData.length, 2); ei2++) {
        var eItem = epData[ei2];
        var eSku = eItem.ItemNumber || eItem.itemNumber || '?';
        var eLotVals = {};
        for (var el = 0; el < lotFields.length; el++) {
          eLotVals[lotFields[el]] = eItem[lotFields[el]];
        }
        Logger.log('  ' + eSku + ': ' + JSON.stringify(eLotVals));
      }

    } catch (e) {
      Logger.log('Status: ERROR — ' + e.message);
    }
    Logger.log('');
  }

  Logger.log('=== PROBE COMPLETE ===');
  Logger.log('Look above for endpoints that returned LOT-RELATED FIELDS.');
  Logger.log('The correct endpoint + field name is what we need for accurate lot tracking display.');

  try {
    SpreadsheetApp.getUi().alert(
      'Probe complete — check View → Logs.\n\n' +
      'Look for "LOT-RELATED FIELDS" entries to find which endpoint returns lot tracking data.'
    );
  } catch (e) {}
}

// ============================================
// AUDIT LOT TRACKING — Quick mismatch finder
// ============================================

/**
 * Compares Katana lot tracking (source of truth) vs WASP settings.
 * Logs items that need lot tracking ENABLED or DISABLED in WASP.
 * Also shows results in a toast + alert for quick reference.
 * Run from menu: "Audit Lot Tracking"
 */
function auditLotTracking() {
  Logger.log('=== AUDIT LOT TRACKING ===');
  Logger.log('Katana = source of truth');
  Logger.log('');

  try { SpreadsheetApp.getActiveSpreadsheet().toast('Auditing lot tracking...', 'Please wait', 60); } catch (e) {}

  // 1. Fetch Katana data + batch info
  var katanaData = fetchAllKatanaData();
  var batchData = fetchAllBatchData(katanaData);

  // 2. Get batch_tracking from Katana products/materials
  var productBatch = {};
  var materialBatch = {};
  try {
    var products = katanaFetchAllPages('/products');
    for (var pi = 0; pi < products.length; pi++) {
      var prod = products[pi];
      var bt = prod.batch_tracking || prod.batch_tracked || prod.is_batch_tracked || prod.track_batches || false;
      if (prod.id) productBatch[prod.id] = bt ? true : false;
    }
  } catch (e) { Logger.log('Product fetch error: ' + e.message); }

  try {
    var materials = katanaFetchAllPages('/materials');
    for (var mi = 0; mi < materials.length; mi++) {
      var mat = materials[mi];
      var mbt = mat.batch_tracking || mat.batch_tracked || mat.is_batch_tracked || mat.track_batches || false;
      if (mat.id) materialBatch[mat.id] = mbt ? true : false;
    }
  } catch (e) { Logger.log('Material fetch error: ' + e.message); }

  // 3. Map variants to batch_tracking
  var variantBatchMap = {};
  try {
    var variants = katanaFetchAllPages('/variants');
    for (var vi = 0; vi < variants.length; vi++) {
      var v = variants[vi];
      if (!v.sku) continue;
      var vbt = false;
      if (v.product_id && productBatch[v.product_id]) vbt = true;
      if (v.material_id && materialBatch[v.material_id]) vbt = true;
      variantBatchMap[v.sku] = vbt;
    }
  } catch (e) { Logger.log('Variant map error: ' + e.message); }

  // 4. Build Katana lot tracking map (unique SKUs)
  var katanaLot = {};
  for (var i = 0; i < katanaData.items.length; i++) {
    var item = katanaData.items[i];
    if (!item.sku || isSkippedSku(item.sku) || katanaLot.hasOwnProperty(item.sku)) continue;
    var kLot = false;
    if (variantBatchMap.hasOwnProperty(item.sku)) kLot = variantBatchMap[item.sku];
    if (batchData.lotTrackedSkus[item.sku]) kLot = true;
    katanaLot[item.sku] = kLot;
  }

  // 5. Fetch WASP lot settings
  var waspLot = fetchWaspItemLotSettings(true);

  // 6. Compare
  var needEnable = [];   // Katana=Yes, WASP=No → need to ENABLE in WASP
  var needDisable = [];  // Katana=No, WASP=Yes → need to DISABLE in WASP
  var notInWasp = [];    // Katana item not in WASP at all
  var matchYes = [];     // Both Yes
  var matchNo = [];      // Both No

  var katanaSkus = Object.keys(katanaLot).sort();
  for (var si = 0; si < katanaSkus.length; si++) {
    var sku = katanaSkus[si];
    var kVal = katanaLot[sku];
    if (!waspLot.hasOwnProperty(sku)) {
      notInWasp.push(sku);
      continue;
    }
    var wVal = waspLot[sku];
    if (kVal && !wVal) {
      needEnable.push(sku);
    } else if (!kVal && wVal) {
      needDisable.push(sku);
    } else if (kVal && wVal) {
      matchYes.push(sku);
    } else {
      matchNo.push(sku);
    }
  }

  // 7. Log results
  Logger.log('--- NEED LOT TRACKING ENABLED IN WASP (' + needEnable.length + ') ---');
  Logger.log('These items are lot-tracked in Katana but NOT in WASP:');
  for (var ne = 0; ne < needEnable.length; ne++) {
    Logger.log('  ENABLE: ' + needEnable[ne]);
  }
  Logger.log('');

  Logger.log('--- NEED LOT TRACKING DISABLED IN WASP (' + needDisable.length + ') ---');
  Logger.log('These items are NOT lot-tracked in Katana but ARE in WASP:');
  for (var nd = 0; nd < needDisable.length; nd++) {
    Logger.log('  DISABLE: ' + needDisable[nd]);
  }
  Logger.log('');

  Logger.log('--- ALREADY CORRECT (' + (matchYes.length + matchNo.length) + ') ---');
  Logger.log('Lot=Yes match: ' + matchYes.length + ' items: ' + matchYes.join(', '));
  Logger.log('Lot=No match: ' + matchNo.length + ' items');
  Logger.log('');

  Logger.log('--- NOT IN WASP (' + notInWasp.length + ') ---');
  if (notInWasp.length > 0) Logger.log(notInWasp.join(', '));
  Logger.log('');

  Logger.log('=== SUMMARY ===');
  Logger.log('Total Katana SKUs: ' + katanaSkus.length);
  Logger.log('Need ENABLE in WASP: ' + needEnable.length);
  Logger.log('Need DISABLE in WASP: ' + needDisable.length);
  Logger.log('Already correct: ' + (matchYes.length + matchNo.length));
  Logger.log('Not in WASP: ' + notInWasp.length);

  // 8. Show alert
  var msg = 'LOT TRACKING AUDIT\n\n';
  if (needEnable.length > 0) {
    msg += 'ENABLE lot tracking (' + needEnable.length + '):\n' + needEnable.join('\n') + '\n\n';
  }
  if (needDisable.length > 0) {
    msg += 'DISABLE lot tracking (' + needDisable.length + '):\n' + needDisable.join('\n') + '\n\n';
  }
  msg += 'Already correct: ' + (matchYes.length + matchNo.length) + '\n';
  msg += 'Not in WASP: ' + notInWasp.length + '\n\n';
  msg += 'Full details in View → Logs';

  try { SpreadsheetApp.getUi().alert(msg); } catch (e) {}
}

// ============================================
// LOT TRACKING REFERENCE TAB
// Creates comparison tab: Katana vs WASP lot settings
// ============================================

/**
 * Creates/updates a "Lot Tracking" reference tab showing lot tracking status
 * for all items. Katana = source of truth. Highlights mismatches with WASP.
 * Run from menu: "Write Lot Tracking Tab"
 */
function writeLotTrackingTab() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.toast('Building Lot Tracking tab... (fetching data)', 'Please wait', 60);

  // 1. Fetch Katana data + batch tracking info
  var katanaData = fetchAllKatanaData();
  var batchData = fetchAllBatchData(katanaData);

  // 2. Get batch_tracking setting from Katana products and materials
  var katanaBatchSetting = {};
  try {
    var products = katanaFetchAllPages('/products');
    for (var pi = 0; pi < products.length; pi++) {
      var prod = products[pi];
      var bt = prod.batch_tracking || prod.batch_tracked || prod.is_batch_tracked || prod.track_batches || false;
      if (prod.id) katanaBatchSetting['product_' + prod.id] = bt ? true : false;
    }
    Logger.log('Lot tab: fetched ' + products.length + ' products for batch_tracking');
  } catch (e) { Logger.log('Product batch_tracking fetch error: ' + e.message); }

  try {
    var materials = katanaFetchAllPages('/materials');
    for (var mi = 0; mi < materials.length; mi++) {
      var mat = materials[mi];
      var mbt = mat.batch_tracking || mat.batch_tracked || mat.is_batch_tracked || mat.track_batches || false;
      if (mat.id) katanaBatchSetting['material_' + mat.id] = mbt ? true : false;
    }
    Logger.log('Lot tab: fetched ' + materials.length + ' materials for batch_tracking');
  } catch (e) { Logger.log('Material batch_tracking fetch error: ' + e.message); }

  // 3. Map variants to batch_tracking via product/material parent
  var variantBatchMap = {};
  try {
    var variants = katanaFetchAllPages('/variants');
    for (var vi = 0; vi < variants.length; vi++) {
      var v = variants[vi];
      if (!v.sku) continue;
      var vbt = false;
      if (v.product_id && katanaBatchSetting['product_' + v.product_id]) vbt = true;
      if (v.material_id && katanaBatchSetting['material_' + v.material_id]) vbt = true;
      variantBatchMap[v.sku] = vbt;
    }
    Logger.log('Lot tab: mapped ' + Object.keys(variantBatchMap).length + ' variant SKUs');
  } catch (e) { Logger.log('Variant batch map error: ' + e.message); }

  // 4. Fetch WASP item lot settings
  var waspLotSettings = fetchWaspItemLotSettings(true);

  // 5. Build unique SKU list from Katana items
  var skuInfo = {};
  for (var i = 0; i < katanaData.items.length; i++) {
    var item = katanaData.items[i];
    if (!item.sku || isSkippedSku(item.sku)) continue;
    if (skuInfo[item.sku]) continue;
    // Katana lot tracking: product/material API setting OR has batch_stock records
    var katanaLot = false;
    if (variantBatchMap.hasOwnProperty(item.sku)) {
      katanaLot = variantBatchMap[item.sku];
    }
    if (batchData.lotTrackedSkus[item.sku]) {
      katanaLot = true;
    }
    skuInfo[item.sku] = { name: item.name, type: item.type, katanaLot: katanaLot };
  }

  // 6. Build rows
  var rows = [];
  var skus = Object.keys(skuInfo).sort();
  var mismatchCount = 0;
  var notInWaspCount = 0;
  var matchCount = 0;

  for (var si = 0; si < skus.length; si++) {
    var sku = skus[si];
    var info = skuInfo[sku];
    var waspLot = waspLotSettings.hasOwnProperty(sku) ? waspLotSettings[sku] : null;
    var katanaLot = info.katanaLot;
    var status = '';
    if (waspLot === null) {
      status = 'NOT IN WASP';
      notInWaspCount++;
    } else if (katanaLot === waspLot) {
      status = 'MATCH';
      matchCount++;
    } else {
      status = 'MISMATCH';
      mismatchCount++;
    }
    rows.push([
      sku,
      info.name,
      info.type,
      katanaLot ? 'Yes' : 'No',
      katanaLot ? 'Yes' : 'No',
      waspLot === null ? 'N/A' : (waspLot ? 'Yes' : 'No'),
      status
    ]);
  }

  // Also add WASP-only items (exist in WASP but not in Katana tracked locations)
  var waspOnlyCount = 0;
  var waspSkus = Object.keys(waspLotSettings);
  for (var ws = 0; ws < waspSkus.length; ws++) {
    if (!skuInfo[waspSkus[ws]] && !isSkippedSku(waspSkus[ws])) {
      rows.push([
        waspSkus[ws],
        '',
        '',
        'N/A',
        'N/A',
        waspLotSettings[waspSkus[ws]] ? 'Yes' : 'No',
        'WASP ONLY'
      ]);
      waspOnlyCount++;
    }
  }

  // 7. Create or clear sheet
  var sheet = ss.getSheetByName('Lot Tracking');
  if (sheet) {
    sheet.clear();
    sheet.clearFormats();
    sheet.clearConditionalFormatRules();
  } else {
    sheet = ss.insertSheet('Lot Tracking');
  }

  // 8. Write headers
  var headers = ['SKU', 'Name', 'Type', 'Katana Lot', 'Katana DateCode', 'WASP Lot', 'Status'];
  sheet.getRange(1, 1, 1, 7).setValues([headers]);
  sheet.getRange(1, 1, 1, 7).setFontWeight('bold').setBackground('#424242').setFontColor('#ffffff');
  sheet.setFrozenRows(1);

  // 9. Write data rows
  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, 7).setValues(rows);
  }

  // 10. Column widths
  sheet.setColumnWidth(1, 140);
  sheet.setColumnWidth(2, 260);
  sheet.setColumnWidth(3, 100);
  sheet.setColumnWidth(4, 100);
  sheet.setColumnWidth(5, 120);
  sheet.setColumnWidth(6, 100);
  sheet.setColumnWidth(7, 120);

  // 11. Conditional formatting on Status column (col 7)
  var numRows = Math.max(rows.length, 1);
  var statusRange = sheet.getRange(2, 7, numRows, 1);
  var katanaLotRange = sheet.getRange(2, 4, numRows, 2);
  var rules = [];

  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('MISMATCH')
    .setBackground('#ffcdd2')
    .setFontColor('#b71c1c')
    .setRanges([statusRange])
    .build());

  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('MATCH')
    .setBackground('#c8e6c9')
    .setRanges([statusRange])
    .build());

  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('NOT IN WASP')
    .setBackground('#ffe0b2')
    .setRanges([statusRange])
    .build());

  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('WASP ONLY')
    .setBackground('#bbdefb')
    .setRanges([statusRange])
    .build());

  // "Yes" in Katana Lot / DateCode columns — light blue highlight
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Yes')
    .setBackground('#e3f2fd')
    .setRanges([katanaLotRange])
    .build());

  sheet.setConditionalFormatRules(rules);
  SpreadsheetApp.flush();

  var summary = rows.length + ' items: ' + matchCount + ' match, ' + mismatchCount + ' mismatch, ' + notInWaspCount + ' not in WASP, ' + waspOnlyCount + ' WASP-only';
  Logger.log('Lot Tracking tab written: ' + summary);
  try {
    ss.toast('Lot Tracking tab complete: ' + summary, 'Done', 10);
  } catch (e) {}
}

// ============================================
// DIAGNOSTIC: Test ic/item/create endpoint
// Run from GAS editor to debug 404 errors.
// ============================================

/**
 * Tests WASP ic/item/create with multiple payload variants.
 * Run from GAS editor → View → Logs to see results.
 *
 * Tests:
 * 1. Minimal payload (matches Python sync_all_v5.py)
 * 2. Current GAS payload (includes SiteName, LocationCode, LotTracking)
 * 3. Without /public-api/ prefix (alternate URL)
 * 4. Payload wrapped in array (like transaction endpoints)
 */
function testWaspCreateEndpoint() {
  var testSku = 'TEST-DELETE-ME-' + new Date().getTime();
  var base = getWaspBase(); // https://...waspinventorycloud.com/public-api/
  var instance = PropertiesService.getScriptProperties().getProperty('WASP_INSTANCE') || 'mymagichealer';
  var altBase = 'https://' + instance + '.waspinventorycloud.com/';
  var token = getStoredWaspToken_();

  Logger.log('=== WASP ic/item/create DIAGNOSTIC ===');
  Logger.log('Test SKU: ' + testSku);
  Logger.log('Base URL (GAS): ' + base);
  Logger.log('Alt URL (no public-api): ' + altBase);
  Logger.log('Token present: ' + (token ? 'YES (' + token.length + ' chars)' : 'NO'));
  Logger.log('');

  // Payload A: Minimal (matches Python sync_all_v5.py)
  var payloadMinimal = {
    ItemNumber: testSku,
    Description: 'Diagnostic test item',
    Category: 'FINISHED GOODS',
    ItemType: 'Inventory',
    Cost: 0,
    SalePrice: 0,
    Price: 0,
    ListPrice: 0,
    Active: true,
    Taxable: true,
    Notes: 'Diagnostic test ' + new Date().toISOString()
  };

  // Payload B: Current GAS payload (with site/location/lot fields)
  var payloadGAS = {
    ItemNumber: testSku + '-B',
    Description: 'Diagnostic test item B',
    SiteName: 'MMH Kelowna',
    LocationCode: 'UNSORTED',
    ItemType: 'Inventory',
    Active: true,
    LotTracking: false
  };

  // Payload C: Even more minimal — just ItemNumber + Description + ItemType
  var payloadBare = {
    ItemNumber: testSku + '-C',
    Description: 'Diagnostic test item C',
    ItemType: 'Inventory'
  };

  var tests = [
    { name: 'Test 1: Python-style payload, /public-api/', url: base + 'ic/item/create', payload: payloadMinimal },
    { name: 'Test 2: GAS payload (site+location+lot), /public-api/', url: base + 'ic/item/create', payload: payloadGAS },
    { name: 'Test 3: Bare payload, /public-api/', url: base + 'ic/item/create', payload: payloadBare },
    { name: 'Test 4: Python-style, NO /public-api/', url: altBase + 'ic/item/create', payload: payloadMinimal },
    { name: 'Test 5: Python-style, array-wrapped', url: base + 'ic/item/create', payload: [payloadMinimal] },
    { name: 'Test 6: Verify working endpoint (advancedinventorysearch)', url: base + 'ic/item/advancedinventorysearch', payload: { PageSize: 1, PageNumber: 1 } }
  ];

  for (var t = 0; t < tests.length; t++) {
    var test = tests[t];
    Logger.log('--- ' + test.name + ' ---');
    Logger.log('URL: ' + test.url);
    Logger.log('Payload: ' + JSON.stringify(test.payload).substring(0, 300));

    try {
      var options = {
        method: 'POST',
        headers: { 'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json' },
        payload: JSON.stringify(test.payload),
        muteHttpExceptions: true
      };
      var response = UrlFetchApp.fetch(test.url, options);
      var code = response.getResponseCode();
      var body = response.getContentText().substring(0, 500);
      var headers = response.getHeaders();

      Logger.log('Status: ' + code);
      Logger.log('Response: ' + body);
      if (code === 404) {
        Logger.log('Content-Type: ' + (headers['Content-Type'] || headers['content-type'] || 'unknown'));
      }
      if (code === 200) {
        Logger.log('SUCCESS — this payload/URL combination works');
      }
    } catch (e) {
      Logger.log('EXCEPTION: ' + e.message);
    }
    Logger.log('');
  }

  // Cleanup: try to delete test items
  Logger.log('--- Cleanup: deleting test items ---');
  var deleteSkus = [testSku, testSku + '-B', testSku + '-C'];
  for (var d = 0; d < deleteSkus.length; d++) {
    try {
      var delResult = waspApiCall(base + 'ic/item/delete', { ItemNumber: deleteSkus[d] });
      Logger.log('Delete ' + deleteSkus[d] + ': ' + (delResult.success ? 'OK' : 'skip'));
    } catch (e) {
      Logger.log('Delete ' + deleteSkus[d] + ': ' + e.message);
    }
  }

  Logger.log('=== DIAGNOSTIC COMPLETE ===');
  Logger.log('Check results above. If Test 1 or 3 succeeds but Test 2 fails,');
  Logger.log('the cause is SiteName/LocationCode/LotTracking fields in payload.');
  Logger.log('If Test 4 succeeds, the /public-api/ prefix is the issue.');
  Logger.log('If Test 5 succeeds, the endpoint needs array-wrapped payload.');

  try {
    SpreadsheetApp.getUi().alert(
      'Diagnostic complete — check View → Logs for results.\n\n' +
      'Look for which test returned Status: 200 (SUCCESS).'
    );
  } catch (e) { /* timer context */ }
}

/**
 * Probes WASP API for alternative item create endpoints.
 * ic/item/create is 404 as of Feb 20, 2026 — need to find replacement.
 * Run from GAS editor → View → Logs.
 */
function discoverWaspCreateEndpoint() {
  var token = getStoredWaspToken_();
  var base = getWaspBase();
  var testSku = 'TEST-PROBE-' + new Date().getTime();

  var payload = {
    ItemNumber: testSku,
    Description: 'Endpoint probe test',
    ItemType: 'Inventory',
    Active: true
  };

  // Candidate endpoints to try
  var candidates = [
    'ic/item/create',
    'ic/item/add',
    'ic/item/insert',
    'ic/item/new',
    'ic/item/save',
    'ic/item/upsert',
    'ic/item/createitem',
    'ic/item/additem',
    'ic/item/createorupdate',
    'ic/items/create',
    'ic/items/add',
    'inventory/item/create',
    'inventory/item/add',
    'items/create',
    'items/add',
    'item/create',
    'item/add',
    'ic/item/update',
    'ic/item/edit'
  ];

  Logger.log('=== WASP ENDPOINT DISCOVERY ===');
  Logger.log('Testing ' + candidates.length + ' candidate endpoints...');
  Logger.log('Test SKU: ' + testSku);
  Logger.log('');

  var found = [];

  for (var i = 0; i < candidates.length; i++) {
    var endpoint = candidates[i];
    var url = base + endpoint;

    try {
      var options = {
        method: 'POST',
        headers: { 'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json' },
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
      };
      var response = UrlFetchApp.fetch(url, options);
      var code = response.getResponseCode();
      var body = response.getContentText().substring(0, 300);
      var isHtml = body.indexOf('<!DOCTYPE') > -1 || body.indexOf('<html') > -1;
      var isJson = false;
      try { JSON.parse(response.getContentText()); isJson = true; } catch (e) {}

      var status = code + (isJson ? ' JSON' : (isHtml ? ' HTML' : ' ???'));

      if (code === 200 && isJson) {
        Logger.log('>>> HIT: ' + endpoint + ' → ' + status);
        Logger.log('    Response: ' + body);
        found.push(endpoint);
      } else if (code === 200 && isHtml) {
        Logger.log('    skip: ' + endpoint + ' → ' + status + ' (web page, not API)');
      } else if (code === 404) {
        Logger.log('    miss: ' + endpoint + ' → 404');
      } else if (code === 400 || code === 422) {
        // 400/422 means endpoint EXISTS but payload was rejected — good signal!
        Logger.log('>>> POSSIBLE: ' + endpoint + ' → ' + status + ' (endpoint exists, payload rejected)');
        Logger.log('    Response: ' + body);
        found.push(endpoint + ' (needs correct payload)');
      } else {
        Logger.log('    other: ' + endpoint + ' → ' + status);
        Logger.log('    Response: ' + body.substring(0, 150));
      }
    } catch (e) {
      Logger.log('    error: ' + endpoint + ' → ' + e.message);
    }
  }

  Logger.log('');
  Logger.log('=== DISCOVERY RESULTS ===');
  if (found.length > 0) {
    Logger.log('Found ' + found.length + ' potential endpoints:');
    for (var f = 0; f < found.length; f++) {
      Logger.log('  ' + found[f]);
    }
  } else {
    Logger.log('No working create endpoints found in /public-api/ namespace.');
    Logger.log('WASP may have moved item creation to a different API version or path.');
    Logger.log('Next steps:');
    Logger.log('  1. Check WASP admin UI for API documentation link');
    Logger.log('  2. Contact WASP support about ic/item/create removal');
    Logger.log('  3. Try WASP Inventory Import CSV as workaround');
  }

  // Cleanup
  for (var d = 0; d < found.length; d++) {
    if (found[d].indexOf('needs correct') === -1) {
      try {
        waspApiCall(base + 'ic/item/delete', { ItemNumber: testSku });
      } catch (e) {}
    }
  }

  Logger.log('=== DISCOVERY COMPLETE ===');

  try {
    var msg = found.length > 0
      ? 'Found ' + found.length + ' potential endpoints! Check View → Logs.'
      : 'No create endpoints found. Check View → Logs for next steps.';
    SpreadsheetApp.getUi().alert(msg);
  } catch (e) {}
}
