// ============================================
// 98_ItemComparison.gs - ITEM COMPARISON UTILITY
// ============================================
// Pulls items from Katana and compares with WASP CSV data.
// Populates the "WASP-Katana Item Comparison" Google Sheet.
// DELETE THIS FILE when no longer needed.
// ============================================

var COMPARISON_SHEET_ID = '1mqdZ1Yp9fIzpSMxlJYgfbYGSe1ii2N836L1stc4AtNU';

// ============================================
// STEP 1: POPULATE WASP RAW DATA
// ============================================
// Reads the WASP CSV from Downloads (or uses current WASP export)
// and writes to the "WASP Raw" tab.
// ============================================

/**
 * Populate the WASP Raw tab from the existing Items CSV data.
 * User should paste CSV data into "WASP Raw" tab column A first,
 * OR run this function which reads from WASP API (item lots).
 */

// ============================================
// STEP 2: PULL KATANA ITEMS & BUILD COMPARISON
// ============================================

/**
 * PREVIEW — Shows how many items in Katana vs WASP
 */
function itemCompare_PREVIEW() {
  var katanaItems = fetchAllKatanaVariants();
  var waspData = readWaspRawTab();

  Logger.log('===== ITEM COMPARISON PREVIEW =====');
  Logger.log('Katana variants found: ' + katanaItems.length);
  Logger.log('WASP unique SKUs: ' + Object.keys(waspData).length);
  Logger.log('');

  // Find mismatches
  var onlyKatana = [];
  var onlyWasp = [];
  var inBoth = [];

  for (var i = 0; i < katanaItems.length; i++) {
    var kItem = katanaItems[i];
    if (waspData[kItem.sku]) {
      inBoth.push(kItem.sku);
    } else {
      onlyKatana.push(kItem.sku);
    }
  }

  for (var sku in waspData) {
    var found = false;
    for (var j = 0; j < katanaItems.length; j++) {
      if (katanaItems[j].sku === sku) { found = true; break; }
    }
    if (!found) onlyWasp.push(sku);
  }

  Logger.log('In BOTH systems: ' + inBoth.length);
  Logger.log('Only in Katana: ' + onlyKatana.length);
  Logger.log('Only in WASP: ' + onlyWasp.length);
  Logger.log('');

  if (onlyKatana.length > 0) {
    Logger.log('--- Only in Katana (missing from WASP) ---');
    for (var k = 0; k < onlyKatana.length; k++) {
      Logger.log('  ' + onlyKatana[k]);
    }
  }

  if (onlyWasp.length > 0) {
    Logger.log('');
    Logger.log('--- Only in WASP (missing from Katana) ---');
    for (var w = 0; w < onlyWasp.length; w++) {
      Logger.log('  ' + onlyWasp[w]);
    }
  }

  Logger.log('');
  Logger.log('Run itemCompare_POPULATE() to write comparison to the sheet.');
}

/**
 * POPULATE — Writes full comparison to the Comparison tab
 *
 * Katana UOM source of truth (in priority order):
 *   1. KATANA_CONVERSION_ITEMS — hardcoded stock UOM for items with purchase→stock conversions
 *      (e.g. B-WAX purchased in kg but stocked in g — Katana API returns kg, override gives g)
 *   2. materials endpoint — for raw material variants (variant.material_id is set)
 *   3. products endpoint  — for finished-good variants (variant.product_id is set)
 *
 * WASP UOM comes from WASP Raw tab column 7 (StockingUnit field from WASP export).
 */
function itemCompare_POPULATE() {
  var katanaItems = fetchAllKatanaVariants();
  var katanaProducts = fetchAllKatanaProducts();
  var katanaMaterials = fetchAllKatanaMaterials();
  var waspData = readWaspRawTab();

  // Build product UOM map (productId → uom)
  var productUomMap = {};
  for (var p = 0; p < katanaProducts.length; p++) {
    var prod = katanaProducts[p];
    productUomMap[prod.id] = prod.unit_of_measure || prod.uom || '';
  }

  // Build material UOM map (materialId → uom)
  // Raw material variants have material_id — their UOM lives on the material record, not product
  var materialUomMap = {};
  for (var mat = 0; mat < katanaMaterials.length; mat++) {
    var material = katanaMaterials[mat];
    materialUomMap[material.id] = material.unit_of_measure || material.uom || '';
  }

  // Build conversion override map from KATANA_CONVERSION_ITEMS (sku → stockUom)
  // Supplier-level conversions (e.g. kg purchase → g stock) are NOT in the Katana API.
  // This hardcoded list is the source of truth for stock UOM on these items.
  var conversionMap = {};
  if (typeof KATANA_CONVERSION_ITEMS !== 'undefined') {
    for (var ci = 0; ci < KATANA_CONVERSION_ITEMS.length; ci++) {
      var convItem = KATANA_CONVERSION_ITEMS[ci];
      conversionMap[convItem.sku] = convItem.stockUom;
    }
  }

  // Build Katana SKU map with accurate stock UOM
  var katanaMap = {};
  for (var i = 0; i < katanaItems.length; i++) {
    var kItem = katanaItems[i];
    if (!kItem.sku) continue;
    var productId = kItem.product_id || '';
    var materialId = kItem.material_id || '';

    // Resolve stock UOM:
    // 1. KATANA_CONVERSION_ITEMS override (most accurate for purchase-conversion items)
    // 2. Materials endpoint (for raw materials)
    // 3. Products endpoint (for finished goods)
    var stockUom = '';
    if (conversionMap[kItem.sku]) {
      stockUom = conversionMap[kItem.sku];
    } else if (materialId && materialUomMap[materialId]) {
      stockUom = materialUomMap[materialId];
    } else if (productId && productUomMap[productId]) {
      stockUom = productUomMap[productId];
    }

    katanaMap[kItem.sku] = {
      name: kItem.name || kItem.sku,
      sku: kItem.sku,
      uom: stockUom,
      variantId: kItem.id,
      productId: productId,
      materialId: materialId
    };
  }

  // Merge all SKUs
  var allSkus = {};
  for (var sku1 in katanaMap) { allSkus[sku1] = true; }
  for (var sku2 in waspData) { allSkus[sku2] = true; }

  // Sort SKUs
  var sortedSkus = Object.keys(allSkus).sort();

  // Build comparison rows (14 columns — layout unchanged)
  var rows = [];
  for (var s = 0; s < sortedSkus.length; s++) {
    var sku = sortedSkus[s];
    var kData = katanaMap[sku] || null;
    var wData = waspData[sku] || null;

    var description = kData ? kData.name : (wData ? wData.description : '');
    var category = wData ? wData.category : '';
    var waspUom = wData ? wData.uom : '';
    var katanaUom = kData ? kData.uom : '';
    var uomMatch = '';
    if (waspUom && katanaUom) {
      uomMatch = normalizeUom(waspUom) === normalizeUom(katanaUom) ? 'YES' : 'MISMATCH';
    }
    var waspQty = wData ? wData.totalQty : 0;
    var katanaQty = ''; // Filled by itemCompare_ADD_KATANA_STOCK()
    var qtyDiff = '';
    var waspLocations = wData ? wData.locations : '';
    var inKatana = kData ? 'YES' : 'NO';
    var inWasp = wData ? 'YES' : 'NO';
    var waspLots = wData ? wData.lots : '';

    var status = '';
    if (!kData) status = 'WASP ONLY';
    else if (!wData) status = 'KATANA ONLY';
    else if (uomMatch === 'MISMATCH') status = 'UOM MISMATCH';
    else status = 'OK';

    rows.push([
      sku, description, category,
      waspUom, katanaUom, uomMatch,
      waspQty, katanaQty, qtyDiff,
      waspLocations, inKatana, inWasp, status,
      waspLots
    ]);
  }

  // Write to Comparison tab
  var ss = SpreadsheetApp.openById(COMPARISON_SHEET_ID);
  var compTab = ss.getSheetByName('Sheet1') || ss.getSheets()[0];

  var headers = [
    'SKU', 'Description', 'Category',
    'WASP UOM', 'Katana UOM', 'UOM Match?',
    'WASP Total Qty', 'Katana Qty', 'Qty Diff',
    'WASP Locations', 'In Katana?', 'In WASP?', 'Status',
    'WASP Lots'
  ];
  compTab.getRange(1, 1, 1, 14).setValues([headers]);

  // Clear old data (keep header)
  if (compTab.getLastRow() > 1) {
    compTab.getRange(2, 1, compTab.getLastRow() - 1, 14).clearContent();
  }

  // Write rows
  if (rows.length > 0) {
    compTab.getRange(2, 1, rows.length, 14).setValues(rows);
  }

  Logger.log('===== COMPARISON COMPLETE =====');
  Logger.log('Total SKUs: ' + rows.length);
  Logger.log('Written to: WASP-Katana Item Comparison sheet');
  Logger.log('');
  Logger.log('Next: Run itemCompare_ADD_KATANA_STOCK() to fetch Katana stock levels.');
}

/**
 * ADD KATANA STOCK — Fetches stock levels from Katana for each item
 * (Run separately because it makes many API calls)
 */
function itemCompare_ADD_KATANA_STOCK() {
  var ss = SpreadsheetApp.openById(COMPARISON_SHEET_ID);
  var compTab = ss.getSheetByName('Sheet1') || ss.getSheets()[0];
  var data = compTab.getDataRange().getValues();

  var updated = 0;
  var errors = 0;

  for (var i = 1; i < data.length; i++) {
    var sku = data[i][0];
    var inKatana = data[i][10]; // col 11 = In Katana?

    if (inKatana !== 'YES') continue;

    // Look up variant to get stock
    var variant = getKatanaVariantBySku(sku);
    if (!variant) {
      errors++;
      continue;
    }

    // Get stock level from variant or inventory
    var katanaStock = 0;
    if (variant.in_stock !== undefined) {
      katanaStock = variant.in_stock;
    } else if (variant.stock !== undefined) {
      katanaStock = variant.stock;
    }

    // Write Katana stock (column H = col 8)
    compTab.getRange(i + 1, 8).setValue(katanaStock);

    // Calculate diff (column I = col 9)
    var waspQty = data[i][6] || 0;
    var diff = Number(waspQty) - Number(katanaStock);
    compTab.getRange(i + 1, 9).setValue(diff);

    // Update status if qty doesn't match
    if (diff !== 0 && data[i][12] === 'OK') {
      compTab.getRange(i + 1, 13).setValue('QTY DIFF');
    }

    updated++;

    // Rate limit — Katana API allows ~5 req/sec
    if (updated % 5 === 0) {
      Utilities.sleep(1200);
    }
  }

  Logger.log('===== KATANA STOCK UPDATE COMPLETE =====');
  Logger.log('Updated: ' + updated + '  Errors: ' + errors);
}

// ============================================
// WASP EXPORT SHEET — Direct read for zero/sync
// ============================================
var WASP_EXPORT_SHEET_ID = '1j0dfwTzlmGoAu8axju88xsRSyA1qJwUp7tpbr3nsYso';

/**
 * Convert a date value (Date object, ISO string, or M/D/YYYY) to YYYY-MM-DD
 */
function formatDateCode(val) {
  if (!val) return '';
  // Date object from GAS getValues()
  if (val instanceof Date) {
    var y = val.getFullYear();
    var m = ('0' + (val.getMonth() + 1)).slice(-2);
    var d = ('0' + val.getDate()).slice(-2);
    return y + '-' + m + '-' + d;
  }
  var s = String(val);
  // ISO string: 2027-07-06T07:00:00.000Z
  if (s.indexOf('T') > 4 && s.indexOf('-') === 4) {
    return s.split('T')[0];
  }
  // M/D/YYYY: 7/6/2027
  var parts = s.split('/');
  if (parts.length === 3 && parts[2].length === 4) {
    return parts[2] + '-' + ('0' + parts[0]).slice(-2) + '-' + ('0' + parts[1]).slice(-2);
  }
  return s; // Return as-is
}

/**
 * ZERO FROM EXPORT — Reads directly from WASP export sheet.
 * Uses "Total In House" (col F) as qty — NOT "Total Available".
 * Only processes rows with Location + qty > 0.
 * Handles Date objects from Sheets correctly.
 *
 * Export columns: Item Number(A), Description(B), Total Available(C),
 *                 Lot(D), Date Code(E), Total In House(F), Location(G), Site(H)
 */
function zeroWasp_FROM_EXPORT() {
  var ss = SpreadsheetApp.openById(WASP_EXPORT_SHEET_ID);
  var sheet = ss.getSheets()[0]; // First tab
  var data = sheet.getDataRange().getValues();

  var results = { success: 0, failed: 0, skipped: 0, errors: [] };

  Logger.log('===== ZERO WASP FROM EXPORT =====');
  Logger.log('Rows to process: ' + (data.length - 1));

  for (var i = 1; i < data.length; i++) {
    var sku = String(data[i][0] || '').trim();
    var lot = data[i][3] ? String(data[i][3]).trim() : '';
    var dateCode = formatDateCode(data[i][4]);
    var qty = parseFloat(data[i][5]) || 0; // Total In House (col F)
    var location = String(data[i][6] || '').trim();
    var site = String(data[i][7] || '').trim();

    // Skip rows with no location, no qty, or no SKU
    if (!sku || qty <= 0 || !location) {
      results.skipped++;
      continue;
    }

    var result;
    if (lot) {
      result = waspRemoveInventoryWithLot(sku, qty, location, lot, 'Zero out — export sync', site, dateCode);
    } else {
      result = waspRemoveInventory(sku, qty, location, 'Zero out — export sync', site);
    }

    if (result.success) {
      results.success++;
      Logger.log('  OK: ' + sku + ' -' + qty + ' [' + (lot || 'no lot') + '] @ ' + site + '/' + location);
    } else {
      results.failed++;
      var errMsg = (result.response || '').substring(0, 120);
      results.errors.push(sku + ' [' + (lot || '-') + '] @ ' + location + ': ' + errMsg);
      Logger.log('  FAIL: ' + sku + ' [' + (lot || '-') + '] @ ' + location + ' — ' + errMsg);
    }

    // Rate limit
    if ((results.success + results.failed) % 10 === 0) {
      Utilities.sleep(500);
    }
  }

  Logger.log('');
  Logger.log('===== ZERO FROM EXPORT COMPLETE =====');
  Logger.log('Success: ' + results.success);
  Logger.log('Failed: ' + results.failed);
  Logger.log('Skipped (no location/qty): ' + results.skipped);

  if (results.errors.length > 0) {
    Logger.log('');
    Logger.log('ERRORS:');
    for (var e = 0; e < results.errors.length; e++) {
      Logger.log('  ' + results.errors[e]);
    }
  }
}

/**
 * PREVIEW — Shows what zeroWasp_FROM_EXPORT will remove
 */
function zeroWasp_EXPORT_PREVIEW() {
  var ss = SpreadsheetApp.openById(WASP_EXPORT_SHEET_ID);
  var sheet = ss.getSheets()[0];
  var data = sheet.getDataRange().getValues();

  var totalItems = 0;
  var totalQty = 0;
  var lotted = 0;
  var nonLotted = 0;

  Logger.log('===== ZERO PREVIEW (from WASP Export) =====');

  for (var i = 1; i < data.length; i++) {
    var sku = String(data[i][0] || '').trim();
    var lot = data[i][3] ? String(data[i][3]).trim() : '';
    var qty = parseFloat(data[i][5]) || 0;
    var location = String(data[i][6] || '').trim();
    var site = String(data[i][7] || '').trim();

    if (!sku || qty <= 0 || !location) continue;

    Logger.log('  ' + sku + ': -' + qty + ' [' + (lot || 'no lot') + '] @ ' + site + '/' + location);
    totalItems++;
    totalQty += qty;
    if (lot) { lotted++; } else { nonLotted++; }
  }

  Logger.log('');
  Logger.log('Total rows to remove: ' + totalItems);
  Logger.log('Total qty to remove: ' + totalQty);
  Logger.log('Lotted: ' + lotted + '  Non-lotted: ' + nonLotted);
  Logger.log('');
  Logger.log('Run zeroWasp_FROM_EXPORT() to execute.');
}

// ============================================
// ZERO OUT WASP — Set all items to qty 0
// ============================================

/**
 * PREVIEW — Shows what will be zeroed out
 */
function zeroWasp_PREVIEW() {
  var waspData = readWaspRawTab();
  var skus = Object.keys(waspData).sort();

  Logger.log('===== ITEMS TO ZERO OUT IN WASP =====');
  var totalItems = 0;
  var totalQty = 0;

  for (var i = 0; i < skus.length; i++) {
    var sku = skus[i];
    var item = waspData[sku];
    if (item.totalQty > 0) {
      Logger.log('  ' + sku + ': ' + item.totalQty + ' ' + item.uom + ' @ ' + item.locations);
      totalItems++;
      totalQty += item.totalQty;
    }
  }

  Logger.log('');
  Logger.log('Total items with stock: ' + totalItems);
  Logger.log('Total quantity to zero: ' + totalQty);
  Logger.log('');
  Logger.log('WARNING: This will remove ALL inventory from WASP!');
  Logger.log('Run zeroWasp_EXECUTE() to proceed.');
}

/**
 * EXECUTE — Removes all inventory from WASP
 * Reads the WASP Raw tab for per-location/lot data to remove correctly
 */
function zeroWasp_EXECUTE() {
  var ss = SpreadsheetApp.openById(COMPARISON_SHEET_ID);
  var rawTab = ss.getSheetByName('WASP Raw');
  if (!rawTab) {
    Logger.log('ERROR: "WASP Raw" tab not found. Populate it first.');
    return;
  }

  var rawData = rawTab.getDataRange().getValues();
  var results = { success: 0, failed: 0, skipped: 0, errors: [] };

  // Process each row (skip header)
  for (var i = 1; i < rawData.length; i++) {
    var sku = rawData[i][0];
    var qty = parseFloat(rawData[i][2]) || 0;
    var lot = rawData[i][3] ? String(rawData[i][3]) : '';
    var dateCode = rawData[i][4] || '';
    var location = rawData[i][5] || '';
    var site = rawData[i][6] || '';

    if (!sku || qty <= 0 || !location) {
      results.skipped++;
      continue;
    }

    // Format dateCode to YYYY-MM-DD if it's an ISO string
    if (dateCode && String(dateCode).indexOf('T') > -1) {
      dateCode = String(dateCode).split('T')[0];
    }

    var result;
    if (lot) {
      result = waspRemoveInventoryWithLot(sku, qty, location, lot, 'Zero out for item sync', site, dateCode);
    } else {
      result = waspRemoveInventory(sku, qty, location, 'Zero out for item sync', site);
    }

    if (result.success) {
      results.success++;
      Logger.log('  OK: ' + sku + ' -' + qty + ' [' + (lot || 'no lot') + '] @ ' + location);
    } else {
      results.failed++;
      results.errors.push(sku + ' @ ' + location + ': ' + (result.response || '').substring(0, 100));
      Logger.log('  FAIL: ' + sku + ' @ ' + location + ' — ' + (result.response || '').substring(0, 100));
    }

    // Rate limit
    if ((results.success + results.failed) % 10 === 0) {
      Utilities.sleep(500);
    }
  }

  Logger.log('');
  Logger.log('===== ZERO OUT COMPLETE =====');
  Logger.log('Success: ' + results.success);
  Logger.log('Failed: ' + results.failed);
  Logger.log('Skipped: ' + results.skipped);

  if (results.errors.length > 0) {
    Logger.log('');
    Logger.log('ERRORS:');
    for (var e = 0; e < results.errors.length; e++) {
      Logger.log('  ' + results.errors[e]);
    }
  }
}

/**
 * RETRY — Only processes rows that have lot/dateCode (the ones that failed last time)
 * Run this after deploying the DateCode fix to waspRemoveInventoryWithLot
 */
function zeroWasp_RETRY_LOTTED() {
  var ss = SpreadsheetApp.openById(COMPARISON_SHEET_ID);
  var rawTab = ss.getSheetByName('WASP Raw');
  if (!rawTab) {
    Logger.log('ERROR: "WASP Raw" tab not found.');
    return;
  }

  var rawData = rawTab.getDataRange().getValues();
  var results = { success: 0, failed: 0, skipped: 0, errors: [] };

  for (var i = 1; i < rawData.length; i++) {
    var sku = rawData[i][0];
    var qty = parseFloat(rawData[i][2]) || 0;
    var lot = rawData[i][3] ? String(rawData[i][3]) : '';
    var dateCode = rawData[i][4] || '';
    var location = rawData[i][5] || '';
    var site = rawData[i][6] || '';

    // Only process rows that HAVE a lot number (these are the ones that failed)
    if (!sku || qty <= 0 || !location || !lot) {
      results.skipped++;
      continue;
    }

    // Format dateCode to YYYY-MM-DD if it's an ISO string
    if (dateCode && String(dateCode).indexOf('T') > -1) {
      dateCode = String(dateCode).split('T')[0];
    }

    var result = waspRemoveInventoryWithLot(sku, qty, location, lot, 'Zero out for item sync', site, dateCode);

    if (result.success) {
      results.success++;
      Logger.log('  OK: ' + sku + ' -' + qty + ' [' + lot + '] dc=' + dateCode + ' @ ' + location);
    } else {
      results.failed++;
      results.errors.push(sku + ' @ ' + location + ' lot=' + lot + ' dc=' + dateCode + ': ' + (result.response || '').substring(0, 120));
      Logger.log('  FAIL: ' + sku + ' @ ' + location + ' lot=' + lot + ' dc=' + dateCode + ' — ' + (result.response || '').substring(0, 120));
    }

    if ((results.success + results.failed) % 10 === 0) {
      Utilities.sleep(500);
    }
  }

  Logger.log('');
  Logger.log('===== RETRY LOTTED COMPLETE =====');
  Logger.log('Success: ' + results.success);
  Logger.log('Failed: ' + results.failed);
  Logger.log('Skipped (no lot): ' + results.skipped);

  if (results.errors.length > 0) {
    Logger.log('');
    Logger.log('ERRORS:');
    for (var e = 0; e < results.errors.length; e++) {
      Logger.log('  ' + results.errors[e]);
    }
  }
}

// ============================================
// HELPER FUNCTIONS
// ============================================

/**
 * Fetch ALL Katana variants (paginated)
 * Katana API may cap per_page at 50 — use dynamic break based on actual page size
 */
function fetchAllKatanaVariants() {
  var allVariants = [];
  var page = 1;
  var hasMore = true;
  var perPage = 200; // Request max; API may return fewer

  while (hasMore) {
    var result = katanaApiCall('variants?per_page=' + perPage + '&page=' + page);
    if (!result) {
      Logger.log('  ERROR: null response on page ' + page);
      hasMore = false;
      break;
    }

    // Handle both { data: [...] } and direct array response
    var items = result.data || result;
    if (!items || !items.length || items.length === 0) {
      hasMore = false;
      break;
    }

    for (var i = 0; i < items.length; i++) {
      allVariants.push(items[i]);
    }

    Logger.log('  Fetched variants page ' + page + ' (' + items.length + ' items)');

    // Stop if we got fewer than requested — no more pages
    // Use actual count from first page as the page size reference
    if (page === 1) {
      perPage = items.length; // API's actual page size
    }
    if (items.length < perPage) {
      hasMore = false;
    } else {
      page++;
      Utilities.sleep(300); // Rate limit
    }
  }

  Logger.log('Total Katana variants: ' + allVariants.length);
  return allVariants;
}

/**
 * Fetch ALL Katana products (paginated) — for UOM data
 */
function fetchAllKatanaProducts() {
  var allProducts = [];
  var page = 1;
  var hasMore = true;
  var perPage = 200;

  while (hasMore) {
    var result = katanaApiCall('products?per_page=' + perPage + '&page=' + page);
    if (!result) { hasMore = false; break; }

    var items = result.data || result;
    if (!items || !items.length || items.length === 0) {
      hasMore = false;
      break;
    }

    for (var i = 0; i < items.length; i++) {
      allProducts.push(items[i]);
    }

    if (page === 1) { perPage = items.length; }
    if (items.length < perPage) {
      hasMore = false;
    } else {
      page++;
      Utilities.sleep(300);
    }
  }

  Logger.log('Total Katana products: ' + allProducts.length);
  return allProducts;
}

/**
 * Fetch ALL Katana materials (paginated) — for raw material UOM data.
 * Raw material variants carry a material_id; their UOM lives on the material record.
 */
function fetchAllKatanaMaterials() {
  var allMaterials = [];
  var page = 1;
  var hasMore = true;
  var perPage = 200;

  while (hasMore) {
    var result = katanaApiCall('materials?per_page=' + perPage + '&page=' + page);
    if (!result) { hasMore = false; break; }

    var items = result.data || result;
    if (!items || !items.length || items.length === 0) {
      hasMore = false;
      break;
    }

    for (var i = 0; i < items.length; i++) {
      allMaterials.push(items[i]);
    }

    if (page === 1) { perPage = items.length; }
    if (items.length < perPage) {
      hasMore = false;
    } else {
      page++;
      Utilities.sleep(300);
    }
  }

  Logger.log('Total Katana materials: ' + allMaterials.length);
  return allMaterials;
}

/**
 * Read WASP Raw tab and aggregate by SKU
 * Returns { sku: { description, totalQty, uom, category, locations, lots } }
 */
function readWaspRawTab() {
  var ss = SpreadsheetApp.openById(COMPARISON_SHEET_ID);
  var rawTab = ss.getSheetByName('WASP Raw');
  if (!rawTab) {
    Logger.log('WARNING: WASP Raw tab not found');
    return {};
  }

  var data = rawTab.getDataRange().getValues();
  var items = {};

  // Columns: SKU(0), Description(1), Qty(2), Lot(3), DateCode(4), Location(5), Site(6), UOM(7), Category(8)
  for (var i = 1; i < data.length; i++) {
    var sku = String(data[i][0] || '').trim();
    if (!sku) continue;

    var qty = parseFloat(data[i][2]) || 0;
    var lot = String(data[i][3] || '').trim();
    var location = String(data[i][5] || '').trim();
    var uom = String(data[i][7] || '').trim();
    var category = String(data[i][8] || '').trim();
    var description = String(data[i][1] || '').trim();

    if (!items[sku]) {
      items[sku] = {
        description: description,
        totalQty: 0,
        uom: uom,
        category: category,
        locationSet: {},
        lotSet: {}
      };
    }

    items[sku].totalQty += qty;
    if (location) items[sku].locationSet[location] = true;
    if (lot) items[sku].lotSet[lot] = true;
  }

  // Convert sets to strings
  for (var s in items) {
    items[s].locations = Object.keys(items[s].locationSet).join(', ');
    items[s].lots = Object.keys(items[s].lotSet).join(', ');
    delete items[s].locationSet;
    delete items[s].lotSet;
  }

  return items;
}

/**
 * Normalize UOM for comparison
 */
function normalizeUom(uom) {
  if (!uom) return '';
  var u = uom.toLowerCase().trim();

  // Common normalizations
  if (u === 'each' || u === 'ea' || u === 'pcs' || u === 'pc') return 'each';
  if (u === 'grams' || u === 'gram' || u === 'g' || u === 'gr') return 'grams';
  if (u === 'kg' || u === 'kilogram' || u === 'kilograms') return 'kg';
  if (u === 'box' || u === 'bx') return 'box';
  if (u === 'rolls' || u === 'roll') return 'rolls';
  if (u === 'pk' || u === 'pack' || u === 'packs') return 'pk';

  return u;
}

// ============================================
// STEP 4: RE-ADD FROM KATANA (Per-Location)
// ============================================
// Fetches Katana inventory per location + item type,
// builds a plan, then executes against WASP.
// Location mapping:
//   MMH Kelowna + product → SHIPPING-DOCK
//   MMH Kelowna + material/unknown → RECEIVING-DOCK
//   Shopify + any → SHOPIFY (new WASP location)
//   Storage Warehouse → SW-STORAGE
//   Amazon USA → SKIP (not physically at our facility)
// ============================================

/**
 * Fetch ALL Katana inventory records (per variant, per location)
 * Uses GET /v1/inventory — returns { variant_id, location_id, quantity_in_stock, ... }
 */
function fetchAllKatanaInventory() {
  var allRecords = [];
  var page = 1;
  var hasMore = true;
  var limit = 200;

  while (hasMore) {
    var result = katanaApiCall('inventory?limit=' + limit + '&page=' + page);
    if (!result) { hasMore = false; break; }

    var items = result.data || result;
    if (!items || !items.length || items.length === 0) {
      hasMore = false;
      break;
    }

    for (var i = 0; i < items.length; i++) {
      allRecords.push(items[i]);
    }

    Logger.log('  Fetched inventory page ' + page + ' (' + items.length + ' records)');

    if (page === 1) { limit = items.length; }
    if (items.length < limit) {
      hasMore = false;
    } else {
      page++;
      Utilities.sleep(300);
    }
  }

  Logger.log('Total Katana inventory records: ' + allRecords.length);
  return allRecords;
}

/**
 * Fetch all Katana locations
 * Returns { locationId: locationName }
 */
function fetchAllKatanaLocations() {
  var result = katanaApiCall('locations?limit=100');
  if (!result) return {};

  var items = result.data || result;
  if (!items || !items.length) return {};

  var map = {};
  for (var i = 0; i < items.length; i++) {
    var loc = items[i];
    map[loc.id] = loc.name || ('Location-' + loc.id);
  }

  Logger.log('Katana locations found: ' + Object.keys(map).length);
  return map;
}

/**
 * BUILD RE-ADD PLAN — Fetches Katana inventory per location,
 * maps each item to the correct WASP site/location,
 * and writes the plan to the "Re-Add Plan" tab.
 *
 * Run this first. Review the plan, then run reAddPlan_ADD_BATCHES()
 * and finally reAddPlan_EXECUTE().
 */
function reAddPlan_BUILD() {
  Logger.log('===== BUILDING RE-ADD PLAN =====');

  // 1. Fetch all data from Katana
  Logger.log('Fetching Katana inventory (per location)...');
  var inventory = fetchAllKatanaInventory();

  Logger.log('Fetching Katana locations...');
  var locationMap = fetchAllKatanaLocations();

  Logger.log('Fetching Katana variants...');
  var variants = fetchAllKatanaVariants();

  Logger.log('Fetching Katana products...');
  var products = fetchAllKatanaProducts();

  // 2. Build lookup maps
  var variantMap = {};
  for (var v = 0; v < variants.length; v++) {
    var vr = variants[v];
    variantMap[vr.id] = {
      sku: vr.sku || '',
      name: vr.name || '',
      productId: vr.product_id || ''
    };
  }

  var productMap = {};
  for (var p = 0; p < products.length; p++) {
    var prod = products[p];
    productMap[prod.id] = {
      type: prod.type || prod.product_type || prod.category || 'unknown',
      uom: prod.unit_of_measure || prod.uom || '',
      name: prod.name || ''
    };
  }

  // 3. Build plan rows
  var planRows = [];
  var skipped = 0;

  for (var i = 0; i < inventory.length; i++) {
    var inv = inventory[i];
    var qty = inv.quantity_in_stock || 0;
    if (qty <= 0) { skipped++; continue; }

    var variantId = inv.variant_id;
    var locationId = inv.location_id;

    var vInfo = variantMap[variantId];
    if (!vInfo || !vInfo.sku) { skipped++; continue; }

    // Skip OP- items
    if (isSkippedSku(vInfo.sku)) { skipped++; continue; }

    var katanaLocName = locationMap[locationId] || 'Unknown';
    var pInfo = productMap[vInfo.productId] || { type: 'unknown', uom: '' };

    // Determine WASP destination based on Katana location + item type
    // Skip Amazon USA — stock is at Amazon's fulfillment center, not our facility
    if (katanaLocName === 'Amazon USA') { skipped++; continue; }

    var waspSite = 'MMH Kelowna';
    var waspLocation = 'SHIPPING-DOCK';

    if (katanaLocName === 'Storage Warehouse') {
      waspSite = 'Storage Warehouse';
      waspLocation = 'SW-STORAGE';
    } else if (katanaLocName === 'Shopify') {
      waspLocation = 'SHOPIFY';
    } else if (pInfo.type === 'material' || pInfo.type === 'unknown') {
      waspLocation = 'RECEIVING-DOCK';
    }

    planRows.push([
      vInfo.sku,
      vInfo.name,
      qty,
      katanaLocName,
      pInfo.type,
      pInfo.uom,
      waspSite,
      waspLocation,
      '', // Lot — filled by reAddPlan_ADD_BATCHES
      '', // DateCode — filled by reAddPlan_ADD_BATCHES
      'PENDING'
    ]);
  }

  // 4. Write to Re-Add Plan tab
  var ss = SpreadsheetApp.openById(COMPARISON_SHEET_ID);
  var planTab = ss.getSheetByName('Re-Add Plan');
  if (!planTab) {
    planTab = ss.insertSheet('Re-Add Plan');
  }

  var headers = [
    'SKU', 'Name', 'Qty', 'Katana Location', 'Item Type',
    'Katana UOM', 'WASP Site', 'WASP Location',
    'Lot', 'DateCode', 'Status'
  ];

  planTab.clear();
  planTab.getRange(1, 1, 1, headers.length).setValues([headers]);

  // Sort by SKU
  planRows.sort(function(a, b) { return a[0] < b[0] ? -1 : a[0] > b[0] ? 1 : 0; });

  if (planRows.length > 0) {
    planTab.getRange(2, 1, planRows.length, headers.length).setValues(planRows);
  }

  Logger.log('');
  Logger.log('===== RE-ADD PLAN COMPLETE =====');
  Logger.log('Plan rows: ' + planRows.length);
  Logger.log('Skipped (zero qty / no SKU / OP-): ' + skipped);
  Logger.log('Locations found: ' + JSON.stringify(locationMap));
  Logger.log('');
  Logger.log('NEXT STEPS:');
  Logger.log('1. Review the Re-Add Plan tab');
  Logger.log('2. Run reAddPlan_ADD_BATCHES() for lot/expiry data');
  Logger.log('3. Zero WASP first: zeroWasp_EXECUTE()');
  Logger.log('4. Execute: reAddPlan_EXECUTE()');
}

/**
 * ADD BATCHES — Fetches Katana batch_stocks for each variant,
 * updates Re-Add Plan rows with Lot and DateCode.
 * For multi-lot items, splits into separate rows.
 * Run separately — makes many API calls (~5/sec rate limit).
 */
function reAddPlan_ADD_BATCHES() {
  var ss = SpreadsheetApp.openById(COMPARISON_SHEET_ID);
  var planTab = ss.getSheetByName('Re-Add Plan');
  if (!planTab) {
    Logger.log('ERROR: Re-Add Plan tab not found. Run reAddPlan_BUILD() first.');
    return;
  }

  var data = planTab.getDataRange().getValues();
  var updated = 0;
  var errors = 0;

  // Build SKU → variantId map
  var variants = fetchAllKatanaVariants();
  var skuToVariantId = {};
  for (var v = 0; v < variants.length; v++) {
    if (variants[v].sku) {
      skuToVariantId[variants[v].sku] = variants[v].id;
    }
  }

  // Track which SKUs we already fetched batches for (avoid duplicate calls)
  var fetchedSkus = {};
  var extraRows = [];

  for (var i = 1; i < data.length; i++) {
    var sku = data[i][0];
    var planQty = Number(data[i][2]) || 0;
    if (!sku || planQty <= 0) continue;

    // Skip if already has lot data
    if (data[i][8]) continue;

    var variantId = skuToVariantId[sku];
    if (!variantId) continue;

    // Skip if we already fetched for this SKU (multi-location)
    if (fetchedSkus[sku]) continue;
    fetchedSkus[sku] = true;

    // Fetch batch stocks
    var batchResult = katanaApiCall('batch_stocks?variant_id=' + variantId);
    if (!batchResult) { errors++; continue; }

    var batches = batchResult.data || batchResult || [];
    if (!batches || !batches.length || batches.length === 0) continue;

    // Filter batches with stock > 0
    var activeBatches = [];
    for (var b = 0; b < batches.length; b++) {
      var batch = batches[b];
      var batchQty = batch.in_stock || batch.quantity || batch.quantity_in_stock || 0;
      if (batchQty > 0) {
        var dc = batch.best_before || batch.expiry_date || batch.date_code || '';
        if (dc && String(dc).indexOf('T') > -1) dc = String(dc).split('T')[0];
        activeBatches.push({
          lot: batch.batch_number || batch.nr || batch.number || '',
          qty: batchQty,
          dateCode: dc
        });
      }
    }

    if (activeBatches.length === 0) continue;

    if (activeBatches.length === 1) {
      // Single lot — update existing row
      planTab.getRange(i + 1, 9).setValue(activeBatches[0].lot);
      planTab.getRange(i + 1, 10).setValue(activeBatches[0].dateCode);
      updated++;
    } else {
      // Multiple lots — update first row with first batch, append rest
      planTab.getRange(i + 1, 3).setValue(activeBatches[0].qty);
      planTab.getRange(i + 1, 9).setValue(activeBatches[0].lot);
      planTab.getRange(i + 1, 10).setValue(activeBatches[0].dateCode);

      for (var nb = 1; nb < activeBatches.length; nb++) {
        var ab = activeBatches[nb];
        extraRows.push([
          sku, data[i][1], ab.qty, data[i][3], data[i][4],
          data[i][5], data[i][6], data[i][7],
          ab.lot, ab.dateCode, 'PENDING'
        ]);
      }
      updated++;
    }

    // Rate limit — Katana API ~5 req/sec
    if ((updated + errors) % 5 === 0) {
      Utilities.sleep(1200);
    }
  }

  // Append extra rows for multi-lot items
  for (var r = 0; r < extraRows.length; r++) {
    planTab.appendRow(extraRows[r]);
  }

  Logger.log('');
  Logger.log('===== BATCH UPDATE COMPLETE =====');
  Logger.log('SKUs updated with lot data: ' + updated);
  Logger.log('Extra rows added (multi-lot): ' + extraRows.length);
  Logger.log('Errors: ' + errors);
}

/**
 * EXECUTE RE-ADD PLAN — Reads the Re-Add Plan tab and adds
 * inventory to WASP. Run AFTER zeroing WASP with zeroWasp_EXECUTE().
 *
 * For items with Lot/DateCode: uses waspAddInventoryWithLot()
 * For items without: uses waspAddInventory()
 */
function reAddPlan_EXECUTE() {
  var ss = SpreadsheetApp.openById(COMPARISON_SHEET_ID);
  var planTab = ss.getSheetByName('Re-Add Plan');
  if (!planTab) {
    Logger.log('ERROR: Re-Add Plan tab not found.');
    return;
  }

  var data = planTab.getDataRange().getValues();
  var results = { success: 0, failed: 0, skipped: 0, errors: [] };

  for (var i = 1; i < data.length; i++) {
    var sku = data[i][0];
    var qty = Number(data[i][2]) || 0;
    var waspSite = data[i][6];
    var waspLocation = data[i][7];
    var lot = String(data[i][8] || '').trim();
    var dateCode = formatDateCode(data[i][9]);
    var status = String(data[i][10] || '');

    if (!sku || qty <= 0 || status === 'DONE') {
      results.skipped++;
      continue;
    }

    var result;
    if (lot) {
      result = waspAddInventoryWithLot(sku, qty, waspLocation, lot, dateCode, 'Katana sync re-add', waspSite);
    } else {
      result = waspAddInventory(sku, qty, waspLocation, 'Katana sync re-add', waspSite);
    }

    if (result.success) {
      results.success++;
      planTab.getRange(i + 1, 11).setValue('DONE');
      Logger.log('  OK: ' + sku + ' +' + qty + ' [' + (lot || 'no lot') + '] @ ' + waspSite + '/' + waspLocation);
    } else {
      results.failed++;
      var errMsg = (result.response || '').substring(0, 100);
      planTab.getRange(i + 1, 11).setValue('ERROR: ' + errMsg);
      results.errors.push(sku + ': ' + errMsg);
      Logger.log('  FAIL: ' + sku + ' — ' + errMsg);
    }

    // Rate limit
    if ((results.success + results.failed) % 10 === 0) {
      Utilities.sleep(500);
    }
  }

  Logger.log('');
  Logger.log('===== RE-ADD COMPLETE =====');
  Logger.log('Success: ' + results.success);
  Logger.log('Failed: ' + results.failed);
  Logger.log('Skipped: ' + results.skipped);

  if (results.errors.length > 0) {
    Logger.log('');
    Logger.log('ERRORS:');
    for (var e = 0; e < results.errors.length; e++) {
      Logger.log('  ' + results.errors[e]);
    }
  }
}

/**
 * RETRY FAILED — Re-processes only rows with ERROR status
 */
function reAddPlan_RETRY_FAILED() {
  var ss = SpreadsheetApp.openById(COMPARISON_SHEET_ID);
  var planTab = ss.getSheetByName('Re-Add Plan');
  if (!planTab) {
    Logger.log('ERROR: Re-Add Plan tab not found.');
    return;
  }

  var data = planTab.getDataRange().getValues();
  var results = { success: 0, failed: 0, skipped: 0, errors: [] };

  for (var i = 1; i < data.length; i++) {
    var sku = data[i][0];
    var qty = Number(data[i][2]) || 0;
    var waspSite = data[i][6];
    var waspLocation = data[i][7];
    var lot = String(data[i][8] || '').trim();
    var dateCode = formatDateCode(data[i][9]);
    var status = String(data[i][10] || '');

    // Only retry ERROR rows
    if (!sku || qty <= 0 || status === 'DONE' || status === 'PENDING' || status.indexOf('ERROR') < 0) {
      results.skipped++;
      continue;
    }

    var result;
    if (lot) {
      result = waspAddInventoryWithLot(sku, qty, waspLocation, lot, dateCode, 'Katana sync retry', waspSite);
    } else {
      result = waspAddInventory(sku, qty, waspLocation, 'Katana sync retry', waspSite);
    }

    if (result.success) {
      results.success++;
      planTab.getRange(i + 1, 11).setValue('DONE');
    } else {
      results.failed++;
      var errMsg = (result.response || '').substring(0, 100);
      planTab.getRange(i + 1, 11).setValue('ERROR: ' + errMsg);
      results.errors.push(sku + ': ' + errMsg);
    }

    if ((results.success + results.failed) % 10 === 0) {
      Utilities.sleep(500);
    }
  }

  Logger.log('===== RETRY COMPLETE =====');
  Logger.log('Success: ' + results.success + '  Failed: ' + results.failed + '  Skipped: ' + results.skipped);
}
