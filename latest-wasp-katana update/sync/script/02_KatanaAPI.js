/**
 * Katana MRP API - All Katana API calls
 */

var KATANA_BASE = 'https://api.katanamrp.com/v1';
var TRACKED_LOCATIONS = ['MMH Kelowna', 'MMH Mayfair', 'Storage Warehouse'];
var KATANA_MASTER_CACHE_TTL_MS = 6 * 60 * 60 * 1000;

// Fallback conversion rates verified manually in Katana (Supply details tab).
// Used when the API does not return unit_conversion_rate for a material/product.
// Format: SKU → number of stock units per 1 purchase unit.
var UOM_RATE_FALLBACK = {
  // Essential oils & wax: kg → g
  'B-PROP':      1000,  // 1 kg = 1000 g
  'B-WAX':       1000,  // 1 kg = 1000 g
  'EO-CB':       1000,  // 1 kg = 1000 g
  'EO-CS':       1000,  // 1 kg = 1000 g
  'EO-LAV':      1000,  // 1 kg = 1000 g
  'EO-PEP':      1000,  // 1 kg = 1000 g
  'EO-THY':      1000,  // 1 kg = 1000 g
  'EO-TT':       1000,  // 1 kg = 1000 g
  // Herbs: lbs → g
  'H-ARNF':      454,   // 1 lbs = 454 g
  'H-COML':      454,   // 1 lbs = 454 g
  // Eggs: dozen → pcs
  'EGG-X':       12,    // 1 dozen = 12 pcs
  // Labels (confirmed from Katana supply details)
  'FBA-BARCODE': 500,   // 1 pack = 500 PC
  'FBA-FRAGILE': 500,   // 1 pack = 500 PC
  // Non-inventory consumables (confirmed via browser agent)
  'NI-ATPSWABS':    100,  // 1 BOX = 100 EA
  'NI-NAPKINS':     12,   // 1 box = 12 pack
  'NI-GLOVEMD':     200,  // 1 pack = 200 EA
  'NI-GLOVESM':     200,  // 1 pack = 200 EA
  'NI-GLOVELR':     200,  // 1 pack = 200 EA
  'NI-GLOVECR':     200,  // 1 pack = 200 EA
  'NI-STEELWOOL':   16,   // 1 pack = 16 PC
  'NI-PAPERTOWEL8': 8,    // 1 pcs = 8 ROLLS
  'NI-SCRUB-10':    10,   // 1 pack = 10 pc
  'NI-FIRE-BNKT':   4,    // 1 pack = 4 pc
  'INFBAG-160':     100   // 1 pack = 100 PC
};

function loadKatanaCacheAnyAge_(key) {
  var props = PropertiesService.getScriptProperties();
  var count = parseInt(props.getProperty(key + '_COUNT') || '0', 10);
  if (count === 0) return null;

  var json = '';
  for (var i = 0; i < count; i++) {
    json += (props.getProperty(key + '_' + i) || '');
  }

  try {
    return {
      savedAt: parseInt(props.getProperty(key + '_TS') || '0', 10),
      data: JSON.parse(json)
    };
  } catch (e) {
    Logger.log('Failed to parse Katana cache ' + key + ': ' + e.message);
    return null;
  }
}

function loadKatanaCacheFresh_(key, ttlMs) {
  var cached = loadKatanaCacheAnyAge_(key);
  if (!cached || !cached.savedAt) return null;
  if ((new Date()).getTime() - cached.savedAt > ttlMs) return null;
  return cached.data;
}

function saveKatanaCache_(key, data) {
  var props = PropertiesService.getScriptProperties();
  var oldCount = parseInt(props.getProperty(key + '_COUNT') || '0', 10);
  for (var i = 0; i < oldCount; i++) {
    props.deleteProperty(key + '_' + i);
  }

  var json = JSON.stringify(data);
  var chunkSize = 8000;
  var chunks = Math.ceil(json.length / chunkSize);
  for (var c = 0; c < chunks; c++) {
    props.setProperty(key + '_' + c, json.substring(c * chunkSize, (c + 1) * chunkSize));
  }
  props.setProperty(key + '_COUNT', String(chunks));
  props.setProperty(key + '_TS', String((new Date()).getTime()));
  Logger.log('Saved Katana cache ' + key + ' (' + chunks + ' chunks)');
}

function fetchKatanaCached_(key, ttlMs, fetchFn) {
  var fresh = loadKatanaCacheFresh_(key, ttlMs);
  if (fresh !== null) {
    Logger.log('Using fresh Katana cache: ' + key);
    return fresh;
  }

  try {
    var data = fetchFn();
    saveKatanaCache_(key, data);
    return data;
  } catch (e) {
    var stale = loadKatanaCacheAnyAge_(key);
    if (stale && stale.data !== null) {
      Logger.log('Using stale Katana cache for ' + key + ' after fetch failure: ' + e.message);
      return stale.data;
    }
    throw e;
  }
}

function katanaFetch(endpoint, params) {
  var token = PropertiesService.getScriptProperties().getProperty('KATANA_API_KEY');
  if (!token) throw new Error('Katana API key not set.');

  var url = KATANA_BASE + endpoint;
  if (params) {
    var parts = [];
    for (var key in params) {
      if (params.hasOwnProperty(key)) {
        parts.push(encodeURIComponent(key) + '=' + encodeURIComponent(params[key]));
      }
    }
    if (parts.length) url += '?' + parts.join('&');
  }

  var options = {
    method: 'get',
    headers: { 'Authorization': 'Bearer ' + token, 'Accept': 'application/json' },
    muteHttpExceptions: true
  };

  var response = UrlFetchApp.fetch(url, options);
  var code = response.getResponseCode();
  if (code === 429) {
    var wait = parseInt(response.getHeaders()['retry-after'] || '5');
    Utilities.sleep(wait * 1000);
    response = UrlFetchApp.fetch(url, options);
    code = response.getResponseCode();
  }
  if (code !== 200 && code !== 201) {
    throw new Error('Katana ' + code + ' on ' + endpoint + ': ' + response.getContentText().substring(0, 300));
  }
  return JSON.parse(response.getContentText());
}

function katanaFetchAllPages(endpoint) {
  var allData = [];
  var page = 1;
  var totalPages = 999;
  var token = PropertiesService.getScriptProperties().getProperty('KATANA_API_KEY');

  while (page <= totalPages) {
    var url = KATANA_BASE + endpoint + (endpoint.indexOf('?') > -1 ? '&' : '?') + 'limit=250&page=' + page;
    var response = UrlFetchApp.fetch(url, {
      method: 'get',
      headers: { 'Authorization': 'Bearer ' + token, 'Accept': 'application/json' },
      muteHttpExceptions: true
    });
    var code = response.getResponseCode();
    if (code === 429) { Utilities.sleep(5000); continue; }
    if (code !== 200) throw new Error('Katana ' + code + ' on ' + endpoint);

    var body = JSON.parse(response.getContentText());
    var items = Array.isArray(body) ? body : (body.data || []);
    for (var i = 0; i < items.length; i++) { allData.push(items[i]); }

    var pHeader = response.getHeaders()['x-pagination'];
    if (pHeader) {
      var p = JSON.parse(pHeader);
      totalPages = p.total_pages || 1;
    } else {
      totalPages = items.length === 250 ? page + 1 : page;
    }
    page++;
    if (page > 50) break;
  }
  return allData;
}

function fetchKatanaBatchStocks(variantId) {
  try {
    var result = katanaFetch('/batch_stocks', { variant_id: variantId });
    var batches = result.data || result || [];
    if (!Array.isArray(batches)) return [];
    return batches;
  } catch (e) {
    Logger.log('batch_stocks error for variant ' + variantId + ': ' + e.message);
    return [];
  }
}

function capitalizeType(type) {
  if (!type) return '';
  var t = String(type).toLowerCase();
  if (t === 'product') return 'Product';
  if (t === 'material') return 'Material';
  if (t === 'intermediate') return 'Intermediate';
  return t.charAt(0).toUpperCase() + t.slice(1);
}

function fetchAllKatanaData() {
  var locations = fetchKatanaCached_('KATANA_LOCATIONS', KATANA_MASTER_CACHE_TTL_MS, function() {
    var rows = katanaFetch('/locations');
    if (rows.data) rows = rows.data;
    if (!Array.isArray(rows)) rows = [];
    return rows;
  });

  var products = fetchKatanaCached_('KATANA_PRODUCTS', KATANA_MASTER_CACHE_TTL_MS, function() {
    return katanaFetchAllPages('/products');
  });
  var materials = fetchKatanaCached_('KATANA_MATERIALS', KATANA_MASTER_CACHE_TTL_MS, function() {
    return katanaFetchAllPages('/materials');
  });
  var variants = fetchKatanaCached_('KATANA_VARIANTS', KATANA_MASTER_CACHE_TTL_MS, function() {
    return katanaFetchAllPages('/variants');
  });
  var inventory = katanaFetchAllPages('/inventory');

  var locationMap = {};
  for (var li = 0; li < locations.length; li++) {
    locationMap[locations[li].id] = locations[li].name;
  }

  var productMap = {};
  for (var pi = 0; pi < products.length; pi++) {
    var prod = products[pi];
    var pBt = prod.batch_tracking || prod.batch_tracked || prod.is_batch_tracked || prod.track_batches || false;
    productMap[prod.id] = { name: prod.name, type: prod.type || prod.product_type || 'product', uom: prod.unit_of_measure || prod.uom || 'ea', purchaseUom: prod.purchase_unit_of_measure || prod.purchase_uom || '', conversionRate: prod.unit_conversion_rate || prod.conversion_rate || null, batchTracking: pBt ? true : false, category: prod.category || prod.category_name || '' };
  }

  var materialMap = {};
  for (var mi = 0; mi < materials.length; mi++) {
    var mat = materials[mi];
    var mBt = mat.batch_tracking || mat.batch_tracked || mat.is_batch_tracked || mat.track_batches || false;
    // Log first material's fields once to identify the correct conversion rate field name
    if (mi === 0) Logger.log('Material fields: ' + Object.keys(mat).join(', '));
    var mRate = mat.unit_conversion_rate || mat.conversion_rate || mat.purchase_conversion_rate || mat.uom_conversion_rate || mat.conversion || null;
    materialMap[mat.id] = { name: mat.name, uom: mat.unit_of_measure || mat.uom || 'ea', purchaseUom: mat.purchase_unit_of_measure || mat.purchase_uom || '', conversionRate: mRate, batchTracking: mBt ? true : false, category: mat.category || mat.category_name || '' };
  }

  var variantMap = {};
  var batchTrackingSkus = {};
  for (var vi = 0; vi < variants.length; vi++) {
    var v = variants[vi];
    var name = '', type = '', uom = 'ea', purchaseUom = '', conversionRate = null, bt = false, cat = '';
    if (v.product_id) {
      var pInfo = productMap[v.product_id];
      name = pInfo ? pInfo.name : 'Unknown Product';
      type = pInfo ? pInfo.type : 'product';
      uom = pInfo ? pInfo.uom : 'ea';
      purchaseUom = pInfo ? (pInfo.purchaseUom || '') : '';
      conversionRate = pInfo ? (pInfo.conversionRate || null) : null;
      bt = pInfo ? pInfo.batchTracking : false;
      cat = pInfo ? pInfo.category : '';
    } else if (v.material_id) {
      var mInfo = materialMap[v.material_id];
      name = mInfo ? mInfo.name : 'Unknown Material';
      type = 'material';
      uom = mInfo ? mInfo.uom : 'ea';
      purchaseUom = mInfo ? (mInfo.purchaseUom || '') : '';
      conversionRate = mInfo ? (mInfo.conversionRate || null) : null;
      bt = mInfo ? mInfo.batchTracking : false;
      cat = mInfo ? mInfo.category : '';
    }
    variantMap[v.id] = { sku: v.sku, name: name, type: type, uom: uom, purchaseUom: purchaseUom, conversionRate: conversionRate, category: cat, salesPrice: v.sales_price || 0, productId: v.product_id || '', materialId: v.material_id || '' };
    if (bt && v.sku) batchTrackingSkus[v.sku] = true;
  }

  // Log first inventory record fields for debugging stock issues
  if (inventory.length > 0) {
    Logger.log('Inventory fields: ' + Object.keys(inventory[0]).join(', '));
    Logger.log('Sample inventory: ' + JSON.stringify(inventory[0]).substring(0, 500));
  }

  var items = [];
  for (var ii = 0; ii < inventory.length; ii++) {
    var inv = inventory[ii];
    var locName = locationMap[inv.location_id] || 'Unknown';
    var isTracked = false;
    for (var tl = 0; tl < TRACKED_LOCATIONS.length; tl++) {
      if (locName === TRACKED_LOCATIONS[tl]) { isTracked = true; break; }
    }
    if (!isTracked) continue;

    var vInfo = variantMap[inv.variant_id] || { sku: 'UNKNOWN', name: 'Unknown', type: '?', uom: 'ea', purchaseUom: '', conversionRate: null, category: '', salesPrice: 0 };
    var rawQty = parseFloat(inv.quantity_in_stock) || 0;
    if (rawQty === undefined || rawQty === null) rawQty = inv.in_stock;
    if (rawQty === undefined || rawQty === null) rawQty = inv.stock;
    if (rawQty === undefined || rawQty === null) rawQty = inv.on_hand;
    if (rawQty === undefined || rawQty === null) rawQty = 0;

    items.push({
      variantId: inv.variant_id, sku: vInfo.sku, name: vInfo.name,
      type: capitalizeType(vInfo.type), uom: vInfo.uom, purchaseUom: vInfo.purchaseUom || '', conversionRate: vInfo.conversionRate || UOM_RATE_FALLBACK[vInfo.sku] || null, category: vInfo.category || '',
      locationId: inv.location_id, locationName: locName,
      qtyInStock: rawQty, qtyCommitted: inv.quantity_committed || inv.committed || 0,
      qtyExpected: inv.quantity_expected || inv.expected || 0, avgCost: inv.average_cost || inv.cost || 0, salesPrice: vInfo.salesPrice
    });
  }

  Logger.log('Katana: ' + locations.length + ' locations, ' + products.length + ' products, ' +
    materials.length + ' materials, ' + variants.length + ' variants, ' +
    inventory.length + ' inventory (' + items.length + ' tracked)');

  var btCount = Object.keys(batchTrackingSkus).length;
  Logger.log('Katana batch_tracking: ' + btCount + ' SKUs marked as batch-tracked');

  return { items: items, locationMap: locationMap, variantMap: variantMap, batchTrackingSkus: batchTrackingSkus, rawCounts: { locations: locations.length, products: products.length, materials: materials.length, variants: variants.length, inventory: inventory.length } };
}

/**
 * Fetch ALL batch_stocks in bulk paginated calls (2-3 instead of 200+).
 * Returns { batchesBySku: { sku: [{batchNumber, qty, expiry, locations}] }, lotTrackedSkus: { sku: true } }
 */
function fetchAllBatchData(katanaData) {
  var allBatchStocks = katanaFetchAllPages('/batch_stocks');
  Logger.log('Fetched ' + allBatchStocks.length + ' batch_stock records (bulk)');

  var batchesBySku = {};
  var lotTrackedSkus = {};

  var variantLocMap = {};
  for (var i = 0; i < katanaData.items.length; i++) {
    var item = katanaData.items[i];
    if (item.sku && !isSkippedSku(item.sku)) {
      if (!variantLocMap[item.variantId]) variantLocMap[item.variantId] = [];
      variantLocMap[item.variantId].push(item.locationName);
    }
  }

  // Log first batch record fields for debugging
  if (allBatchStocks.length > 0) {
    Logger.log('Batch fields: ' + Object.keys(allBatchStocks[0]).join(', '));
    Logger.log('Sample batch: ' + JSON.stringify(allBatchStocks[0]).substring(0, 500));
  }

  var skippedCount = 0;
  for (var b = 0; b < allBatchStocks.length; b++) {
    var batch = allBatchStocks[b];
    var variantId = batch.variant_id;
    var vInfo = katanaData.variantMap[variantId];
    if (!vInfo) {
      if (skippedCount < 10) Logger.log('Skipped batch: variant_id=' + variantId + ' not in variantMap, batch=' + (batch.batch_number || batch.nr || '?'));
      skippedCount++;
      continue;
    }
    if (!vInfo.sku || isSkippedSku(vInfo.sku)) continue;

    var batchQty = parseFloat(batch.quantity_in_stock || batch.in_stock || batch.quantity || 0);
    if (isNaN(batchQty) || batchQty <= 0) continue;

    var sku = vInfo.sku;
    lotTrackedSkus[sku] = true;
    if (!batchesBySku[sku]) batchesBySku[sku] = [];

    var bestBefore = batch.expiration_date || batch.best_before || batch.expiry_date || batch.date_code || '';
    if (bestBefore && String(bestBefore).indexOf('T') > -1) {
      bestBefore = String(bestBefore).split('T')[0];
    }

    // Batch location: try location_id, warehouse_id, stock_location_id, then fallback to variant locations
    var batchLoc = '';
    var locId = batch.location_id || batch.warehouse_id || batch.stock_location_id || '';
    if (locId && katanaData.locationMap[locId]) {
      batchLoc = katanaData.locationMap[locId];
    } else if (locId && katanaData.locationMap[String(locId)]) {
      batchLoc = katanaData.locationMap[String(locId)];
    } else {
      var locs = variantLocMap[variantId] || variantLocMap[String(variantId)] || [];
      batchLoc = locs[0] || '';
    }

    var isTrackedBatchLoc = false;
    for (var tl2 = 0; tl2 < TRACKED_LOCATIONS.length; tl2++) {
      if (batchLoc === TRACKED_LOCATIONS[tl2]) { isTrackedBatchLoc = true; break; }
    }
    if (!isTrackedBatchLoc) continue;

    batchesBySku[sku].push({
      batchNumber: batch.batch_number || batch.nr || batch.number || '',
      qty: batchQty,
      expiry: bestBefore,
      location: batchLoc,
      uom: vInfo.uom || 'ea'
    });
  }

  Logger.log('Batch data: ' + Object.keys(lotTrackedSkus).length + ' lot-tracked SKUs from ' + allBatchStocks.length + ' records');
  return { batchesBySku: batchesBySku, lotTrackedSkus: lotTrackedSkus };
}
