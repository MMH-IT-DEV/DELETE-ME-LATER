/**
 * Wasp InventoryCloud API
 */

var WASP_DEFAULT_SITE = 'MMH Mayfair';

function getWaspBase() {
  var instance = PropertiesService.getScriptProperties().getProperty('WASP_INSTANCE') || 'mymagichealer';
  return 'https://' + instance + '.waspinventorycloud.com/public-api/';
}

function waspFetch(endpoint, payload) {
  var token = PropertiesService.getScriptProperties().getProperty('WASP_API_TOKEN');
  if (!token) throw new Error('Wasp API token not set.');

  var options = {
    method: 'post',
    headers: { 'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json' },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  var response = UrlFetchApp.fetch(getWaspBase() + endpoint, options);
  var code = response.getResponseCode();
  if (code === 429) {
    var resetSec = parseInt(response.getHeaders()['wasp-overloadprotection-reset'] || '5');
    Utilities.sleep(resetSec * 1000);
    response = UrlFetchApp.fetch(getWaspBase() + endpoint, options);
    code = response.getResponseCode();
  }
  if (code !== 200) {
    throw new Error('Wasp ' + code + ' on ' + endpoint + ': ' + response.getContentText().substring(0, 300));
  }
  var body = response.getContentText();
  // WASP occasionally returns an HTML page (server blip) with status 200.
  // Use trim() to handle leading whitespace/BOM before <!DOCTYPE.
  // Detect and retry once after a short pause before failing.
  if (body.trim().charAt(0) === '<') {
    Logger.log('waspFetch: HTML response on ' + endpoint + ' — retrying in 5s');
    Utilities.sleep(5000);
    response = UrlFetchApp.fetch(getWaspBase() + endpoint, options);
    code = response.getResponseCode();
    body = response.getContentText();
    if (body.trim().charAt(0) === '<') {
      throw new Error('Wasp returned HTML twice on ' + endpoint + ' — service temporarily unavailable, try again in a moment.');
    }
  }
  return JSON.parse(body);
}

function waspApiCall(url, payload) {
  var token = PropertiesService.getScriptProperties().getProperty('WASP_API_TOKEN');
  var payloadString = typeof payload === 'string' ? payload : JSON.stringify(payload);
  var options = {
    method: 'POST',
    headers: { 'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json' },
    payload: payloadString,
    muteHttpExceptions: true
  };
  try {
    var response = UrlFetchApp.fetch(url, options);
    var code = response.getResponseCode();
    var body = response.getContentText();
    var success = code === 200 && body.indexOf('HasError\":true') === -1;
    if (!success) Logger.log('waspApiCall FAILED: code=' + code + ' url=' + url + ' body=' + body.substring(0, 500));
    return { success: success, code: code, response: body };
  } catch (error) {
    return { success: false, code: 0, response: error.message };
  }
}

/**
 * Parse a WASP date value into a clean YYYY-MM-DD string.
 * Returns '' for null/undefined/empty, invalid dates, or dates before 2020.
 */
function parseWaspDate(value) {
  if (value === null || value === undefined || value === '') return '';

  var str = '';

  if (value instanceof Date) {
    // Convert Date object to YYYY-MM-DD
    var y = value.getFullYear();
    if (y < 2020) return '';
    var m = value.getMonth() + 1;
    var d = value.getDate();
    str = y + '-' + (m < 10 ? '0' + m : m) + '-' + (d < 10 ? '0' + d : d);
    return str;
  }

  str = String(value);
  if (str === '') return '';

  // ISO string with time component — take date part only
  if (str.indexOf('T') !== -1) {
    str = str.split('T')[0];
  }

  // Reject known bad sentinel date
  if (str === '1899-12-30') return '';

  // Reject if year portion is before 2020
  var parts = str.split('-');
  if (parts.length >= 1) {
    var year = parseInt(parts[0], 10);
    if (!isNaN(year) && year < 2020) return '';
  }

  return str;
}

function fetchAllWaspData() {
  var allItems = [];
  var page = 1;
  var hasMore = true;
  var totalCount = 0;

  while (hasMore) {
    var payload = { PageSize: 500, PageNumber: page };
    if (totalCount > 0) payload.TotalCountFromPriorFetch = totalCount;
    var result = waspFetch('ic/item/advancedinventorysearch', payload);
    var data = result.Data || [];
    for (var d = 0; d < data.length; d++) { allItems.push(data[d]); }
    if (page === 1) totalCount = result.TotalRecordsLongCount || 0;
    hasMore = result.HasSuccessWithMoreDataRemaining === true;
    page++;
    if (page > 50) break;
  }

  // Debug: log ALL field names from first raw item so we can verify the correct fields
  if (allItems.length > 0) {
    var firstItem = allItems[0];
    var keys = Object.keys(firstItem);
    Logger.log('=== WASP RAW ITEM FIELD NAMES (' + keys.length + ' fields) ===');
    Logger.log(keys.join(', '));
    Logger.log('=== WASP RAW ITEM SAMPLE VALUES ===');
    for (var ki = 0; ki < keys.length; ki++) {
      Logger.log('  ' + keys[ki] + ': ' + JSON.stringify(firstItem[keys[ki]]));
    }
  }

  var noSiteCount = 0;
  var items = [];

  for (var i = 0; i < allItems.length; i++) {
    var item = allItems[i];

    // Qty: try TotalInHouse first (confirmed field name), then fallbacks
    var qty = parseFloat(
      item.TotalInHouse !== undefined && item.TotalInHouse !== null ? item.TotalInHouse :
      item.TotalAvailable !== undefined && item.TotalAvailable !== null ? item.TotalAvailable :
      item.QuantityAvailable !== undefined && item.QuantityAvailable !== null ? item.QuantityAvailable :
      item.Available !== undefined && item.Available !== null ? item.Available :
      item.Quantity !== undefined && item.Quantity !== null ? item.Quantity :
      0
    );

    // Site / Location — warn on missing, never silently drop
    var siteName = item.SiteName || item.Site || item.SiteCode || '';
    var locationCode = item.LocationCode || item.Location || item.LocationName || '';
    if (!siteName) {
      noSiteCount++;
      siteName = '(unknown)';
    }

    // Detect lot tracking ONLY when API actually returns the field.
    // advancedinventorysearch returns inventory rows — may NOT include item-level LotTracking.
    // If no field found, ltVal stays undefined so writeWaspTab falls through to lotTrackedSkus heuristic.
    var ltVal;
    if (item.LotTracking !== undefined && item.LotTracking !== null) { ltVal = item.LotTracking ? true : false; }
    else if (item.TrackLot !== undefined && item.TrackLot !== null) { ltVal = item.TrackLot ? true : false; }
    else if (item.IsLotTracked !== undefined && item.IsLotTracked !== null) { ltVal = item.IsLotTracked ? true : false; }
    else if (item.LotTrackingEnabled !== undefined && item.LotTrackingEnabled !== null) { ltVal = item.LotTrackingEnabled ? true : false; }

    items.push({
      itemNumber: item.ItemNumber || item.itemNumber || '',
      description: item.ItemDescription || item.Description || '',
      category: item.CategoryDescription || item.Category || '',
      siteName: siteName,
      locationCode: locationCode,
      qtyAvailable: qty,
      qtyOnHand: parseFloat(item.TotalOnHand || item.QuantityOnHand || item.OnHand || 0),
      lot: item.Lot || item.LotNumber || '',
      dateCode: parseWaspDate(item.DateCode || item.ExpiryDate || ''),
      serialNumber: item.SerialNumber || '',
      cost: parseFloat(item.Cost || 0),
      salesPrice: parseFloat(item.SalesPrice || 0),
      uom: item.StockingUnit || item.StockingUnitDescription || item.UOM || item.UnitOfMeasure || '',
      lotTracking: ltVal,
      lastUpdated: parseWaspDate(item.ItemLastUpdatedDate || item.LastUpdated || '')
    });
  }

  if (noSiteCount > 0) {
    Logger.log('WARNING: ' + noSiteCount + ' WASP items had no SiteName — marked as (unknown)');
  }

  var lotCount = 0;
  var uniqueSkus = {};
  for (var lc = 0; lc < items.length; lc++) {
    if (items[lc].lot) lotCount++;
    uniqueSkus[items[lc].itemNumber] = true;
  }
  Logger.log('Wasp: ' + items.length + ' records, ' + Object.keys(uniqueSkus).length +
    ' unique SKUs, ' + lotCount + ' with lots (before lot fetch), pages: ' + (page - 1));

  // Fetch lot-level inventory separately
  var lotItems = fetchWaspLotInventory();
  if (lotItems.length > 0) {
    // Build a set of existing lot records to avoid duplicates
    var existingKeys = {};
    for (var ek = 0; ek < items.length; ek++) {
      if (items[ek].lot) {
        existingKeys[items[ek].itemNumber + '|' + items[ek].lot + '|' + items[ek].siteName + '|' + items[ek].locationCode] = true;
      }
    }

    var added = 0;
    for (var li = 0; li < lotItems.length; li++) {
      var lk = lotItems[li].itemNumber + '|' + lotItems[li].lot + '|' + lotItems[li].siteName + '|' + lotItems[li].locationCode;
      if (!existingKeys[lk]) {
        items.push(lotItems[li]);
        added++;
      }
    }
    Logger.log('Wasp lots: ' + lotItems.length + ' lot records fetched, ' + added + ' new records merged');
  }

  return { items: items, totalCount: totalCount };
}

function fetchWaspLotInventory() {
  // Try dedicated lot endpoints first
  var endpoints = [
    'ic/lot/search',
    'ic/item/lotdetailsearch',
    'ic/item/lotsearch',
    'ic/item/inventorysearch'
  ];

  for (var ei = 0; ei < endpoints.length; ei++) {
    try {
      Logger.log('Trying Wasp lot endpoint: ' + endpoints[ei]);
      var result = waspFetch(endpoints[ei], { PageSize: 500, PageNumber: 1 });
      var data = result.Data || [];
      if (data.length === 0) { Logger.log('  -> empty result'); continue; }

      Logger.log('  -> SUCCESS! ' + data.length + ' records');
      Logger.log('  -> fields: ' + Object.keys(data[0]).join(', '));

      // Paginate
      var allLotItems = [];
      for (var d = 0; d < data.length; d++) { allLotItems.push(data[d]); }
      var hasMore = result.HasSuccessWithMoreDataRemaining === true;
      var lotPage = 2;
      var lotTotal = result.TotalRecordsLongCount || 0;

      while (hasMore && lotPage <= 50) {
        var payload2 = { PageSize: 500, PageNumber: lotPage };
        if (lotTotal > 0) payload2.TotalCountFromPriorFetch = lotTotal;
        var result2 = waspFetch(endpoints[ei], payload2);
        var data2 = result2.Data || [];
        for (var d2 = 0; d2 < data2.length; d2++) { allLotItems.push(data2[d2]); }
        hasMore = result2.HasSuccessWithMoreDataRemaining === true;
        lotPage++;
      }

      return mapLotItems(allLotItems);

    } catch (e) {
      Logger.log('  -> failed: ' + e.message);
    }
  }

  // Fallback: use advancedinventorysearch, fetch ALL items, filter client-side for lot data
  // Do NOT use Lot:'*' filter — it is unreliable and often returns 0 results.
  Logger.log('Dedicated lot endpoints failed — falling back to advancedinventorysearch client-side lot filter');
  try {
    var fallbackItems = [];
    var fbPage = 1;
    var fbHasMore = true;
    var fbTotal = 0;

    while (fbHasMore) {
      var fbPayload = { PageSize: 500, PageNumber: fbPage };
      if (fbTotal > 0) fbPayload.TotalCountFromPriorFetch = fbTotal;
      var fbResult = waspFetch('ic/item/advancedinventorysearch', fbPayload);
      var fbData = fbResult.Data || [];
      for (var fd = 0; fd < fbData.length; fd++) { fallbackItems.push(fbData[fd]); }
      if (fbPage === 1) fbTotal = fbResult.TotalRecordsLongCount || 0;
      fbHasMore = fbResult.HasSuccessWithMoreDataRemaining === true;
      fbPage++;
      if (fbPage > 50) break;
    }

    // Filter client-side: keep only records that have a non-empty lot field
    var withLots = [];
    for (var fl = 0; fl < fallbackItems.length; fl++) {
      var flItem = fallbackItems[fl];
      var lot = flItem.Lot || flItem.LotNumber || flItem.LotCode || '';
      if (lot) withLots.push(flItem);
    }

    Logger.log('advancedinventorysearch fallback: ' + fallbackItems.length + ' total, ' + withLots.length + ' with lot data');
    return mapLotItems(withLots);

  } catch (e) {
    Logger.log('advancedinventorysearch fallback also failed: ' + e.message);
  }

  Logger.log('No working lot endpoint found — returning empty array');
  return [];
}

/**
 * Map raw WASP item objects (from any lot-capable endpoint) into the standard format.
 * Only includes records that have a non-empty lot field.
 */
function mapLotItems(rawItems) {
  var items = [];
  for (var i = 0; i < rawItems.length; i++) {
    var item = rawItems[i];
    var lot = item.Lot || item.LotNumber || item.LotCode || '';
    if (!lot) continue;

    var siteName = item.SiteName || item.Site || item.SiteCode || '(unknown)';
    var locationCode = item.LocationCode || item.Location || item.LocationName || '';

    var qty = parseFloat(
      item.TotalInHouse !== undefined && item.TotalInHouse !== null ? item.TotalInHouse :
      item.TotalAvailable !== undefined && item.TotalAvailable !== null ? item.TotalAvailable :
      item.QuantityAvailable !== undefined && item.QuantityAvailable !== null ? item.QuantityAvailable :
      item.Available !== undefined && item.Available !== null ? item.Available :
      item.Quantity !== undefined && item.Quantity !== null ? item.Quantity :
      0
    );

    items.push({
      itemNumber: item.ItemNumber || item.itemNumber || '',
      description: item.ItemDescription || item.Description || '',
      category: item.CategoryDescription || item.Category || '',
      siteName: siteName,
      locationCode: locationCode,
      qtyAvailable: qty,
      qtyOnHand: parseFloat(item.TotalOnHand || item.QuantityOnHand || item.OnHand || 0),
      lot: lot,
      dateCode: parseWaspDate(item.DateCode || item.ExpiryDate || ''),
      serialNumber: item.SerialNumber || '',
      cost: parseFloat(item.Cost || 0),
      salesPrice: parseFloat(item.SalesPrice || 0),
      uom: item.StockingUnit || item.StockingUnitDescription || item.UOM || item.UnitOfMeasure || '',
      lastUpdated: parseWaspDate(item.ItemLastUpdatedDate || item.LastUpdated || '')
    });
  }
  return items;
}

/**
 * Fetches per-item lot tracking settings from WASP item catalog.
 * advancedinventorysearch returns inventory rows (no item settings),
 * so we need a separate item search to get LotTracking flag.
 *
 * @returns {Object} Map of SKU -> true/false for lot tracking
 */
function fetchWaspItemLotSettings() {
  var map = {};
  var endpoints = ['ic/item/search', 'ic/item/itemsearch', 'ic/item/detailsearch'];

  for (var ei = 0; ei < endpoints.length; ei++) {
    try {
      Logger.log('Trying WASP item search endpoint: ' + endpoints[ei]);
      var result = waspFetch(endpoints[ei], { PageSize: 500, PageNumber: 1 });
      var data = result.Data || [];
      if (result.Data && result.Data.ResultList) data = result.Data.ResultList;
      if (data.length === 0) { Logger.log('  -> empty result'); continue; }

      // Log field names from first item for debugging
      Logger.log('  -> SUCCESS! ' + data.length + ' records');
      Logger.log('  -> fields: ' + Object.keys(data[0]).join(', '));

      // Process first page
      for (var d = 0; d < data.length; d++) {
        var item = data[d];
        var sku = item.ItemNumber || '';
        if (sku) {
          map[sku] = (item.LotTracking || item.TrackLot || item.IsLotTracked || item.LotTrackingEnabled || false) ? true : false;
        }
      }

      // Paginate
      var hasMore = result.HasSuccessWithMoreDataRemaining === true;
      var page = 2;
      var totalCount = result.TotalRecordsLongCount || 0;

      while (hasMore && page <= 20) {
        var payload = { PageSize: 500, PageNumber: page };
        if (totalCount > 0) payload.TotalCountFromPriorFetch = totalCount;
        var result2 = waspFetch(endpoints[ei], payload);
        var data2 = result2.Data || [];
        if (result2.Data && result2.Data.ResultList) data2 = result2.Data.ResultList;
        for (var d2 = 0; d2 < data2.length; d2++) {
          var item2 = data2[d2];
          var sku2 = item2.ItemNumber || '';
          if (sku2) {
            map[sku2] = (item2.LotTracking || item2.TrackLot || item2.IsLotTracked || item2.LotTrackingEnabled || false) ? true : false;
          }
        }
        hasMore = result2.HasSuccessWithMoreDataRemaining === true;
        page++;
      }

      var lotCount = 0;
      var skus = Object.keys(map);
      for (var lc = 0; lc < skus.length; lc++) {
        if (map[skus[lc]]) lotCount++;
      }
      Logger.log('WASP item lot settings: ' + skus.length + ' items, ' + lotCount + ' lot-tracked (via ' + endpoints[ei] + ')');
      return map;

    } catch (e) {
      Logger.log('  -> failed: ' + e.message);
    }
  }

  Logger.log('No working WASP item search endpoint found — lot tracking map empty');
  return map;
}

/**
 * Pre-mark a WASP adjustment in the GAS webhook's script cache so F2 skips SA creation.
 * Called before every WASP add/remove API call.
 * Requires ScriptProperty 'GAS_WEBHOOK_URL' to be set.
 */
function enginPreMark_(itemNumber, locationCode, op) {
  var url = PropertiesService.getScriptProperties().getProperty('GAS_WEBHOOK_URL');
  if (!url) return;
  try {
    var resp = UrlFetchApp.fetch(url, {
      method: 'POST',
      contentType: 'application/json',
      payload: JSON.stringify({ action: 'engin_mark', sku: itemNumber, location: locationCode, op: op }),
      muteHttpExceptions: true
    });
    Logger.log('enginPreMark_ ' + op + ' ' + itemNumber + '@' + locationCode + ' → HTTP ' + resp.getResponseCode() + ' ' + resp.getContentText().substring(0, 150));
  } catch (e) {
    Logger.log('enginPreMark_ error: ' + e.message);
  }
}

function waspAddInventoryWithLot(itemNumber, quantity, locationCode, lotNumber, expiryDate, notes, siteName) {
  enginPreMark_(itemNumber, locationCode, 'add');
  var url = getWaspBase() + 'transactions/item/add';
  var lot = lotNumber || 'NO-LOT';
  var dateCode = expiryDate;
  if (!dateCode) { var future = new Date(); future.setFullYear(future.getFullYear() + 2); dateCode = future.toISOString().slice(0, 10); }
  var payload = [{ ItemNumber: itemNumber, Quantity: quantity, SiteName: siteName || WASP_DEFAULT_SITE, LocationCode: locationCode, Lot: lot, DateCode: dateCode, Notes: notes || '' }];
  return waspApiCall(url, payload);
}

function waspAddInventory(itemNumber, quantity, locationCode, notes, siteName) {
  enginPreMark_(itemNumber, locationCode, 'add');
  var url = getWaspBase() + 'transactions/item/add';
  var payload = [{ ItemNumber: itemNumber, Quantity: quantity, SiteName: siteName || WASP_DEFAULT_SITE, LocationCode: locationCode, Notes: notes || '' }];
  return waspApiCall(url, payload);
}

function waspRemoveInventoryWithLot(itemNumber, quantity, locationCode, lotNumber, notes, siteName, dateCode) {
  enginPreMark_(itemNumber, locationCode, 'remove');
  var url = getWaspBase() + 'transactions/item/remove';
  var payload = [{ ItemNumber: itemNumber, Quantity: quantity, SiteName: siteName || WASP_DEFAULT_SITE, LocationCode: locationCode, Lot: lotNumber || '', Notes: notes || '' }];
  if (dateCode) { payload[0].DateCode = dateCode; }
  return waspApiCall(url, payload);
}

function waspRemoveInventory(itemNumber, quantity, locationCode, notes, siteName) {
  enginPreMark_(itemNumber, locationCode, 'remove');
  var url = getWaspBase() + 'transactions/item/remove';
  var payload = [{ ItemNumber: itemNumber, Quantity: quantity, SiteName: siteName || WASP_DEFAULT_SITE, LocationCode: locationCode, Notes: notes || '' }];
  return waspApiCall(url, payload);
}

function waspAdjustInventory(itemNumber, quantity, locationCode, notes, siteName) {
  enginPreMark_(itemNumber, locationCode, 'add');
  var url = getWaspBase() + 'transactions/item/adjust';
  var payload = [{ ItemNumber: itemNumber, Quantity: quantity, SiteName: siteName || WASP_DEFAULT_SITE, LocationCode: locationCode, Notes: notes || '' }];
  return waspApiCall(url, payload);
}
