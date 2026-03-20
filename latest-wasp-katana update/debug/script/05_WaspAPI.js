// ============================================
// 05_WaspAPI.gs - WASP API FUNCTIONS
// ============================================
// All WASP InventoryCloud API interactions
// FIXED: Changed const to var for Google Apps Script compatibility
// ============================================

/**
 * Generic WASP API call (POST)
 */
function waspApiCall(url, payload) {
  var payloadString = typeof payload === 'string'
    ? payload
    : JSON.stringify(payload);

  var options = {
    method: 'POST',
    headers: {
      'Authorization': 'Bearer ' + CONFIG.WASP_TOKEN,
      'Content-Type': 'application/json'
    },
    payload: payloadString,
    muteHttpExceptions: true
  };

  try {
    var response = UrlFetchApp.fetch(url, options);
    var code = response.getResponseCode();
    var body = response.getContentText();

    var success = code === 200 && !body.includes('HasError":true');

    // Only log errors
    if (!success) {
      logToSheet('WASP_API_ERROR', {
        url: url.replace(CONFIG.WASP_BASE_URL, ''),
        statusCode: code
      }, body.substring(0, 500));
    }

    return {
      success: success,
      code: code,
      response: body
    };
  } catch (error) {
    logToSheet('WASP_API_ERROR', { url: url }, { error: error.message });
    return {
      success: false,
      code: 0,
      response: error.message
    };
  }
}

function getWaspResultRows_(response) {
  if (!response) return [];
  if (response.Data && response.Data.ResultList) {
    return Array.isArray(response.Data.ResultList) ? response.Data.ResultList : [response.Data.ResultList];
  }

  var rows = response.Data || response.data || [];
  return Array.isArray(rows) ? rows : [rows];
}

function normalizeBusinessDate_(value) {
  if (!value) return '';

  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }

  var text = String(value).trim();
  if (!text) return '';

  var slashMatch = text.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (slashMatch) {
    var smm = slashMatch[1];
    var sdd = slashMatch[2];
    var syyyy = slashMatch[3];
    if (smm.length === 1) smm = '0' + smm;
    if (sdd.length === 1) sdd = '0' + sdd;
    return syyyy + '-' + smm + '-' + sdd;
  }

  if (/^\d{4}-\d{2}-\d{2}$/.test(text)) return text;

  // Preserve ISO/business date tokens exactly as stored instead of shifting
  // them through the script timezone. Katana expiry timestamps often arrive
  // as midnight UTC, and formatting them in local time can move them back a day.
  var isoPrefix = text.match(/^(\d{4}-\d{2}-\d{2})/);
  if (isoPrefix) return isoPrefix[1];

  var hasTimeComponent = text.indexOf('T') >= 0 || /\d{2}:\d{2}/.test(text) || /[zZ]|[+\-]\d{2}:?\d{2}$/.test(text);
  if (hasTimeComponent) {
    var parsed = new Date(text);
    if (!isNaN(parsed.getTime())) {
      return Utilities.formatDate(parsed, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    }
  }

  return text;
}

function getRawDateToken_(value) {
  if (!value) return '';
  var text = String(value).trim();
  if (!text) return '';

  var slashMatch = text.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (slashMatch) {
    var mm = slashMatch[1];
    var dd = slashMatch[2];
    var yyyy = slashMatch[3];
    if (mm.length === 1) mm = '0' + mm;
    if (dd.length === 1) dd = '0' + dd;
    return yyyy + '-' + mm + '-' + dd;
  }

  var isoPrefix = text.match(/^(\d{4}-\d{2}-\d{2})/);
  if (isoPrefix) return isoPrefix[1];

  return '';
}

function buildComparableDateSet_(value) {
  var seen = {};
  var values = [];

  function add(v) {
    if (!v || seen[v]) return;
    seen[v] = true;
    values.push(v);
  }

  add(normalizeBusinessDate_(value));
  add(getRawDateToken_(value));
  return values;
}

function dateSetsOverlap_(left, right) {
  if (!left || !left.length || !right || !right.length) return false;
  var rightMap = {};
  for (var i = 0; i < right.length; i++) rightMap[right[i]] = true;
  for (var j = 0; j < left.length; j++) {
    if (rightMap[left[j]]) return true;
  }
  return false;
}

function parseComparableIsoDate_(value) {
  if (!value) return null;
  var text = String(value).trim();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(text)) return null;
  var parts = text.split('-');
  var year = Number(parts[0]);
  var month = Number(parts[1]);
  var day = Number(parts[2]);
  if (!year || !month || !day) return null;
  return Date.UTC(year, month - 1, day);
}

function datesWithinToleranceDays_(leftValue, rightValue, maxDays) {
  if (!maxDays || maxDays < 0) return false;
  var leftDate = parseComparableIsoDate_(leftValue);
  var rightDate = parseComparableIsoDate_(rightValue);
  if (leftDate === null || rightDate === null) return false;
  var diffDays = Math.abs(leftDate - rightDate) / 86400000;
  return diffDays <= maxDays;
}

// ============================================
// INVENTORY TRANSACTIONS
// ============================================

/**
 * Add inventory to WASP with lot/date tracking
 */
function waspAddInventoryWithLot(itemNumber, quantity, locationCode, lotNumber, expiryDate, notes, siteName) {
  var url = CONFIG.WASP_BASE_URL + '/public-api/transactions/item/add';
  var normalizedExpiry = normalizeBusinessDate_(expiryDate);

  var payload = [{
    ItemNumber: itemNumber,
    Quantity: quantity,
    SiteName: siteName || CONFIG.WASP_SITE,
    LocationCode: locationCode,
    Lot: lotNumber || '',
    Notes: notes || ''
  }];

  if (normalizedExpiry) {
    payload[0].DateCode = normalizedExpiry;
  }

  // Removed verbose WASP logging

  return waspApiCall(url, payload);
}

/**
 * Add inventory to WASP (simple, no lot)
 */
function waspAddInventory(itemNumber, quantity, locationCode, notes, siteName) {
  var url = CONFIG.WASP_BASE_URL + '/public-api/transactions/item/add';

  var payload = [{
    ItemNumber: itemNumber,
    Quantity: quantity,
    SiteName: siteName || CONFIG.WASP_SITE,
    LocationCode: locationCode,
    Notes: notes || ''
  }];

  return waspApiCall(url, payload);
}

/**
 * Remove inventory from WASP
 */
function waspRemoveInventory(itemNumber, quantity, locationCode, notes, siteName) {
  var url = CONFIG.WASP_BASE_URL + '/public-api/transactions/item/remove';

  var payload = [{
    ItemNumber: itemNumber,
    Quantity: quantity,
    SiteName: siteName || CONFIG.WASP_SITE,
    LocationCode: locationCode,
    Notes: notes || ''
  }];

  return waspApiCall(url, payload);
}

/**
 * Remove inventory from WASP with lot tracking
 * Required for lot-tracked items — WASP needs Lot to identify which record to decrement
 */
function waspRemoveInventoryWithLot(itemNumber, quantity, locationCode, lotNumber, notes, siteName, dateCode) {
  var url = CONFIG.WASP_BASE_URL + '/public-api/transactions/item/remove';
  var normalizedDateCode = normalizeBusinessDate_(dateCode);

  var payload = [{
    ItemNumber: itemNumber,
    Quantity: quantity,
    SiteName: siteName || CONFIG.WASP_SITE,
    LocationCode: locationCode,
    Lot: lotNumber || '',
    Notes: notes || ''
  }];

  // WASP requires DateCode for lot-tracked items
  if (normalizedDateCode) {
    payload[0].DateCode = normalizedDateCode;
  }

  return waspApiCall(url, payload);
}

/**
 * Look up item in WASP to find lot number at a location.
 * Uses advancedinventorysearch which returns inventory rows with Lot + DateCode.
 * Paginates through all pages (SearchPattern doesn't filter by ItemNumber).
 * Returns first lot found at the specified location, or null.
 */
function waspLookupItemLots(itemNumber, locationCode, siteName) {
  var url = CONFIG.WASP_BASE_URL + '/public-api/ic/item/advancedinventorysearch';
  var maxPages = 20; // Safety limit

  try {
    for (var page = 1; page <= maxPages; page++) {
      var payload = {
        SearchPattern: itemNumber,
        PageSize: 100,
        PageNumber: page
      };

      var result = waspApiCall(url, payload);
      if (!result.success) {
        logToSheet('WASP_LOT_LOOKUP_FAIL', { item: itemNumber, location: locationCode, page: page }, { error: 'advancedinventorysearch failed' });
        return null;
      }

      var response = JSON.parse(result.response);
      var rows = getWaspResultRows_(response);

      // No more results — stop paginating
      if (rows.length === 0) break;

      for (var i = 0; i < rows.length; i++) {
        var row = rows[i];
        if (!row) continue;

        var rowItem = row.ItemNumber || row.itemNumber || '';
        if (rowItem !== itemNumber) continue;

        var rowSite = row.SiteName || row.siteName || '';
        if (siteName && rowSite && rowSite !== siteName) continue;

        var rowLoc = row.LocationCode || row.locationCode || '';
        if (locationCode && rowLoc !== locationCode) continue;

        var rowLot = row.Lot || row.lot || row.LotNumber || row.lotNumber || '';
        if (rowLot) return rowLot;
      }

      // If fewer than PageSize results, this was the last page
      if (rows.length < 100) break;
    }

    return null;
  } catch (e) {
    logToSheet('WASP_LOT_LOOKUP_ERROR', { item: itemNumber, error: e.message }, {});
    return null;
  }
}

/**
 * Look up item in WASP to find lot AND dateCode at a location.
 * Paginates through all pages (SearchPattern doesn't filter by ItemNumber).
 * Returns { lot: string, dateCode: string } or null.
 */
function waspLookupItemLotAndDate(itemNumber, locationCode, siteName) {
  var url = CONFIG.WASP_BASE_URL + '/public-api/ic/item/advancedinventorysearch';
  var maxPages = 20; // Safety limit

  try {
    for (var page = 1; page <= maxPages; page++) {
      var payload = {
        SearchPattern: itemNumber,
        PageSize: 100,
        PageNumber: page
      };

      var result = waspApiCall(url, payload);
      if (!result.success) return null;

      var response = JSON.parse(result.response);
      var rows = getWaspResultRows_(response);

      if (rows.length === 0) break;

      for (var i = 0; i < rows.length; i++) {
        var row = rows[i];
        if (!row) continue;

        var rowItem = row.ItemNumber || row.itemNumber || '';
        if (rowItem !== itemNumber) continue;

        var rowSite = row.SiteName || row.siteName || '';
        if (siteName && rowSite && rowSite !== siteName) continue;

        var rowLoc = row.LocationCode || row.locationCode || '';
        if (locationCode && rowLoc !== locationCode) continue;

        var rowLot = row.Lot || row.lot || row.LotNumber || row.lotNumber || '';
        if (rowLot) {
          return {
            lot: rowLot,
            dateCode: normalizeExactLotDate_(row.DateCode || row.dateCode || '')
          };
        }
      }

      if (rows.length < 100) break;
    }

    return null;
  } catch (e) {
    return null;
  }
}

function normalizeExactLotDate_(value) {
  return normalizeBusinessDate_(value);
}

/**
 * Look up an exact lot row in WASP for one item/location/lot combination.
 * If expectedDateCode is provided, the date must match too.
 * If expectedDateCode is missing and the same lot exists with multiple dates,
 * returns { ambiguous: true } so callers can skip instead of guessing.
 * Otherwise returns { lot, dateCode, siteName, locationCode, qtyAvailable } or null.
 */
function waspLookupExactLotAndDate(itemNumber, locationCode, lotNumber, siteName, expectedDateCode) {
  if (!itemNumber || !lotNumber) return null;

  var url = CONFIG.WASP_BASE_URL + '/public-api/ic/item/advancedinventorysearch';
  var maxPages = 20;
  var normalizedExpectedDate = normalizeExactLotDate_(expectedDateCode);
  var expectedDateSet = buildComparableDateSet_(expectedDateCode);
  var matches = [];

  try {
    for (var page = 1; page <= maxPages; page++) {
      var payload = {
        SearchPattern: itemNumber,
        PageSize: 100,
        PageNumber: page
      };

      var result = waspApiCall(url, payload);
      if (!result.success) return null;

      var response = JSON.parse(result.response);
      var rows = getWaspResultRows_(response);

      if (rows.length === 0) break;

      for (var i = 0; i < rows.length; i++) {
        var row = rows[i];
        if (!row) continue;

        var rowItem = row.ItemNumber || row.itemNumber || '';
        if (rowItem !== itemNumber) continue;

        var rowSite = row.SiteName || row.siteName || '';
        if (siteName && rowSite && rowSite !== siteName) continue;

        var rowLoc = row.LocationCode || row.locationCode || '';
        if (locationCode && rowLoc !== locationCode) continue;

        var rowLot = row.Lot || row.lot || row.LotNumber || row.lotNumber || '';
        if (rowLot !== lotNumber) continue;

        var rowDateRaw = row.DateCode || row.dateCode || '';
        var rowDateCode = normalizeExactLotDate_(rowDateRaw);
        var rowDateSet = buildComparableDateSet_(rowDateRaw);
        var exactDateMatch = !expectedDateSet.length || dateSetsOverlap_(expectedDateSet, rowDateSet);
        var tolerantDateMatch = !exactDateMatch && normalizedExpectedDate &&
          datesWithinToleranceDays_(normalizedExpectedDate, rowDateCode, 1);
        if (expectedDateSet.length && !exactDateMatch && !tolerantDateMatch) continue;

        matches.push({
          lot: rowLot,
          dateCode: rowDateCode,
          siteName: rowSite,
          locationCode: rowLoc,
          dateToleranceApplied: tolerantDateMatch,
          qtyAvailable: parseFloat(
            row.QtyAvailable || row.TotalInHouse || row.TotalAvailable || row.Quantity || 0
          ) || 0
        });
      }

      if (rows.length < 100) break;
    }

    if (matches.length === 0) return null;

    if (!expectedDateSet.length) {
      var canonicalDate = matches[0].dateCode || '';
      for (var m = 1; m < matches.length; m++) {
        if ((matches[m].dateCode || '') !== canonicalDate) {
          return { ambiguous: true, lot: lotNumber };
        }
      }
    }

    return matches[0];
  } catch (e) {
    return null;
  }
}

/**
 * Look up all WASP rows for one exact item/location/lot combination, regardless of date.
 * Returns an array of { lot, dateCode, siteName, locationCode, qtyAvailable }.
 */
function waspLookupLotRows(itemNumber, locationCode, siteName, lotNumber) {
  if (!itemNumber || !lotNumber) return [];

  var url = CONFIG.WASP_BASE_URL + '/public-api/ic/item/advancedinventorysearch';
  var maxPages = 20;
  var matches = [];

  try {
    for (var page = 1; page <= maxPages; page++) {
      var payload = {
        SearchPattern: itemNumber,
        PageSize: 100,
        PageNumber: page
      };

      var result = waspApiCall(url, payload);
      if (!result.success) return [];

      var response = JSON.parse(result.response);
      var rows = getWaspResultRows_(response);
      if (rows.length === 0) break;

      for (var i = 0; i < rows.length; i++) {
        var row = rows[i];
        if (!row) continue;

        var rowItem = row.ItemNumber || row.itemNumber || '';
        if (rowItem !== itemNumber) continue;

        var rowSite = row.SiteName || row.siteName || '';
        if (siteName && rowSite && rowSite !== siteName) continue;

        var rowLoc = row.LocationCode || row.locationCode || '';
        if (locationCode && rowLoc !== locationCode) continue;

        var rowLot = row.Lot || row.lot || row.LotNumber || row.lotNumber || '';
        if (rowLot !== lotNumber) continue;

        matches.push({
          lot: rowLot,
          dateCode: normalizeExactLotDate_(row.DateCode || row.dateCode || ''),
          siteName: rowSite,
          locationCode: rowLoc,
          qtyAvailable: parseFloat(
            row.QtyAvailable || row.TotalInHouse || row.TotalAvailable || row.Quantity || 0
          ) || 0
        });
      }

      if (rows.length < 100) break;
    }
  } catch (e) {
    return [];
  }

  return matches;
}

/**
 * Query WASP for ALL lot records for an item at a specific location.
 * Used as a fallback when Katana's batch_stocks API can't resolve a lot.
 * Returns array of { lot, expiry, quantity } — one entry per distinct lot.
 * Returns empty array on error or if no lots found.
 */
function waspLookupAllLots_(itemNumber, locationCode, siteName) {
  try {
    var url = CONFIG.WASP_BASE_URL + '/public-api/ic/item/inventorysearch';
    var payload = { ItemNumber: itemNumber, SiteName: siteName || CONFIG.WASP_SITE };
    if (locationCode) payload.LocationCode = locationCode;

    var result = waspApiCall(url, payload);
    if (!result.success) return [];

    var response = JSON.parse(result.response);
    var data = response.Data || response.data || [];
    var lots = [];

    for (var i = 0; i < data.length; i++) {
      var rec = data[i];
      var recLoc = rec.LocationCode || rec.Location || '';
      if (locationCode && recLoc !== locationCode) continue;

      var lot = rec.Lot || rec.LotNumber || rec.BatchNumber || '';
      if (!lot) continue;

      var expiry = normalizeBusinessDate_(rec.DateCode || rec.ExpiryDate || rec.ExpDate || '');
      var qty = parseFloat(rec.Quantity || rec.QuantityOnHand || 0) || 0;
      lots.push({ lot: lot, expiry: expiry, quantity: qty });
    }

    return lots;
  } catch (e) {
    Logger.log('waspLookupAllLots_ error: ' + e.message);
    return [];
  }
}

/**
 * MO lot-only lookup.
 * Finds a WASP row by item/location/lot while intentionally ignoring expiry as
 * a matching criterion. If multiple rows share the same lot, prefer a row with
 * enough quantity for the requested deduction; otherwise use the largest row.
 */
function waspLookupLotForRemoval(itemNumber, locationCode, siteName, lotNumber, requiredQty) {
  var rows = waspLookupLotRows(itemNumber, locationCode, siteName, lotNumber);
  if (!rows || rows.length === 0) return null;

  var neededQty = parseFloat(requiredQty || 0) || 0;
  var totalQty = 0;
  var bestAny = null;
  var bestAnyQty = -1;
  var bestFit = null;
  var bestFitQty = -1;

  for (var i = 0; i < rows.length; i++) {
    var row = rows[i] || {};
    var rowQty = parseFloat(row.qtyAvailable || 0) || 0;
    totalQty += rowQty;

    if (rowQty > bestAnyQty) {
      bestAny = row;
      bestAnyQty = rowQty;
    }

    if (neededQty > 0 && rowQty >= neededQty && rowQty > bestFitQty) {
      bestFit = row;
      bestFitQty = rowQty;
    }
  }

  var picked = bestFit || bestAny || rows[0];
  return {
    lot: picked.lot || String(lotNumber || '').trim(),
    dateCode: picked.dateCode || '',
    siteName: picked.siteName || siteName || '',
    locationCode: picked.locationCode || locationCode || '',
    qtyAvailable: parseFloat(picked.qtyAvailable || 0) || 0,
    totalQtyAvailable: totalQty,
    rowCount: rows.length
  };
}

/**
 * Look up all WASP lot rows for one item/location, regardless of lot/date.
 * Returns an array of { lot, dateCode, siteName, locationCode, qtyAvailable }.
 */
function waspLookupAllLotRows(itemNumber, locationCode, siteName) {
  if (!itemNumber) return [];

  var url = CONFIG.WASP_BASE_URL + '/public-api/ic/item/advancedinventorysearch';
  var maxPages = 20;
  var matches = [];

  try {
    for (var page = 1; page <= maxPages; page++) {
      var payload = {
        SearchPattern: itemNumber,
        PageSize: 100,
        PageNumber: page
      };

      var result = waspApiCall(url, payload);
      if (!result.success) return [];

      var response = JSON.parse(result.response);
      var rows = getWaspResultRows_(response);
      if (rows.length === 0) break;

      for (var i = 0; i < rows.length; i++) {
        var row = rows[i];
        if (!row) continue;

        var rowItem = row.ItemNumber || row.itemNumber || '';
        if (rowItem !== itemNumber) continue;

        var rowSite = row.SiteName || row.siteName || '';
        if (siteName && rowSite && rowSite !== siteName) continue;

        var rowLoc = row.LocationCode || row.locationCode || '';
        if (locationCode && rowLoc !== locationCode) continue;

        var rowLot = row.Lot || row.lot || row.LotNumber || row.lotNumber || '';
        if (!rowLot) continue;

        matches.push({
          lot: rowLot,
          dateCode: normalizeExactLotDate_(row.DateCode || row.dateCode || ''),
          siteName: rowSite,
          locationCode: rowLoc,
          qtyAvailable: parseFloat(
            row.QtyAvailable || row.TotalInHouse || row.TotalAvailable || row.Quantity || 0
          ) || 0
        });
      }

      if (rows.length < 100) break;
    }
  } catch (e) {
    return [];
  }

  return matches;
}

/**
 * Adjust inventory in WASP (can be + or -)
 */
function waspAdjustInventory(itemNumber, quantity, locationCode, notes) {
  var url = CONFIG.WASP_BASE_URL + '/public-api/transactions/item/adjust';

  var payload = [{
    ItemNumber: itemNumber,
    Quantity: quantity,
    SiteName: CONFIG.WASP_SITE,
    LocationCode: locationCode,
    Notes: notes || ''
  }];

  return waspApiCall(url, payload);
}

// ============================================
// PICK ORDERS
// ============================================

/**
 * Create pick order in WASP
 */
function waspCreatePickOrder(payload) {
  var url = CONFIG.WASP_BASE_URL + '/public-api/ic/pickpackshiporder/create';
  return waspApiCall(url, payload);
}

/**
 * Get pick order by number
 */
function waspGetPickOrderByNumber(orderNumber) {
  var url = CONFIG.WASP_BASE_URL + '/public-api/ic/pickpackshiporder/getordersbynumber';

  var payload = {
    OrderNumbers: [orderNumber]
  };

  return waspApiCall(url, payload);
}

/**
 * Check if a pick order is fully picked
 */
function checkPickOrderComplete(pickOrderNumber) {
  var result = waspGetPickOrderByNumber(pickOrderNumber);

  if (!result.success) {
    logToSheet('PICK_ORDER_CHECK_ERROR', { pickOrderNumber: pickOrderNumber }, result);
    return false;
  }

  try {
    var response = JSON.parse(result.response);
    var orders = response.Data || response.data || [];

    if (orders.length === 0) {
      logToSheet('PICK_ORDER_NOT_FOUND_API', { pickOrderNumber: pickOrderNumber }, {});
      return false;
    }

    var order = orders[0];
    var lines = order.PickOrderLines || [];

    var allPicked = true;
    for (var i = 0; i < lines.length; i++) {
      var line = lines[i];
      var outstanding = line.OutstandingQuantity || 0;
      if (outstanding > 0) {
        allPicked = false;
        logToSheet('PICK_LINE_PENDING', {
          item: line.ItemNumber,
          requested: line.Quantity,
          outstanding: outstanding
        }, {});
      }
    }

    return allPicked;

  } catch (e) {
    logToSheet('PICK_ORDER_PARSE_ERROR', {
      pickOrderNumber: pickOrderNumber,
      error: e.message
    }, {});
    return false;
  }
}

// ============================================
// ITEM COST SYNC
// ============================================

/**
 * Update a single item's unit cost in WASP to match Katana.
 * Called during F1 (PO receive) to keep costs in sync on every receive.
 * Also called by syncCostsFromSheet() for a one-time bulk update.
 *
 * UOM alignment (StockingUnit/PurchaseUnit/SalesUnit) is handled by the daily
 * syncAllItemMetadata() — this function only touches cost to avoid clobbering UOM on every receive.
 *
 * @param {string} itemNumber - WASP ItemNumber / SKU
 * @param {number} cost       - Unit cost
 * @param {string} uom        - Unit of measure — only sent if non-empty (avoids overwriting with wrong value)
 */
function waspUpdateItemCost(itemNumber, cost, uom) {
  if (!itemNumber || !(cost > 0)) return { success: false, code: 0, response: 'skipped: invalid params' };
  var payload = {
    ItemNumber:    itemNumber,
    Cost:          cost,
    DimensionInfo: {
      DimensionUnit: 'Inch',
      WeightUnit:    'Pound',
      VolumeUnit:    'Cubic Inch'
    }
  };
  // StockingUnit must always be present — WASP returns -57072 if omitted.
  // Use provided UOM; fall back to 'Each' only when caller passes nothing.
  var unitToUse = uom || 'Each';
  payload.StockingUnit = unitToUse;
  payload.PurchaseUnit = unitToUse;
  payload.SalesUnit    = unitToUse;
  var url = CONFIG.WASP_BASE_URL + '/public-api/ic/item/updateInventoryItems';
  return waspApiCall(url, [payload]);
}
