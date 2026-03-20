/**
 * Wasp InventoryCloud API
 */

var WASP_DEFAULT_SITE = 'MMH Mayfair';
var WASP_LOT_SETTINGS_CACHE_KEY = 'WASP_LOT_SETTINGS';
var WASP_LOT_SETTINGS_CACHE_TTL_MS = 30 * 60 * 1000;

function loadWaspCacheAnyAge_(key) {
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
    Logger.log('Failed to parse WASP cache ' + key + ': ' + e.message);
    return null;
  }
}

function loadWaspCacheFresh_(key, ttlMs) {
  var cached = loadWaspCacheAnyAge_(key);
  if (!cached || !cached.savedAt) return null;
  if ((new Date()).getTime() - cached.savedAt > ttlMs) return null;
  return cached.data;
}

function saveWaspCache_(key, data) {
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
  Logger.log('Saved WASP cache ' + key + ' (' + chunks + ' chunks)');
}

function fetchWaspCached_(key, ttlMs, fetchFn, forceRefresh) {
  if (!forceRefresh) {
    var fresh = loadWaspCacheFresh_(key, ttlMs);
    if (fresh !== null) {
      Logger.log('Using fresh WASP cache: ' + key);
      return fresh;
    }
  }

  try {
    var data = fetchFn();
    saveWaspCache_(key, data);
    return data;
  } catch (e) {
    var stale = loadWaspCacheAnyAge_(key);
    if (stale && stale.data !== null) {
      Logger.log('Using stale WASP cache for ' + key + ' after fetch failure: ' + e.message);
      return stale.data;
    }
    throw e;
  }
}

function normalizeWaspBaseUrl_(rawUrl) {
  var url = String(rawUrl || '').trim();
  if (!url) return '';
  if (!/^https?:\/\//i.test(url)) url = 'https://' + url;
  if (url.slice(-1) !== '/') url += '/';
  if (url.toLowerCase().indexOf('/public-api/') === -1) {
    if (url.toLowerCase().slice(-11) === '/public-api') url += '/';
    else url += 'public-api/';
  }
  return url;
}

function getWaspBase() {
  var explicitBase = normalizeWaspBaseUrl_(PropertiesService.getScriptProperties().getProperty('WASP_BASE_URL'));
  if (explicitBase) return explicitBase;
  var instance = PropertiesService.getScriptProperties().getProperty('WASP_INSTANCE') || 'mymagichealer';
  return normalizeWaspBaseUrl_('https://' + instance + '.waspinventorycloud.com');
}

function getWaspTenantOrigin_() {
  return getWaspBase().replace(/public-api\/?$/i, '');
}

function getWaspTenantName_() {
  var match = getWaspTenantOrigin_().match(/^https?:\/\/([^./]+)/i);
  return match ? String(match[1] || '').toLowerCase() : '';
}

function getStoredWaspToken_() {
  var props = PropertiesService.getScriptProperties();
  return String(props.getProperty('WASP_API_TOKEN') || props.getProperty('WASP_TOKEN') || '').trim();
}

function getStoredWaspTokenSource_() {
  var props = PropertiesService.getScriptProperties();
  if (props.getProperty('WASP_API_TOKEN')) return 'WASP_API_TOKEN';
  if (props.getProperty('WASP_TOKEN')) return 'WASP_TOKEN';
  return '';
}

function waspLooksLikeHtml_(body) {
  return String(body || '').trim().charAt(0) === '<';
}

function waspShouldRetryCode_(code) {
  return code === 429 || code === 502 || code === 503 || code === 504;
}

function waspRetryDelayMs_(attempt, headers) {
  var resetSec = parseInt((headers && (headers['wasp-overloadprotection-reset'] || headers['Wasp-OverloadProtection-Reset'])) || '', 10);
  if (!isNaN(resetSec) && resetSec > 0) return resetSec * 1000;
  var delays = [2000, 5000, 10000, 15000];
  return delays[Math.min(Math.max(attempt - 1, 0), delays.length - 1)];
}

function waspTotalPages_(totalCount, pageSize, fallbackPages) {
  var total = parseInt(totalCount || 0, 10);
  var size = parseInt(pageSize || 500, 10);
  var fallback = parseInt(fallbackPages || 200, 10);
  if (total > 0 && size > 0) {
    return Math.min(Math.max(Math.ceil(total / size), 1), 1000);
  }
  return Math.min(Math.max(fallback, 1), 1000);
}

function waspBodyHasError_(body) {
  var text = String(body || '');
  return /"HasError"\s*:\s*true/i.test(text) || text.indexOf('HasError\\":true') >= 0;
}

function getWaspTokenInfo_() {
  var rawToken = getStoredWaspToken_();
  var info = {
    present: !!rawToken,
    rawLength: rawToken.length,
    propertyName: getStoredWaspTokenSource_(),
    clientId: '',
    email: '',
    issued: '',
    expires: '',
    isExpired: false,
    roles: '',
    decodeError: ''
  };

  if (!rawToken) return info;

  try {
    var decoded = Utilities.newBlob(Utilities.base64Decode(rawToken)).getDataAsString();
    var parts = decoded.split('&');
    for (var i = 0; i < parts.length; i++) {
      var part = parts[i];
      var eq = part.indexOf('=');
      if (eq < 0) continue;
      var key = part.substring(0, eq);
      var value = part.substring(eq + 1);

      if (key === 'client_id') info.clientId = value;
      else if (key === '.issued') info.issued = value;
      else if (key === '.expires') info.expires = value;
      else if (key === 'roles') info.roles = value;
      else if (key === 'http://schemas.xmlsoap.org/ws/2005/05/identity/claims/name') info.email = value;
    }

    if (info.expires) {
      var expiryDate = new Date(info.expires);
      if (!isNaN(expiryDate.getTime())) info.isExpired = expiryDate.getTime() <= Date.now();
    }
  } catch (e) {
    info.decodeError = e.message;
  }

  return info;
}

function getWaspInventorySearchCandidates_() {
  return ['ic/item/advancedinventorysearch', 'ic/item/inventorysearch'];
}

function getWaspItemSearchCandidates_() {
  // The older search/itemsearch/detailsearch endpoints have consistently 404'd in this project.
  // Keep the candidate list tight so Sync and health checks do not waste minutes probing dead paths.
  return ['ic/item/infosearch', 'ic/item/advancedinfosearch'];
}

function waspFetchFirstSuccessful_(endpoints, payload) {
  var errors = [];
  for (var i = 0; i < endpoints.length; i++) {
    try {
      return {
        endpoint: endpoints[i],
        result: waspFetch(endpoints[i], payload)
      };
    } catch (e) {
      errors.push(endpoints[i] + ' - ' + e.message);
    }
  }
  throw new Error(errors.join(' | '));
}

function waspProbeOnce_(endpoint, payload) {
  var token = getStoredWaspToken_();
  if (!token) throw new Error('Wasp API token not set.');

  var response = UrlFetchApp.fetch(getWaspBase() + endpoint, {
    method: 'post',
    headers: { 'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json', 'Accept': 'application/json' },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  var code = response.getResponseCode();
  var body = response.getContentText();

  if (code === 200 && !waspLooksLikeHtml_(body)) {
    try {
      return JSON.parse(body);
    } catch (parseErr) {
      throw new Error('Wasp returned invalid JSON on ' + endpoint + ': ' + parseErr.message);
    }
  }

  if (code === 200 && waspLooksLikeHtml_(body)) {
    throw new Error('Wasp returned HTML on ' + endpoint + ': ' + waspDescribeHtmlResponse_(body));
  }

  throw new Error('Wasp ' + code + ' on ' + endpoint + ': ' + body.substring(0, 300));
}

function waspProbeFirstSuccessful_(endpoints, payload) {
  var errors = [];
  for (var i = 0; i < endpoints.length; i++) {
    try {
      return {
        endpoint: endpoints[i],
        result: waspProbeOnce_(endpoints[i], payload)
      };
    } catch (e) {
      errors.push(endpoints[i] + ' - ' + e.message);
    }
  }
  throw new Error(errors.join(' | '));
}

function waspExtractHtmlTitle_(body) {
  var match = String(body || '').match(/<title[^>]*>([\s\S]*?)<\/title>/i);
  if (!match) return '';
  return String(match[1] || '').replace(/\s+/g, ' ').trim();
}

function waspCompactSnippet_(body, maxLen) {
  var text = String(body || '').replace(/\s+/g, ' ').trim();
  return text.substring(0, maxLen || 240);
}

function waspDescribeHtmlResponse_(body) {
  var title = waspExtractHtmlTitle_(body);
  var snippet = waspCompactSnippet_(body, 180).toLowerCase();

  if (title === 'Wasp Inventory') {
    return 'the Wasp web app shell was returned instead of a JSON API response';
  }
  if (snippet.indexOf('resource not found') >= 0) {
    return 'an HTML "Resource not found" page was returned instead of JSON';
  }
  if (snippet.indexOf('login') >= 0 || snippet.indexOf('sign in') >= 0) {
    return 'an HTML login page was returned instead of JSON';
  }
  if (title) {
    return 'HTML page "' + title + '" was returned instead of JSON';
  }
  return 'HTML was returned instead of JSON';
}

function waspInspectRequest_(url, method, payload, includeAuth) {
  var httpMethod = String(method || 'get').toLowerCase();
  var headers = { 'Accept': 'application/json' };

  if (includeAuth) {
    var token = getStoredWaspToken_();
    if (!token) throw new Error('Wasp API token not set.');
    headers.Authorization = 'Bearer ' + token;
  }

  if (httpMethod !== 'get') headers['Content-Type'] = 'application/json';

  var options = {
    method: httpMethod,
    headers: headers,
    muteHttpExceptions: true
  };

  if (httpMethod !== 'get' && payload !== undefined && payload !== null) {
    options.payload = JSON.stringify(payload);
  }

  var response = UrlFetchApp.fetch(url, options);
  var body = response.getContentText();
  var isHtml = waspLooksLikeHtml_(body);
  var parsed = null;
  var isJson = false;
  var hasError = false;
  var message = '';
  var dataCount = '';
  var totalRecords = '';

  if (!isHtml) {
    try {
      parsed = JSON.parse(body);
      isJson = true;
      hasError = parsed && parsed.HasError === true;
      if (parsed && parsed.Message !== null && parsed.Message !== undefined && parsed.Message !== '') {
        message = String(parsed.Message);
      }
      if (parsed && parsed.TotalRecordsLongCount !== undefined && parsed.TotalRecordsLongCount !== null && parsed.TotalRecordsLongCount !== '') {
        totalRecords = parsed.TotalRecordsLongCount;
      } else if (parsed && parsed.TotalRecordsCount !== undefined && parsed.TotalRecordsCount !== null && parsed.TotalRecordsCount !== '') {
        totalRecords = parsed.TotalRecordsCount;
      }
      if (parsed && parsed.Data && parsed.Data.length !== undefined) dataCount = parsed.Data.length;
    } catch (e) {}
  }

  return {
    url: url,
    code: response.getResponseCode(),
    isHtml: isHtml,
    isJson: isJson,
    hasError: hasError,
    title: isHtml ? waspExtractHtmlTitle_(body) : '',
    snippet: waspCompactSnippet_(body, 260),
    message: message,
    dataCount: dataCount,
    totalRecords: totalRecords
  };
}

function waspInspectEndpoint_(endpoint, payload) {
  var probe = waspInspectRequest_(getWaspBase() + endpoint, 'post', payload || {}, true);
  probe.endpoint = endpoint;
  return probe;
}

function waspProbeLooksSuccessful_(probe) {
  return !!probe && !probe.error && probe.code === 200 && probe.isJson && !probe.hasError;
}

function waspProbesMatchHtmlShell_(left, right) {
  if (!left || !right) return false;
  if (!left.isHtml || !right.isHtml) return false;

  var leftTitle = String(left.title || '').trim();
  var rightTitle = String(right.title || '').trim();
  if (!leftTitle || !rightTitle || leftTitle !== rightTitle) return false;

  var leftSnippet = String(left.snippet || '').replace(/\s+/g, ' ').toLowerCase().substring(0, 120);
  var rightSnippet = String(right.snippet || '').replace(/\s+/g, ' ').toLowerCase().substring(0, 120);
  return leftSnippet === rightSnippet;
}

function waspFetch(endpoint, payload) {
  var token = getStoredWaspToken_();
  if (!token) throw new Error('Wasp API token not set.');

  var options = {
    method: 'post',
    headers: { 'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json', 'Accept': 'application/json' },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  var lastError = '';

  for (var attempt = 1; attempt <= 4; attempt++) {
    var response;
    try {
      response = UrlFetchApp.fetch(getWaspBase() + endpoint, options);
    } catch (fetchErr) {
      lastError = 'Wasp fetch exception on ' + endpoint + ': ' + fetchErr.message;
      if (attempt === 4) throw new Error(lastError);
      Utilities.sleep(waspRetryDelayMs_(attempt, null));
      continue;
    }

    var code = response.getResponseCode();
    var headers = response.getHeaders();
    var body = response.getContentText();

    if (code === 200 && !waspLooksLikeHtml_(body)) {
      try {
        return JSON.parse(body);
      } catch (parseErr) {
        lastError = 'Wasp returned invalid JSON on ' + endpoint + ': ' + parseErr.message;
        if (attempt === 4) throw new Error(lastError);
        Utilities.sleep(waspRetryDelayMs_(attempt, headers));
        continue;
      }
    }

    if (code === 200 && waspLooksLikeHtml_(body)) {
      lastError = 'Wasp returned HTML on ' + endpoint + ' — service temporarily unavailable';
      lastError = 'Wasp returned HTML on ' + endpoint + ': ' + waspDescribeHtmlResponse_(body);
    } else if (waspShouldRetryCode_(code)) {
      lastError = 'Wasp ' + code + ' on ' + endpoint + ': ' + body.substring(0, 300);
    } else {
      throw new Error('Wasp ' + code + ' on ' + endpoint + ': ' + body.substring(0, 300));
    }

    if (attempt === 4) throw new Error(lastError);
    Logger.log('waspFetch retry ' + attempt + ' on ' + endpoint + ': ' + lastError);
    Utilities.sleep(waspRetryDelayMs_(attempt, headers));
  }

  throw new Error(lastError || ('Wasp request failed on ' + endpoint));
}

function waspApiCall(url, payload) {
  var token = getStoredWaspToken_();
  var payloadString = typeof payload === 'string' ? payload : JSON.stringify(payload);
  var options = {
    method: 'POST',
    headers: { 'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json', 'Accept': 'application/json' },
    payload: payloadString,
    muteHttpExceptions: true
  };
  var lastCode = 0;
  var lastBody = '';

  for (var attempt = 1; attempt <= 4; attempt++) {
    try {
      var response = UrlFetchApp.fetch(url, options);
      var code = response.getResponseCode();
      var headers = response.getHeaders();
      var body = response.getContentText();
      lastCode = code;
      lastBody = body;

      if (code === 200 && !waspLooksLikeHtml_(body) && !waspBodyHasError_(body)) {
        return { success: true, code: code, response: body };
      }

      if (code === 200 && waspLooksLikeHtml_(body)) {
        Logger.log('waspApiCall HTML retry ' + attempt + ' url=' + url);
      } else if (waspShouldRetryCode_(code)) {
        Logger.log('waspApiCall transient retry ' + attempt + ': code=' + code + ' url=' + url);
      } else {
        Logger.log('waspApiCall FAILED: code=' + code + ' url=' + url + ' body=' + body.substring(0, 500));
        return { success: false, code: code, response: body };
      }

      if (attempt === 4) break;
      Utilities.sleep(waspRetryDelayMs_(attempt, headers));
    } catch (error) {
      lastCode = 0;
      lastBody = error.message;
      if (attempt === 4) break;
      Utilities.sleep(waspRetryDelayMs_(attempt, null));
    }
  }

  Logger.log('waspApiCall FAILED after retries: code=' + lastCode + ' url=' + url + ' body=' + String(lastBody).substring(0, 500));
  return { success: false, code: lastCode, response: lastBody };
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
  var firstPage = waspFetchFirstSuccessful_(getWaspInventorySearchCandidates_(), { PageSize: 500, PageNumber: 1 });
  var endpoint = firstPage.endpoint;
  var result = firstPage.result;
  var data = result.Data || [];
  var totalCount = result.TotalRecordsLongCount || 0;
  var hasMore = result.HasSuccessWithMoreDataRemaining === true;
  var page = 2;
  var maxPages = waspTotalPages_(totalCount, 500, 200);

  Logger.log('Using WASP inventory endpoint: ' + endpoint);
  for (var d = 0; d < data.length; d++) { allItems.push(data[d]); }

  while (hasMore && page <= maxPages) {
    var payload = { PageSize: 500, PageNumber: page };
    if (totalCount > 0) payload.TotalCountFromPriorFetch = totalCount;
    result = waspFetch(endpoint, payload);
    data = result.Data || [];
    for (var d2 = 0; d2 < data.length; d2++) { allItems.push(data[d2]); }
    hasMore = result.HasSuccessWithMoreDataRemaining === true;
    page++;
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
      var lotMaxPages = waspTotalPages_(lotTotal, 500, 200);

      while (hasMore && lotPage <= lotMaxPages) {
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
  Logger.log('Dedicated lot endpoints failed - falling back to inventory search client-side lot filter');
  try {
    var fallbackItems = [];
    var inventoryPage = waspFetchFirstSuccessful_(getWaspInventorySearchCandidates_(), { PageSize: 500, PageNumber: 1 });
    var inventoryEndpoint = inventoryPage.endpoint;
    var fbResult = inventoryPage.result;
    var fbData = fbResult.Data || [];
    var fbTotal = fbResult.TotalRecordsLongCount || 0;
    var fbHasMore = fbResult.HasSuccessWithMoreDataRemaining === true;
    var fbPage = 2;
    var fbMaxPages = waspTotalPages_(fbTotal, 500, 200);

    for (var fd = 0; fd < fbData.length; fd++) { fallbackItems.push(fbData[fd]); }

    while (fbHasMore && fbPage <= fbMaxPages) {
      var fbPayload = { PageSize: 500, PageNumber: fbPage };
      if (fbTotal > 0) fbPayload.TotalCountFromPriorFetch = fbTotal;
      fbResult = waspFetch(inventoryEndpoint, fbPayload);
      fbData = fbResult.Data || [];
      for (var fd2 = 0; fd2 < fbData.length; fd2++) { fallbackItems.push(fbData[fd2]); }
      fbHasMore = fbResult.HasSuccessWithMoreDataRemaining === true;
      fbPage++;
    }

    // Filter client-side: keep only records that have a non-empty lot field
    var withLots = [];
    for (var fl = 0; fl < fallbackItems.length; fl++) {
      var flItem = fallbackItems[fl];
      var lot = flItem.Lot || flItem.LotNumber || flItem.LotCode || '';
      if (lot) withLots.push(flItem);
    }

    Logger.log('inventory fallback via ' + inventoryEndpoint + ': ' + fallbackItems.length + ' total, ' + withLots.length + ' with lot data');
    return mapLotItems(withLots);

  } catch (e) {
    Logger.log('Inventory-search fallback also failed: ' + e.message);
  }

  Logger.log('No working lot endpoint found - returning empty array');
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
function fetchWaspItemLotSettingsFresh_() {
  var map = {};
  var endpoints = getWaspItemSearchCandidates_();

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
      var maxPages = waspTotalPages_(totalCount, 500, 100);

      while (hasMore && page <= maxPages) {
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
 * Fetches per-item lot tracking settings from the WASP item catalog.
 * Uses a short-lived cache because these settings rarely change, but the
 * underlying item search is one of the slowest parts of Sync Now.
 *
 * @param {boolean} forceRefresh - bypass cache for explicit audits/reference runs
 * @returns {Object} Map of SKU -> true/false for lot tracking
 */
function fetchWaspItemLotSettings(forceRefresh) {
  return fetchWaspCached_(
    WASP_LOT_SETTINGS_CACHE_KEY,
    WASP_LOT_SETTINGS_CACHE_TTL_MS,
    fetchWaspItemLotSettingsFresh_,
    forceRefresh === true
  );
}

/**
 * Pre-mark a WASP adjustment in the GAS webhook's script cache so F2 skips SA creation.
 * Called before every WASP add/remove API call.
 * Requires ScriptProperty 'GAS_WEBHOOK_URL' to be set.
 */
function enginPreMark_(itemNumber, locationCode, op) {
  var url = PropertiesService.getScriptProperties().getProperty('GAS_WEBHOOK_URL');
  if (!url) return;
  var payload = JSON.stringify({ action: 'engin_mark', sku: itemNumber, location: locationCode, op: op });
  var maxAttempts = 2;
  for (var attempt = 1; attempt <= maxAttempts; attempt++) {
    try {
      var resp = UrlFetchApp.fetch(url, {
        method: 'POST',
        contentType: 'application/json',
        payload: payload,
        muteHttpExceptions: true
      });
      var code = resp.getResponseCode();
      if (code === 200) {
        Logger.log('enginPreMark_ ' + op + ' ' + itemNumber + '@' + locationCode + ' → OK (attempt ' + attempt + ')');
        return;
      }
      Logger.log('enginPreMark_ ' + op + ' ' + itemNumber + '@' + locationCode + ' → HTTP ' + code + ' (attempt ' + attempt + ')');
    } catch (e) {
      Logger.log('enginPreMark_ error (attempt ' + attempt + '): ' + e.message);
    }
    if (attempt < maxAttempts) Utilities.sleep(500);
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

/**
 * Transfer inventory with lot from one site/location to another.
 * Implemented as remove-from-source then add-to-destination (WASP has no native transfer API).
 * WARNING: if the add step fails after remove succeeds, stock will be missing — check logs.
 */
function waspTransferWithLot(itemNumber, quantity, lot, dateCode, fromSite, fromLocation, toSite, toLocation, notes) {
  var noteStr = notes || ('Transfer ' + fromSite + '/' + fromLocation + ' → ' + toSite + '/' + toLocation);
  var removeResult = waspRemoveInventoryWithLot(itemNumber, quantity, fromLocation, lot, noteStr, fromSite, dateCode);
  if (!removeResult || removeResult.HasError) {
    throw new Error('Transfer remove failed: ' + JSON.stringify(removeResult));
  }
  var addResult = waspAddInventoryWithLot(itemNumber, quantity, toLocation, lot, dateCode, noteStr, toSite);
  if (!addResult || addResult.HasError) {
    throw new Error('Transfer add failed (remove already done — check WASP): ' + JSON.stringify(addResult));
  }
  return { removeResult: removeResult, addResult: addResult };
}

/**
 * Transfer inventory (no lot) from one site/location to another.
 */
function waspTransfer(itemNumber, quantity, fromSite, fromLocation, toSite, toLocation, notes) {
  var noteStr = notes || ('Transfer ' + fromSite + '/' + fromLocation + ' → ' + toSite + '/' + toLocation);
  var removeResult = waspRemoveInventory(itemNumber, quantity, fromLocation, noteStr, fromSite);
  if (!removeResult || removeResult.HasError) {
    throw new Error('Transfer remove failed: ' + JSON.stringify(removeResult));
  }
  var addResult = waspAddInventory(itemNumber, quantity, toLocation, noteStr, toSite);
  if (!addResult || addResult.HasError) {
    throw new Error('Transfer add failed (remove already done — check WASP): ' + JSON.stringify(addResult));
  }
  return { removeResult: removeResult, addResult: addResult };
}
