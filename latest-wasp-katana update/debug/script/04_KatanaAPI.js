// ============================================
// 04_KatanaAPI.gs - KATANA API FUNCTIONS
// ============================================
// All Katana MRP API interactions
// UPDATED: Added batch_transactions support for batch-tracked items
// FIXED: Changed const to var for Google Apps Script compatibility
// FIXED: Improved batch lookup with case-insensitive matching
// ============================================

/**
 * Generic Katana API call (GET) with retry.
 * Retries up to 2 times (3 total) on 5xx, 429, or network errors.
 * Delays: 3s before retry 2, 6s before retry 3.
 * Client errors (4xx except 429) are not retried — they indicate bad requests.
 * Only logs after all attempts are exhausted.
 */
function katanaApiCall(endpoint) {
  var url = CONFIG.KATANA_BASE_URL + '/' + endpoint;

  var options = {
    method: 'GET',
    headers: {
      'Authorization': 'Bearer ' + CONFIG.KATANA_API_KEY,
      'Accept': 'application/json'
    },
    muteHttpExceptions: true
  };

  var lastCode = 0;
  var lastBody = '';

  for (var attempt = 0; attempt < 4; attempt++) {
    if (attempt > 0) {
      // On 429: exponential backoff — 10s, 20s, 30s
      // On other errors: shorter wait — 2s, 4s, 6s
      var wait = (lastCode === 429) ? (attempt * 10000) : (attempt * 2000);
      Utilities.sleep(wait);
    }

    try {
      var response = UrlFetchApp.fetch(url, options);
      lastCode = response.getResponseCode();
      lastBody = response.getContentText();

      if (lastCode === 200) return JSON.parse(lastBody);

      // Client errors (4xx except 429 rate-limit) — don't retry
      if (lastCode >= 400 && lastCode < 500 && lastCode !== 429) break;

      // 5xx or 429 — fall through to retry

    } catch (err) {
      lastBody = err.message;
      // Network / timeout error — retry
    }
  }

  logToSheet('KATANA_API_ERROR', { endpoint: endpoint, statusCode: lastCode }, lastBody);
  return null;
}

// ============================================
// FETCH FUNCTIONS
// ============================================

function fetchKatanaPO(poId) {
  return katanaApiCall('purchase_orders/' + poId);
}

function fetchKatanaPORows(poId) {
  return katanaApiCall('purchase_order_rows?purchase_order_id=' + poId + '&include=batch_transactions');
}

function fetchKatanaPORow(poRowId) {
  return katanaApiCall('purchase_order_rows/' + poRowId + '?include=batch_transactions');
}

function fetchKatanaVariant(variantId) {
  return katanaApiCall('variants/' + variantId);
}

function fetchKatanaProduct(productId) {
  return katanaApiCall('products/' + productId);
}

function fetchKatanaMaterial(materialId) {
  return katanaApiCall('materials/' + materialId);
}

function fetchKatanaSalesOrder(soId) {
  return katanaApiCall('sales_orders/' + soId);
}

function fetchKatanaCustomer(customerId) {
  if (!customerId) return null;
  return katanaApiCall('customers/' + customerId);
}

function fetchKatanaSalesOrderFull(soId) {
  return katanaApiCall('sales_orders/' + soId);
}

function fetchKatanaSalesOrderRows(soId) {
  return katanaApiCall('sales_order_rows?sales_order_id=' + soId);
}

function fetchKatanaSalesOrderAddresses(soId) {
  if (!soId) return null;
  return katanaApiCall('sales_order_addresses?sales_order_id=' + soId);
}

/**
 * Fetch a single batch stock record by ID.
 * Used to resolve batch_id → batch_number from manufacturing_order_recipe_rows batch_transactions.
 * Handles depleted batches: tries direct lookup first, then include_deleted=true,
 * then query by batch_id parameter (Katana removes zero-stock batches from default results).
 * Returns { batch_number, expiration_date, ... } or null.
 */
function fetchKatanaBatchStock(batchId) {
  if (!batchId) return null;

  // Try 1: Direct lookup by ID
  var result = katanaApiCall('batch_stocks/' + batchId);
  if (result) {
    return result.data ? result.data : result;
  }

  // Try 2: Direct lookup with include_deleted (batch may be depleted/soft-deleted)
  result = katanaApiCall('batch_stocks/' + batchId + '?include_deleted=true');
  if (result) {
    return result.data ? result.data : result;
  }

  // Try 3: Query by batch_id parameter with include_deleted.
  // Katana may not filter by batch_id — scan results to verify ID match.
  // Returning batches[0] unverified caused wrong-batch bugs (e.g. UFC409A returned for UFC413A lookup).
  result = katanaApiCall('batch_stocks?batch_id=' + batchId + '&include_deleted=true');
  if (result) {
    var batches = result.data || result || [];
    for (var t3i = 0; t3i < batches.length; t3i++) {
      var t3b = batches[t3i];
      var t3Id = t3b.batch_id || t3b.id;
      if (String(t3Id) === String(batchId)) return t3b;
    }
  }

  return null;
}

function fetchKatanaMO(moId) {
  return katanaApiCall('manufacturing_orders/' + moId);
}

function fetchKatanaMOIngredients(moId) {
  // Include batch_transactions to get lot/batch allocation data for batch-tracked ingredients
  return katanaApiCall('manufacturing_order_recipe_rows?manufacturing_order_id=' + moId + '&include=batch_transactions');
}

function fetchKatanaObjectFromCollection_(collectionEndpoint, objectId, maxPages) {
  if (!objectId) return null;
  var pages = Math.max(1, Number(maxPages || 5));
  for (var page = 1; page <= pages; page++) {
    var result = katanaApiCall(collectionEndpoint + '?per_page=100&page=' + page + '&sort=-updated_at');
    var rows = result && result.data ? result.data : (result || []);
    if (!rows || rows.length === 0) break;
    for (var i = 0; i < rows.length; i++) {
      if (String(rows[i].id || '') === String(objectId)) return rows[i];
    }
  }
  return null;
}

function fetchKatanaStockAdjustment(saId) {
  return fetchKatanaObjectFromCollection_('stock_adjustments', saId, 6);
}

function fetchKatanaStockAdjustmentRows(saId) {
  var sa = fetchKatanaStockAdjustment(saId);
  var rows = sa ? (sa.stock_adjustment_rows || sa.rows || []) : [];
  return { data: rows };
}

function fetchKatanaStockTransfer(stId) {
  return fetchKatanaObjectFromCollection_('stock_transfers', stId, 6);
}

// ============================================
// LOOKUP FUNCTIONS
// ============================================

/**
 * Get Katana variant by SKU
 */
function getKatanaVariantBySku(sku) {
  var url = CONFIG.KATANA_BASE_URL + '/variants?sku=' + encodeURIComponent(sku);

  var options = {
    method: 'GET',
    headers: {
      'Authorization': 'Bearer ' + CONFIG.KATANA_API_KEY,
      'Content-Type': 'application/json'
    },
    muteHttpExceptions: true
  };

  try {
    var response = UrlFetchApp.fetch(url, options);
    var result = JSON.parse(response.getContentText());

    if (result && result.data && result.data.length > 0) {
      return result.data[0];
    }
    return null;
  } catch (error) {
    logToSheet('KATANA_VARIANT_LOOKUP_ERROR', {
      sku: sku,
      error: error.message
    }, {});
    return null;
  }
}

/**
 * Get the average cost for a variant at a specific Katana location.
 * Returns 0 if the variant has no inventory or the field is missing.
 * Used as cost_per_unit fallback for stock ADD adjustments when the
 * variant has no purchase_price configured.
 */
function getKatanaAverageCost_(variantId, locationId) {
  try {
    var url = CONFIG.KATANA_BASE_URL + '/inventory?variant_id=' + variantId;
    var options = {
      method: 'GET',
      headers: {
        'Authorization': 'Bearer ' + CONFIG.KATANA_API_KEY,
        'Content-Type': 'application/json'
      },
      muteHttpExceptions: true
    };
    var response = UrlFetchApp.fetch(url, options);
    var result = JSON.parse(response.getContentText());
    var rows = result.data || result || [];
    // Prefer the row matching the specific location; fall back to any row
    var best = 0;
    for (var i = 0; i < rows.length; i++) {
      var avg = parseFloat(rows[i].average_cost || rows[i].cost || 0);
      if (rows[i].location_id === locationId && avg > 0) return avg;
      if (avg > 0 && best === 0) best = avg;
    }
    return best;
  } catch (e) {
    return 0;
  }
}

/**
 * Get Katana location by name
 */
function getKatanaLocationByName(locationName) {
  var url = CONFIG.KATANA_BASE_URL + '/locations?name=' + encodeURIComponent(locationName);

  var options = {
    method: 'GET',
    headers: {
      'Authorization': 'Bearer ' + CONFIG.KATANA_API_KEY,
      'Content-Type': 'application/json'
    },
    muteHttpExceptions: true
  };

  try {
    var response = UrlFetchApp.fetch(url, options);
    var result = JSON.parse(response.getContentText());

    if (result && result.data && result.data.length > 0) {
      return result.data[0];
    }
    return null;
  } catch (error) {
    logToSheet('KATANA_LOCATION_LOOKUP_ERROR', {
      name: locationName,
      error: error.message
    }, {});
    return null;
  }
}

/**
 * Find batch ID by lot number for a variant
 * Required for batch-tracked items in Katana
 * IMPROVED: Case-insensitive matching and better field checking
 */
function findBatchIdByNumber(variantId, lotNumber) {
  if (!lotNumber) return null;

  var url = CONFIG.KATANA_BASE_URL + '/batch_stocks?variant_id=' + variantId;

  var options = {
    method: 'GET',
    headers: {
      'Authorization': 'Bearer ' + CONFIG.KATANA_API_KEY,
      'Accept': 'application/json'
    },
    muteHttpExceptions: true
  };

  try {
    var response = UrlFetchApp.fetch(url, options);
    var code = response.getResponseCode();

    if (code !== 200) {
      logToSheet('BATCH_LOOKUP_API_ERROR', { variantId: variantId, lot: lotNumber, code: code }, {});
      return null;
    }

    var result = JSON.parse(response.getContentText());
    var batches = result.data || result || [];
    var lotUpper = lotNumber.toUpperCase();

    // Check all batches for a match
    for (var i = 0; i < batches.length; i++) {
      var batch = batches[i];

      // Try different possible field names
      var batchNumber = batch.batch_number || batch.batch_nr || batch.nr || batch.number || '';
      var batchNumberUpper = String(batchNumber).toUpperCase();

      // Case-insensitive match
      if (batchNumberUpper === lotUpper) {
        var foundBatchId = batch.batch_id || batch.id;
        var foundQty = parseFloat(batch.in_stock || batch.quantity || batch.quantity_in_stock || 0);
        var foundExpiry = normalizeBusinessDate_(batch.expiration_date || batch.expiry_date || batch.best_before_date || '');
        return {
          id: foundBatchId,
          qty: foundQty,
          lot: batchNumber || '',
          expiry: foundExpiry
        };
      }
    }

    // Only log if batch not found (important for debugging)
    logToSheet('BATCH_NOT_FOUND', {
      variantId: variantId,
      lot: lotNumber,
      batchesChecked: batches.length
    }, {});
    return null;
  } catch (error) {
    logToSheet('BATCH_LOOKUP_ERROR', { variantId: variantId, lot: lotNumber, error: error.message }, {});
    return null;
  }
}

// fetchKatanaLocation() lives in 07_F3_AutoPoll.gs (with caching)

// ============================================
// ACTION FUNCTIONS
// ============================================

/**
 * Create stock adjustment in Katana
 * UPDATED: Now supports batch_transactions for batch-tracked items
 *
 * @param {number} variantId - Katana variant ID
 * @param {number} quantity - Positive=add, negative=remove
 * @param {number} locationId - Katana location ID
 * @param {string} notes - Notes for logging
 * @param {string} lotNumber - Optional lot number from WASP
 * @param {string} expiryDate - Optional expiry date from WASP
 */
function createKatanaStockAdjustment(variantId, quantity, locationId, notes, lotNumber, expiryDate) {
  var url = CONFIG.KATANA_BASE_URL + '/stock_adjustments';

  // Simple SA number
  var adjNumber = 'WASP - Imported';

  // Build the SA row
  var saRow = {
    variant_id: variantId,
    quantity: quantity
  };

  // Check if item has batch tracking - look up existing batch
  var batchId = null;
  if (lotNumber) {
    var batchResult = findBatchIdByNumber(variantId, lotNumber);
    batchId = batchResult ? batchResult.id : null;
    if (batchId) {
      // Add batch_transactions for batch-tracked items
      saRow.batch_transactions = [{
        batch_id: batchId,
        quantity: quantity
      }];

      logToSheet('KATANA_SA_BATCH_ADDED', {
        variantId: variantId,
        lotNumber: lotNumber,
        batchId: batchId,
        quantity: quantity
      }, {});
    } else {
      logToSheet('KATANA_SA_NO_BATCH_MATCH', {
        variantId: variantId,
        lotNumber: lotNumber,
        message: 'Lot number not found in Katana batches - will fail if batch-tracked'
      }, {});
    }
  }

  var payload = {
    stock_adjustment_number: adjNumber,
    stock_adjustment_date: new Date().toISOString(),
    location_id: locationId,
    reason: 'WASP Imported',
    additional_info: notes || '',
    stock_adjustment_rows: [saRow]
  };

  var options = {
    method: 'POST',
    headers: {
      'Authorization': 'Bearer ' + CONFIG.KATANA_API_KEY,
      'Content-Type': 'application/json'
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  // Logging moved to response only

  try {
    var response = UrlFetchApp.fetch(url, options);
    var code = response.getResponseCode();
    var text = response.getContentText();

    // Only log errors, success is logged in Activity sheet
    if (code < 200 || code >= 300) {
      logToSheet('KATANA_SA_ERROR', {
        statusCode: code,
        variantId: variantId
      }, text);
    }

    return {
      success: code >= 200 && code < 300,
      code: code,
      response: text,
      adjNumber: adjNumber,
      lot: lotNumber || '',
      expiry: expiryDate || '',
      batchId: batchId
    };
  } catch (error) {
    logToSheet('KATANA_ADJUSTMENT_ERROR', { error: error.message }, {});
    return { success: false, error: error.message };
  }
}

/**
 * Create batch stock adjustment in Katana with multiple rows
 * Used for grouping multiple items into one SA
 *
 * @param {Array} saRows - Array of SA row objects with variant_id, quantity, and optional batch_transactions
 * @param {number} locationId - Katana location ID
 * @param {string} additionalInfo - Notes for the SA
 * @param {string} transactionId - Transaction ID for tracking
 */
function createKatanaBatchStockAdjustment(saRows, locationId, additionalInfo, transactionId, saNumber, reason) {
  var url = CONFIG.KATANA_BASE_URL + '/stock_adjustments';

  var payload = {
    stock_adjustment_number: saNumber || 'WASP-ADJ',
    stock_adjustment_date: new Date().toISOString(),
    location_id: locationId,
    reason: reason || 'WASP Adjustment',
    additional_info: '',
    stock_adjustment_rows: saRows
  };

  var options = {
    method: 'POST',
    headers: {
      'Authorization': 'Bearer ' + CONFIG.KATANA_API_KEY,
      'Content-Type': 'application/json'
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  // Logging moved to response only

  try {
    var response = UrlFetchApp.fetch(url, options);
    var code = response.getResponseCode();
    var text = response.getContentText();

    var success = code >= 200 && code < 300;

    // Only log errors, success is logged in Activity sheet
    if (!success) {
      logToSheet('KATANA_SA_ERROR', {
        transactionId: transactionId,
        statusCode: code
      }, { error: text });
    }

    // Extract the numeric SA ID from Katana response (used for direct link)
    var saId = null;
    if (success) {
      try {
        var parsed = JSON.parse(text);
        saId = parsed.id || null;
      } catch (parseErr) {}
    }

    return {
      success: success,
      code: code,
      response: text,
      transactionId: transactionId,
      saId: saId
    };
  } catch (error) {
    logToSheet('KATANA_BATCH_SA_ERROR', {
      transactionId: transactionId,
      error: error.message
    }, {});
    return { success: false, error: error.message };
  }
}

/**
 * Delete a Katana stock adjustment by its numeric ID.
 * @param {number|string} saId - Katana SA numeric ID
 * @returns {{ success: boolean, code: number, error: string }}
 */
function deleteKatanaSA(saId) {
  var url = CONFIG.KATANA_BASE_URL + '/stock_adjustments/' + saId;
  var options = {
    method: 'DELETE',
    headers: {
      'Authorization': 'Bearer ' + CONFIG.KATANA_API_KEY,
      'Content-Type': 'application/json'
    },
    muteHttpExceptions: true
  };
  try {
    var response = UrlFetchApp.fetch(url, options);
    var code = response.getResponseCode();
    var success = code >= 200 && code < 300;
    return { success: success, code: code, error: success ? '' : response.getContentText().substring(0, 200) };
  } catch (e) {
    return { success: false, code: 0, error: e.message };
  }
}

/**
 * Mark Sales Order as delivered in Katana
 */
function markKatanaSalesOrderDelivered(soId) {
  var url = CONFIG.KATANA_BASE_URL + '/sales_orders/' + soId + '/deliver';

  var options = {
    method: 'POST',
    headers: {
      'Authorization': 'Bearer ' + CONFIG.KATANA_API_KEY,
      'Content-Type': 'application/json'
    },
    payload: JSON.stringify({}),
    muteHttpExceptions: true
  };

  try {
    var response = UrlFetchApp.fetch(url, options);
    var code = response.getResponseCode();
    var body = response.getContentText();

    logToSheet('KATANA_DELIVER_RESPONSE', {
      soId: soId,
      statusCode: code
    }, body);

    return {
      success: code >= 200 && code < 300,
      code: code,
      response: body
    };
  } catch (error) {
    logToSheet('KATANA_DELIVER_ERROR', { soId: soId }, { error: error.message });
    return {
      success: false,
      code: 0,
      response: error.message
    };
  }
}
