// ============================================
// 02_Handlers.gs - KATANA EVENT HANDLERS
// ============================================
// F1: PO Receiving
// F3: SO Delivered (Amazon Transfer)
// F4: MO Complete
// F5: SO Created (Order Fulfillment)
// FIXED: Changed const/let to var for Google Apps Script compatibility
// ============================================

// ============================================
// F1: PURCHASE ORDER HANDLERS
// ============================================

/**
 * Handle PO Created event
 * Just logs and notifies - no WASP action
 */
function handlePurchaseOrderCreated(payload) {
  var poId = payload.object ? payload.object.id : null;
  logToSheet('PO_CREATED', { poId: poId }, { status: 'logged' });
  try {
    if (poId) {
      var poData = fetchKatanaPO(poId);
      var poRef = poData && (poData.order_no || ('PO-' + poId));
      var poReceiveState = poRef ? loadPOReceiveState_(poRef) : null;
      var hasExistingOpenSnapshot = !!(poReceiveState && poReceiveState.openRows && Object.keys(poReceiveState.openRows).length > 0);
      var hasExistingReceiveState = !!(poReceiveState && poReceiveState.flat && poReceiveState.flat.length > 0);
      var poStatusText = String((poData && poData.status) || '').toUpperCase();
      if (poRef && !hasExistingOpenSnapshot && !hasExistingReceiveState && poStatusText === 'NOT_RECEIVED') {
        var poRowsData = fetchKatanaPORows(poId);
        if (poRowsData) {
          persistPOOpenRowsSnapshot_(poRef, poRowsData.data || poRowsData || []);
        }
      }
    }
  } catch (e) {
    Logger.log('PO_CREATED open-row snapshot error: ' + e.message);
  }
  sendSlackNotification('PO Created in Katana\nID: ' + poId);
  return { status: 'logged', poId: poId };
}

function handlePurchaseOrderRowReceived(payload) {
  var poRowId = payload && payload.object ? payload.object.id : null;
  if (!poRowId) {
    return { status: 'error', message: 'No PO row ID in webhook' };
  }

  var poRowData = fetchKatanaPORow(poRowId);
  var poRow = poRowData && poRowData.data ? poRowData.data : poRowData;
  if (!poRow) {
    return { status: 'error', message: 'Failed to fetch PO row' };
  }

  var poId = poRow.purchase_order_id || poRow.purchaseOrderId || null;
  if (!poId) {
    return { status: 'error', message: 'No purchase_order_id on PO row' };
  }

  return handlePurchaseOrderReceived({
    action: 'purchase_order_row.received',
    resource_type: 'purchase_order_row',
    object: {
      id: poId,
      status: 'PARTIALLY_RECEIVED',
      href: payload && payload.object ? payload.object.href : ''
    },
    source_row_id: poRowId
  });
}

/**
 * Handle PO Received event
 * Adds received items to WASP RECEIVING-DOCK
 */
function handlePurchaseOrderReceived(payload) {
  var poId = payload.object ? payload.object.id : null;

  if (!poId) {
    logToSheet('PO_RECEIVED_ERROR', payload, { error: 'No PO ID found' });
    return { status: 'error', message: 'No PO ID in webhook' };
  }

  var poStateGuardKey = 'po_state_guard_' + poId;
  var poReceiveGuardWait = waitForExecutionGuardToClear_(poStateGuardKey, [0, 500, 1000, 2000, 4000, 6000, 8000]);
  if (poReceiveGuardWait.waitedMs > 0) {
    logToSheet('PO_RECEIVE_GUARD_WAIT', {
      poId: poId,
      waitedMs: poReceiveGuardWait.waitedMs,
      cleared: poReceiveGuardWait.cleared,
      action: payload && payload.action ? payload.action : ''
    }, 'Waiting for PO state guard before receive processing');
  }
  if (!acquireExecutionGuard_(poStateGuardKey, 120000)) {
    var poReceiveRetryWait = waitForExecutionGuardToClear_(poStateGuardKey, [2000, 4000, 6000]);
    if (poReceiveRetryWait.waitedMs > 0) {
      logToSheet('PO_RECEIVE_GUARD_RETRY_WAIT', {
        poId: poId,
        waitedMs: poReceiveRetryWait.waitedMs,
        cleared: poReceiveRetryWait.cleared,
        action: payload && payload.action ? payload.action : ''
      }, 'Retry waiting for PO state guard before receive processing');
    }
    if (!acquireExecutionGuard_(poStateGuardKey, 120000)) {
      return { status: 'skipped', reason: 'PO receive already in progress', poId: poId };
    }
  }

  markPOReceiveRecent_(poId, 20);

  try {

  // Fetch PO header
  var poData = fetchKatanaPO(poId);
  if (!poData) {
    return { status: 'error', message: 'Failed to fetch PO header' };
  }

  var poNumber = poData.order_no || ('PO-' + poId);
  var payloadAction = String(payload.action || '').toLowerCase();
  var payloadStatus = String((payload.object && payload.object.status) || '').toUpperCase();
  var poReceiveState = loadPOReceiveState_(poNumber);
  var f1PartialKey = poReceiveState.keys.partialKey;
  var hadPriorPartialState = PropertiesService.getScriptProperties().getProperty(f1PartialKey) === 'true';
  var partialReceiveHint = (
    payloadAction === 'purchase_order.partially_received' ||
    payloadAction === 'purchase_order_row.received' ||
    payloadStatus === 'PARTIALLY_RECEIVED' ||
    String(poData.status || '').toUpperCase() === 'PARTIALLY_RECEIVED'
  );
  var receiveContinuationHint = partialReceiveHint || hadPriorPartialState;
  var variantCache = {};
  var locationCache = {};

  // Load any previously stored data for this PO (from earlier partial receives)
  var previouslySaved = poReceiveState.flat || [];
  var existingReceiptBlocks = poReceiveState.blocks || [];
  var previousBatchIds = {};
  for (var psi = 0; psi < previouslySaved.length; psi++) {
    if (previouslySaved[psi].batchId) previousBatchIds[String(previouslySaved[psi].batchId)] = true;
  }

  // Dedup: full receives use 600s + Activity log check.
  // Partial receives dedup after row fetch using an open-row signature so
  // distinct receive blocks on the same PO are not blocked by one shared cache key.
  var cache = CacheService.getScriptCache();
  var poDedupKey = 'po_received_' + poNumber;
  var partialDedupKey = '';
  if (!receiveContinuationHint) {
    if (cache.get(poDedupKey)) {
      return { status: 'skipped', reason: 'PO already processed (cache)', poNumber: poNumber };
    }
    cache.put(poDedupKey, 'true', 600);
    if (isPOAlreadyReceived(poNumber)) {
      logToSheet('PO_RECEIVED_DUPLICATE', { poId: poId, poRef: poNumber }, { reason: 'Already received in Activity log' });
      return { status: 'skipped', reason: 'PO already received', poId: poId, poRef: poNumber };
    }
  }

  // Receive counter — label partial follows as PO-563/2, PO-563/3, etc.
  // Resolve WASP destination from PO "Ship to" location
  var poLocationId = poData.location_id || null;
  var waspDest = { site: CONFIG.WASP_SITE, location: FLOWS.PO_RECEIVING_LOCATION };

  if (poLocationId) {
    var poLocation = locationCache[String(poLocationId)];
    if (poLocation === undefined) {
      poLocation = fetchKatanaLocation(poLocationId);
      locationCache[String(poLocationId)] = poLocation || null;
    }
    var locName = poLocation ? (poLocation.name || '') : '';
    if (locName && KATANA_LOCATION_TO_WASP[locName]) {
      waspDest = KATANA_LOCATION_TO_WASP[locName];
    }
  }

  // Fetch PO rows unless a synthetic receive block was delegated from purchase_order.updated
  var poRowsData = null;
  var openRowsForSnapshot = payload.current_open_rows || null;
  if (payload.override_rows && payload.override_rows.length) {
    poRowsData = { data: payload.override_rows };
  } else {
    poRowsData = fetchKatanaPORows(poId);
    if (!poRowsData) {
      return { status: 'error', message: 'Failed to fetch PO rows' };
    }
    openRowsForSnapshot = poRowsData.data || poRowsData || [];
  }

  if (receiveContinuationHint && !(payload.override_rows && payload.override_rows.length)) {
    var poReadyDelays = payloadAction === 'purchase_order_row.received' ? [0, 500, 1000, 1500] : [0, 1000, 2000, 3000];
    var poReady = waitForPOReceiveReadiness_(poId, poData, poRowsData, previouslySaved, poReadyDelays);
    if (poReady.waitedMs > 0) {
      logToSheet('PO_PARTIAL_WAIT', {
        poId: poId,
        poRef: poNumber,
        waitedMs: poReady.waitedMs,
        ready: poReady.readiness && poReady.readiness.ready,
        receivableCount: poReady.readiness ? poReady.readiness.receivableCount : 0
      }, 'Adaptive wait for Katana partial PO rows before receive processing');
    }
    poData = poReady.poData || poData;
    poRowsData = poReady.rowsData || poRowsData;
    openRowsForSnapshot = poRowsData.data || poRowsData || [];
  }

  var rows = poRowsData.data || poRowsData || [];
  if (receiveContinuationHint) {
    partialDedupKey = poDedupKey + '_partial_' + buildPOOpenRowsSignature_(openRowsForSnapshot || rows);
    if (cache.get(partialDedupKey)) {
      return { status: 'skipped', reason: 'PO already processed (cache)', poNumber: poNumber };
    }
    cache.put(partialDedupKey, 'true', 8);
  }
  var receiveMode = inferPOReceiveMode_(rows, poData.status);
  var isPartialReceive = partialReceiveHint || receiveMode.isPartial;
  var usePartialLifecycle = isPartialReceive || hadPriorPartialState;
  var results = [];
  var previousNonBatchRowIds = {};
  for (var psri = 0; psri < previouslySaved.length; psri++) {
    var prevSavedRow = previouslySaved[psri] || {};
    if (!prevSavedRow.batchId && prevSavedRow.poRowId) {
      previousNonBatchRowIds[String(prevSavedRow.poRowId)] = true;
    }
  }
  var skuCostMap = {}; // sku → cost; populated during loop, used after to update WASP item cost once per SKU

  // Process each row
  for (var i = 0; i < rows.length; i++) {
    var row = rows[i];
    var variantId = row.variant_id;
    var poRowId = row.id || row.purchase_order_row_id || row.po_row_id || null;
    var quantity = row.quantity;
    // Convert purchase UoM → stock UoM (e.g. 1 dozen × 12 = 12 pcs)
    var puomRate = parseFloat(row.purchase_uom_conversion_rate) || 1;
    var puomNote = '';
    if (puomRate !== 1) {
      var rawQty = row.quantity;
      quantity = rawQty * puomRate;
    }

    // Get SKU from variant
    var variant = variantCache[String(variantId)];
    if (variant === undefined) {
      variant = fetchKatanaVariant(variantId);
      variantCache[String(variantId)] = variant || null;
    }
    var sku = variant ? variant.sku : null;

    // Track cost for WASP item cost update (applied once per unique SKU after loop)
    if (sku && variant && !isSkippedSku(sku)) {
      var skuCostVal = parseFloat(variant.purchase_price || variant.default_purchase_price || variant.cost || 0);
      if (skuCostVal > 0) skuCostMap[sku] = { cost: skuCostVal, uom: variant.uom || '' };
    }

    // Build puomNote with full unit names after variant is available
    if (puomRate !== 1) {
      var puomUomDisplay = {
        'g': 'grams', 'G': 'grams',
        'kg': 'kg', 'KG': 'kg',
        'lbs': 'lbs', 'LBS': 'lbs',
        'pcs': 'pcs', 'PCS': 'pcs', 'Pcs': 'pcs', 'PC': 'pcs', 'pc': 'pcs',
        'EA': 'pcs', 'ea': 'pcs',
        'dozen': 'dozen', 'dz': 'dozen'
      };
      // Fallback: infer stock unit from purchase UOM when variant has no UOM set
      var puomToStockFallback = { 'dozen': 'pcs', 'dz': 'pcs', 'box': 'pcs', 'case': 'pcs', 'pack': 'pcs', 'bag': 'pcs', 'bundle': 'pcs' };
      var stockUom = resolveVariantUom(variant) || (row.uom || '');
      var stockUomLabel = puomUomDisplay[stockUom] || stockUom;
      if (!stockUomLabel) {
        stockUomLabel = puomToStockFallback[(row.purchase_uom || '').toLowerCase()] || '';
      }
      puomNote = row.quantity + ' ' + (row.purchase_uom || 'unit') + ' \u2192 ' + quantity;
      if (stockUomLabel) puomNote += ' ' + stockUomLabel;
      Logger.log('F1 UoM convert: ' + row.purchase_uom + ' x' + puomRate + ' = ' + quantity + ' ' + (stockUom || stockUomLabel));
    }

    // Processing PO row

    if (sku && quantity > 0) {
      var poNumber = poData.order_no || ('PO-' + poId);

      // Determine per-item destination: materials/ingredients → PRODUCTION, others → default
      var itemLocation = waspDest.location;
      var itemSite = waspDest.site;
      if (waspDest.site === CONFIG.WASP_SITE) {
        var itemType = getKatanaItemType(variant);
        if (itemType === 'material' || itemType === 'intermediate') {
          itemLocation = LOCATIONS.PRODUCTION;
        }
      }

      // Check for multi-batch (PO row split across 2+ batches)
      var bt = row.batch_transactions || row.batchTransactions || [];

      // On partial receive: for non-batch rows, use received_quantity delta instead of total row qty.
      // This keeps non-lot-tracked partial receives in sync without double-adding prior receipts.
      if (usePartialLifecycle && bt.length === 0) {
        if (poRowId && isKatanaPOReceivedRow_(row)) {
          if (previousNonBatchRowIds[String(poRowId)]) continue;
        } else if (typeof HOTFIX_FLAGS !== 'undefined' && HOTFIX_FLAGS.F1_PARTIAL_NON_BATCH_DELTA && poRowId) {
          var cumulativeReceivedStockQty = getPORowReceivedStockQty_(row, poData.status, { allowPoStatusFallback: false });
          var previouslyStoredQty = 0;
          for (var psq = 0; psq < previouslySaved.length; psq++) {
            var prevRowId = previouslySaved[psq].poRowId || null;
            if (!prevRowId || String(prevRowId) !== String(poRowId)) continue;
            if (previouslySaved[psq].batchId) continue;
            previouslyStoredQty += parseFloat(previouslySaved[psq].qty || 0) || 0;
          }
          var deltaReceivedQty = cumulativeReceivedStockQty - previouslyStoredQty;
          if (deltaReceivedQty <= 0) continue;
          quantity = deltaReceivedQty;
        } else {
          continue;
        }
      }

      if (bt.length > 1) {
        // MULTI-BATCH: Process each batch_transaction separately
        markSyncedToWasp(sku, itemLocation, 'add');

        for (var btIdx = 0; btIdx < bt.length; btIdx++) {
          var btBatchIdCheck = bt[btIdx].batch_id || bt[btIdx].batchId || null;
          if (btBatchIdCheck && previousBatchIds[String(btBatchIdCheck)]) continue;
          var btEntry = bt[btIdx];
          var btQty = btEntry.quantity || 0;
          if (btQty <= 0) continue;

          // Resolve lot and expiry from this batch_transaction
          var btLot = extractKatanaBatchNumber_(btEntry);
          var btExpiry = extractKatanaExpiryDate_(btEntry);

          if (!btLot && btEntry.batch_stock) {
            btLot = extractKatanaBatchNumber_(btEntry.batch_stock);
          }
          var btBatchId = btEntry.batch_id || btEntry.batchId || null;
          if (btBatchId && (!btLot || !btExpiry)) {
            var btBatchInfo = fetchKatanaBatchStock(btBatchId);
            if (btBatchInfo) {
              btLot = btLot || extractKatanaBatchNumber_(btBatchInfo);
              btExpiry = btExpiry || extractKatanaExpiryDate_(btBatchInfo);
            }
          }
          // If still no lot: query all batch_stocks for this variant (recovers depleted batches)
          if (!btLot && btBatchId && variantId) {
            var allVarBatchesMB = katanaApiCall('batch_stocks?variant_id=' + variantId + '&include_deleted=true');
            if (allVarBatchesMB) {
              var vBatchListMB = allVarBatchesMB.data || allVarBatchesMB || [];
              for (var vbMbIdx = 0; vbMbIdx < vBatchListMB.length; vbMbIdx++) {
                var vBatchMB = vBatchListMB[vbMbIdx];
                if (String(vBatchMB.id) === String(btBatchId)) {
                  btLot = extractKatanaBatchNumber_(vBatchMB);
                  if (!btExpiry) btExpiry = extractKatanaExpiryDate_(vBatchMB);
                  break;
                }
              }
            }
          }
          // No fallback lot — send with empty lot; WASP will reject with -57041 if lot is required

          var btResult = waspAddInventoryWithLot(
            sku, btQty, itemLocation, btLot, btExpiry,
            'PO Received: ' + poNumber, itemSite
          );

          results.push({
            sku: sku, quantity: btQty, result: btResult,
            location: itemLocation, site: itemSite, lot: btLot, expiry: btExpiry,
            variantId: variantId, batchId: btBatchId, poRowId: poRowId,
            multiBatch: true, multiBatchTotal: quantity,
            multiBatchFirst: (btIdx === 0), multiBatchLast: (btIdx === bt.length - 1),
            puomNote: puomNote, uom: resolveVariantUom(variant)
          });
        }
      } else {
        // SINGLE-BATCH or NO-BATCH
        // For single-batch: skip if this batch_id was already processed in a prior partial
        var sbBtCheck = row.batch_transactions || row.batchTransactions || [];
        if (sbBtCheck.length === 1) {
          var sbBtIdCheck = sbBtCheck[0].batch_id || sbBtCheck[0].batchId || null;
          if (sbBtIdCheck && previousBatchIds[String(sbBtIdCheck)]) continue;
        }

        var lotNumber = extractIngredientBatchNumber(row);
        var expiryDate = extractIngredientExpiryDate(row) || '';

        // If lot not found via standard lookup, try querying ALL batch_stocks for this variant
        // with include_deleted=true. Depleted batches are removed from batch_stocks/{id} endpoint
        // but still appear in variant-level queries, so this recovers the batch_number.
        if (!lotNumber) {
          var rowBt = row.batch_transactions || row.batchTransactions || [];
          if (rowBt.length > 0 && variantId) {
            var rowBtId = rowBt[0].batch_id || rowBt[0].batchId || null;
            if (rowBtId) {
              var allVarBatches = katanaApiCall('batch_stocks?variant_id=' + variantId + '&include_deleted=true');
              if (allVarBatches) {
                var vBatchList = allVarBatches.data || allVarBatches || [];
                for (var vbIdx = 0; vbIdx < vBatchList.length; vbIdx++) {
                  var vBatch = vBatchList[vbIdx];
                  if (String(vBatch.id) === String(rowBtId)) {
                    lotNumber = extractKatanaBatchNumber_(vBatch);
                    if (!expiryDate) expiryDate = extractKatanaExpiryDate_(vBatch);
                    break;
                  }
                }
              }
            }
          }
        }
        // No fallback lot — send with empty lot; WASP will reject with -57041 if lot is required

        markSyncedToWasp(sku, itemLocation, 'add');

        var result = waspAddInventoryWithLot(
          sku, quantity, itemLocation, lotNumber, expiryDate,
          'PO Received: ' + poNumber, itemSite
        );

        var sbRowBt = row.batch_transactions || row.batchTransactions || [];
        var sbBatchId = sbRowBt.length > 0 ? (sbRowBt[0].batch_id || sbRowBt[0].batchId || null) : null;
        results.push({ sku: sku, quantity: quantity, result: result, location: itemLocation, site: itemSite, lot: lotNumber, expiry: expiryDate, variantId: variantId, batchId: sbBatchId, poRowId: poRowId, puomNote: puomNote, uom: resolveVariantUom(variant) });
      }
    } else {
      // Skipped - missing SKU or zero quantity
    }
  }

  // Update WASP item cost for each unique SKU received — keeps cost in sync with Katana.
  // One call per SKU regardless of how many batches/lots were received.
  for (var costSku in skuCostMap) {
    if (skuCostMap.hasOwnProperty(costSku)) {
      var costEntry = skuCostMap[costSku];
      waspUpdateItemCost(costSku, costEntry.cost, costEntry.uom);
    }
  }

  if (results.length === 0) {
    cache.remove(poDedupKey);
    if (partialDedupKey) cache.remove(partialDedupKey);
    logToSheet('PO_RECEIVE_EMPTY', { poId: poId, poRef: poNumber, partial: isPartialReceive }, {
      reason: isPartialReceive ? 'No new partial-receive rows detected after settle wait' : 'No receivable rows detected'
    });
    return {
      status: 'ignored',
      reason: isPartialReceive ? ('No new partial-receive rows detected for ' + poNumber) : ('No receivable rows detected for ' + poNumber),
      poId: poId,
      poRef: poNumber
    };
  }

  var f1CountKey = 'f1_count_' + poNumber.replace(/[^a-zA-Z0-9_]/g, '');
  var receiveCount = parseInt(PropertiesService.getScriptProperties().getProperty(f1CountKey) || '0', 10) + 1;
  PropertiesService.getScriptProperties().setProperty(f1CountKey, String(receiveCount));
  var poRefLabel = receiveCount > 1 ? (poNumber + '/' + receiveCount) : poNumber;

  var receiptBlock = buildPOReceiptBlockFromResults_(poRefLabel, poId, results, rows, payloadAction);
  if (receiptBlock.items.length > 0) {
    existingReceiptBlocks.push(receiptBlock);
    persistPOReceiveState_(poNumber, existingReceiptBlocks);
  }
  persistPOOpenRowsSnapshot_(poNumber, openRowsForSnapshot || rows);

  // Track partial receive state so isPOAlreadyReceived allows a follow-up receive
  receiveMode = inferPOReceiveMode_(rows, poData.status);
  if (receiveMode.isPartial || String(poData.status || '').toUpperCase() === 'PARTIALLY_RECEIVED') {
    PropertiesService.getScriptProperties().setProperty(f1PartialKey, 'true');
  } else {
    PropertiesService.getScriptProperties().deleteProperty(f1PartialKey);
  }

  // Activity Log — F1 Receiving
  var f1Success = 0;
  var f1Fail = 0;
  var f1SubItems = [];
  for (var r = 0; r < results.length; r++) {
    var res = results[r];
    var itemOk = res.result && res.result.success;
    if (itemOk) f1Success++; else f1Fail++;

    if (res.multiBatch && res.multiBatchFirst) {
      // Multi-batch parent row
      var f1BatchCount = 0;
      for (var f1bc = r; f1bc < results.length; f1bc++) {
        if (results[f1bc].sku === res.sku && results[f1bc].multiBatch) f1BatchCount++;
        else if (f1bc > r) break;
      }
      f1SubItems.push({
        sku: res.sku,
        qty: res.multiBatchTotal,
        uom: normalizeUom(res.uom || ''),
        success: true,
        status: '',
        error: '',
        action: '→ ' + (res.location || waspDest.location) + (res.puomNote ? '  (' + res.puomNote + ')' : ''),
        qtyColor: 'grey',
        isParent: true,
        batchCount: f1BatchCount
      });
      f1SubItems[f1SubItems.length - 1].action = buildActivityCompactMeta_(
        waspDest.site,
        res.location || waspDest.location,
        '',
        '',
        [res.puomNote ? '(' + res.puomNote + ')' : '']
      );
    }

    if (res.multiBatch) {
      // Nested batch sub-row
      var f1NestedAction = '';
      if (res.lot) f1NestedAction += 'lot:' + res.lot;
      if (res.expiry) f1NestedAction += (f1NestedAction ? '  ' : '') + 'exp:' + res.expiry;
      f1SubItems.push({
        sku: '',
        qty: res.quantity,
        uom: normalizeUom(res.uom || ''),
        success: itemOk,
        status: itemOk ? 'Added' : 'Failed',
        error: itemOk ? '' : (res.result ? (res.result.error || parseWaspError(res.result.response, 'Add', res.sku)) : ''),
        action: f1NestedAction,
        qtyColor: 'green',
        nested: true
      });
    } else {
      f1SubItems.push({
        sku: res.sku,
        qty: res.quantity,
        uom: normalizeUom(res.uom || ''),
        success: itemOk,
        status: itemOk ? 'Added' : '',
        error: itemOk ? '' : (res.result ? (res.result.error || parseWaspError(res.result.response, 'Add', res.sku)) : ''),
        action: itemOk ? ('→ ' + (res.location || waspDest.location) + (res.lot ? '  lot:' + res.lot : '') + (res.expiry ? '  exp:' + res.expiry : '') + (res.puomNote ? '  (' + res.puomNote + ')' : '')) : '',
        qtyColor: 'green'
      });
      f1SubItems[f1SubItems.length - 1].action = buildActivityCompactMeta_(
        waspDest.site,
        res.location || waspDest.location,
        res.lot,
        res.expiry,
        [res.puomNote ? '(' + res.puomNote + ')' : '']
      );
    }
  }
  var f1Status = f1Fail === 0 ? 'success' : f1Success === 0 ? 'failed' : 'partial';
  var poRef = extractCanonicalActivityRef_(poData.order_no, 'PO-', poId);

  // Build location summary for logging (handles mixed destinations)
  var f1LocMap = {};
  for (var lr = 0; lr < results.length; lr++) {
    var rl = results[lr].location || waspDest.location;
    f1LocMap[getActivityDisplayLocation_(rl)] = true;
  }
  var f1LocKeys = Object.keys(f1LocMap);
  var f1LocSummary = f1LocKeys.join(', ') + ' @ ' + waspDest.site;

  var f1Detail;
  if (results.length === 1) {
    f1Detail = poRefLabel + '  ' + results[0].sku + ' x' + results[0].quantity;
    if (f1Fail > 0 && f1SubItems[0] && f1SubItems[0].error) f1Detail += '  ' + f1SubItems[0].error;
  } else {
    f1Detail = poRefLabel + '  ' + results.length + ' items';
    if (f1Fail > 0) f1Detail += '  ' + f1Fail + ' error' + (f1Fail > 1 ? 's' : '');
  }
  f1Detail = joinActivitySegments_([
    poRef,
    buildActivityCountSummary_(results.length, 'line', 'lines', 'received')
  ]);
  poRefLabel = poRef;
  var f1ExecId = logActivity('F1', f1Detail, f1Status, buildActivitySourceActionContext_('Katana', isPartialReceive ? 'receive partial' : 'receive', waspDest.site), f1SubItems.length > 0 ? f1SubItems : null, {
    text: poRefLabel,
    url: getKatanaWebUrl('po', poId)
  });

  var f1FlowItems = [];
  for (var fl = 0; fl < results.length; fl++) {
    var fRes = results[fl];
    var fOk = fRes.result && fRes.result.success;

    if (fRes.multiBatch && fRes.multiBatchFirst) {
      // Multi-batch parent row
      var f1fBatchCount = 0;
      for (var f1fbc = fl; f1fbc < results.length; f1fbc++) {
        if (results[f1fbc].sku === fRes.sku && results[f1fbc].multiBatch) f1fBatchCount++;
        else if (f1fbc > fl) break;
      }
      f1FlowItems.push({
        sku: fRes.sku,
        qty: fRes.multiBatchTotal,
        uom: normalizeUom(fRes.uom || ''),
        detail: '→ ' + (fRes.location || waspDest.location) + (fRes.puomNote ? '  uom:' + fRes.puomNote : ''),
        status: '',
        error: '',
        qtyColor: 'grey',
        isParent: true,
        batchCount: f1fBatchCount
      });
    }

    if (fRes.multiBatch) {
      // Nested batch sub-row
      var f1fNestedDetail = '';
      if (fRes.lot) f1fNestedDetail += 'lot:' + fRes.lot;
      if (fRes.expiry) f1fNestedDetail += (f1fNestedDetail ? '  ' : '') + 'exp:' + fRes.expiry;
      f1FlowItems.push({
        sku: '',
        qty: fRes.quantity,
        uom: normalizeUom(fRes.uom || ''),
        expiry: fRes.expiry || '',
        detail: f1fNestedDetail,
        status: fOk ? 'Added' : 'Failed',
        error: fOk ? '' : (fRes.result ? (fRes.result.error || parseWaspError(fRes.result.response, 'Add', fRes.sku)) : ''),
        qtyColor: 'green',
        nested: true
      });
    } else {
      f1FlowItems.push({
        sku: fRes.sku,
        qty: fRes.quantity,
        uom: normalizeUom(fRes.uom || ''),
        expiry: fRes.expiry || '',
        detail: '→ ' + (fRes.location || waspDest.location) + (fRes.lot ? '  lot:' + fRes.lot : '') + (fRes.puomNote ? '  uom:' + fRes.puomNote : ''),
        status: fOk ? 'Added' : 'Failed',
        error: fOk ? '' : (fRes.result ? (fRes.result.error || parseWaspError(fRes.result.response, 'Add', fRes.sku)) : ''),
        qtyColor: 'green'
      });
    }
  }

  for (var f1fd = 0; f1fd < f1FlowItems.length; f1fd++) {
    var f1FlowItem = f1FlowItems[f1fd];
    if (f1FlowItem.expiry && (!f1FlowItem.detail || f1FlowItem.detail.indexOf('exp:') === -1)) {
      f1FlowItem.detail = (f1FlowItem.detail || '') + '  exp:' + f1FlowItem.expiry;
    }
  }

  logFlowDetail('F1', f1ExecId, {
    ref: poRef,
    detail: results.length + ' item' + (results.length !== 1 ? 's' : '') + ' → ' + f1LocSummary,
    status: f1Status === 'success' ? 'Complete' : f1Status === 'failed' ? 'Failed' : 'Partial',
    linkText: poRef,
    linkUrl: getKatanaWebUrl('po', poId)
  }, f1FlowItems);

  sendSlackNotification(
    'PO Received: ' + poRef +
    '\nItems: ' + results.length +
    '\nLocation: ' + f1LocSummary
  );

  markPOReceiveRecent_(poId, 20);

  return {
    status: 'processed',
    poId: poId,
    itemsProcessed: results.length,
    locations: f1LocKeys,
    site: waspDest.site,
    results: results
  };
  } finally {
    releaseExecutionGuard_(poStateGuardKey);
  }
}

// ============================================
// F1: PO REVERT (purchase_order.updated)
// ============================================

/**
 * Handle purchase_order.updated — detects when a PO receive is reverted in Katana.
 * Compares current Katana batch_transactions against what F1 previously sent to WASP.
 * Reverted batches (in stored data but no longer in Katana) are removed from WASP.
 *
 * Setup: subscribe to purchase_order.updated in Katana webhook settings.
 */
function handlePurchaseOrderUpdated(payload) {
  if (!payload || !payload.object) return { status: 'ignored', reason: 'no payload' };
  var poId = payload.object ? payload.object.id : null;
  if (!poId) return { status: 'ignored', reason: 'no PO ID' };
  var releasedForReceiveFallback = false;

  var poStateGuardKey = 'po_state_guard_' + poId;
  var poGuardWait = waitForExecutionGuardToClear_(poStateGuardKey, [0, 1000, 2000, 3000, 4000, 5000, 7000, 9000]);
  if (poGuardWait.waitedMs > 0) {
    logToSheet('PO_UPDATED_RECEIVE_WAIT', {
      poId: poId,
      waitedMs: poGuardWait.waitedMs,
      cleared: poGuardWait.cleared
    }, 'Waiting for in-flight PO receive to finish before revert analysis');
  }

  if (!acquireExecutionGuard_(poStateGuardKey, 120000)) {
    var poUpdateRetryWait = waitForExecutionGuardToClear_(poStateGuardKey, [2000, 4000, 6000]);
    if (poUpdateRetryWait.waitedMs > 0) {
      logToSheet('PO_UPDATED_GUARD_RETRY_WAIT', {
        poId: poId,
        waitedMs: poUpdateRetryWait.waitedMs,
        cleared: poUpdateRetryWait.cleared
      }, 'Retry waiting for PO state guard before updated processing');
    }
    if (!acquireExecutionGuard_(poStateGuardKey, 120000)) {
      return { status: 'skipped', reason: 'PO state change already in progress', poId: poId };
    }
  }

  try {

  var poData = fetchKatanaPO(poId);
  if (!poData) return { status: 'error', message: 'Failed to fetch PO' };

  var poRef = poData.order_no || ('PO-' + poId);
  var poReceiveState = loadPOReceiveState_(poRef);
  var f1SaveKey = poReceiveState.keys.flatKey;
  var f1BlocksKey = poReceiveState.keys.blocksKey;
  var storedJson = poReceiveState.flat && poReceiveState.flat.length ? JSON.stringify(poReceiveState.flat) : '';
  var poRowsDataAll = fetchKatanaPORows(poId);
  var rowsAll = poRowsDataAll ? (poRowsDataAll.data || poRowsDataAll || []) : [];

  if (!storedJson) {
    var poStatusForFallback = String((payload.object && payload.object.status) || poData.status || '').toUpperCase();
    if (poStatusForFallback === 'RECEIVED' || poStatusForFallback === 'PARTIALLY_RECEIVED') {
      var firstReceiveOpenRowDiff = buildSyntheticPOReceiveRowsFromOpenDiff_(poReceiveState.openRows, rowsAll, []);
      if (firstReceiveOpenRowDiff.rows.length > 0) {
        logToSheet('PO_UPDATED_OPEN_DIFF_FIRST', {
          poId: poId,
          poRef: poRef,
          status: poStatusForFallback,
          syntheticRows: firstReceiveOpenRowDiff.rows.length
        }, 'Using Katana open-row delta for first PO receive processing');
        releasedForReceiveFallback = true;
        releaseExecutionGuard_(poStateGuardKey);
        return handlePurchaseOrderReceived({
          action: poStatusForFallback === 'PARTIALLY_RECEIVED' ? 'purchase_order.partially_received' : 'purchase_order.received',
          resource_type: 'purchase_order',
          object: {
            id: poId,
            status: poStatusForFallback,
            href: payload.object && payload.object.href ? payload.object.href : ''
          },
          delegated_from: 'purchase_order.updated.open_rows_first',
          override_rows: firstReceiveOpenRowDiff.rows,
          current_open_rows: rowsAll
        });
      }

      releasedForReceiveFallback = true;
      releaseExecutionGuard_(poStateGuardKey);
      return handlePurchaseOrderReceived({
        action: poStatusForFallback === 'PARTIALLY_RECEIVED' ? 'purchase_order.partially_received' : 'purchase_order.received',
        resource_type: 'purchase_order',
        object: {
          id: poId,
          status: poStatusForFallback,
          href: payload.object && payload.object.href ? payload.object.href : ''
        },
        delegated_from: 'purchase_order.updated'
      });
    }
    return { status: 'ignored', reason: 'No F1 receive data stored for ' + poRef };
  }

  var storedBatches;
  try { storedBatches = JSON.parse(storedJson); } catch (e) {
    return { status: 'error', message: 'Corrupt stored data for ' + poRef };
  }
  var storedBlocks = poReceiveState.blocks || [];

  // Primary revert signal: PO status flipped to NOT_RECEIVED → all stored items reverted.
  // This handles non-lot-tracked items (batchId=null) which the batch comparison can't detect.
  var reverted = [];
  var remaining = [];

  // Rows already fetched above so we can support first-partial diff fallback too.
  var refreshedAfterStatusConfirm = false;

  if (hasRecentPOReceive_(poId) && poData.status !== 'NOT_RECEIVED') {
    // A receive handler JUST processed this PO. Skip the open-row diff to avoid
    // double-processing stale rows. The receive handler already added the correct items.
    // Reverts (NOT_RECEIVED) still process normally — this only skips RECEIVED/PARTIALLY_RECEIVED.
    logToSheet('PO_UPDATED_SKIP_RECENT', {
      poId: poId,
      poRef: poRef,
      status: poData.status
    }, 'Skipping — receive handler just processed this PO');
    releaseExecutionGuard_(poStateGuardKey);
    return { status: 'skipped', reason: 'Recent receive already processed by receive handler', poId: poId, poRef: poRef };
  }

  var poStatusAfterSettle = String(poData.status || '').toUpperCase();
  if (poStatusAfterSettle === 'RECEIVED' || poStatusAfterSettle === 'PARTIALLY_RECEIVED') {
    var immediateOpenRowDiff = buildSyntheticPOReceiveRowsFromOpenDiff_(poReceiveState.openRows, rowsAll, storedBatches);
    if (immediateOpenRowDiff.rows.length > 0) {
      logToSheet('PO_UPDATED_OPEN_DIFF_FAST', {
        poId: poId,
        poRef: poRef,
        status: poStatusAfterSettle,
        syntheticRows: immediateOpenRowDiff.rows.length
      }, 'Using Katana open-row delta for immediate PO receive processing');
      releasedForReceiveFallback = true;
      releaseExecutionGuard_(poStateGuardKey);
      return handlePurchaseOrderReceived({
        action: poStatusAfterSettle === 'PARTIALLY_RECEIVED' ? 'purchase_order.partially_received' : 'purchase_order.received',
        resource_type: 'purchase_order',
        object: {
          id: poId,
          status: poStatusAfterSettle,
          href: payload.object && payload.object.href ? payload.object.href : ''
        },
        delegated_from: 'purchase_order.updated.open_rows_diff.fast',
        override_rows: immediateOpenRowDiff.rows,
        current_open_rows: rowsAll
      });
    }

    var receiveReadyState = waitForPOReceiveReadiness_(poId, poData, poRowsDataAll, storedBatches, poStatusAfterSettle === 'PARTIALLY_RECEIVED' ? [0, 1000, 2000, 3000] : [0, 1000, 2000]);
    if (receiveReadyState.waitedMs > 0) {
      logToSheet('PO_UPDATED_RECEIVE_READY_WAIT', {
        poId: poId,
        poRef: poRef,
        status: poStatusAfterSettle,
        waitedMs: receiveReadyState.waitedMs,
        ready: receiveReadyState.readiness && receiveReadyState.readiness.ready,
        receivableCount: receiveReadyState.readiness ? receiveReadyState.readiness.receivableCount : 0
      }, 'Waiting for Katana PO updated rows to expose new receivable lines');
    }
    poData = receiveReadyState.poData || poData;
    poRowsDataAll = receiveReadyState.rowsData || poRowsDataAll;
    rowsAll = poRowsDataAll ? (poRowsDataAll.data || poRowsDataAll || []) : [];
    poStatusAfterSettle = String(poData.status || '').toUpperCase();
    var receiveReadiness = receiveReadyState.readiness || assessPOReceiveReadiness_(rowsAll, storedBatches, poStatusAfterSettle);
    if (receiveReadiness.ready) {
      releasedForReceiveFallback = true;
      releaseExecutionGuard_(poStateGuardKey);
      return handlePurchaseOrderReceived({
        action: poStatusAfterSettle === 'PARTIALLY_RECEIVED' ? 'purchase_order.partially_received' : 'purchase_order.received',
        resource_type: 'purchase_order',
        object: {
          id: poId,
          status: poStatusAfterSettle,
          href: payload.object && payload.object.href ? payload.object.href : ''
        },
        delegated_from: 'purchase_order.updated',
        delegated_receivable_count: receiveReadiness.receivableCount
      });
    }

    var openRowDiff = buildSyntheticPOReceiveRowsFromOpenDiff_(poReceiveState.openRows, rowsAll, storedBatches);
    if (openRowDiff.rows.length > 0) {
      logToSheet('PO_UPDATED_OPEN_DIFF', {
        poId: poId,
        poRef: poRef,
        status: poStatusAfterSettle,
        syntheticRows: openRowDiff.rows.length
      }, 'Synthesizing PO receive from Katana open-row delta');
      releasedForReceiveFallback = true;
      releaseExecutionGuard_(poStateGuardKey);
      return handlePurchaseOrderReceived({
        action: poStatusAfterSettle === 'PARTIALLY_RECEIVED' ? 'purchase_order.partially_received' : 'purchase_order.received',
        resource_type: 'purchase_order',
        object: {
          id: poId,
          status: poStatusAfterSettle,
          href: payload.object && payload.object.href ? payload.object.href : ''
        },
        delegated_from: 'purchase_order.updated.open_rows_diff',
        override_rows: openRowDiff.rows,
        current_open_rows: rowsAll
      });
    }

    if (poStatusAfterSettle === 'PARTIALLY_RECEIVED') {
      var delayedReceiveReadyState = waitForPOReceiveReadiness_(poId, poData, poRowsDataAll, storedBatches, [4000, 6000]);
      if (delayedReceiveReadyState.waitedMs > 0) {
        logToSheet('PO_UPDATED_RECEIVE_LATE_WAIT', {
          poId: poId,
          poRef: poRef,
          status: poStatusAfterSettle,
          waitedMs: delayedReceiveReadyState.waitedMs,
          ready: delayedReceiveReadyState.readiness && delayedReceiveReadyState.readiness.ready,
          receivableCount: delayedReceiveReadyState.readiness ? delayedReceiveReadyState.readiness.receivableCount : 0
        }, 'Late retry waiting for Katana partial PO rows');
      }
      poData = delayedReceiveReadyState.poData || poData;
      poRowsDataAll = delayedReceiveReadyState.rowsData || poRowsDataAll;
      rowsAll = poRowsDataAll ? (poRowsDataAll.data || poRowsDataAll || []) : [];
      poStatusAfterSettle = String(poData.status || '').toUpperCase();

      var delayedOpenRowDiff = buildSyntheticPOReceiveRowsFromOpenDiff_(poReceiveState.openRows, rowsAll, storedBatches);
      if (delayedOpenRowDiff.rows.length > 0) {
        logToSheet('PO_UPDATED_OPEN_DIFF_LATE', {
          poId: poId,
          poRef: poRef,
          status: poStatusAfterSettle,
          syntheticRows: delayedOpenRowDiff.rows.length
        }, 'Late open-row delta detected for PO receive processing');
        releasedForReceiveFallback = true;
        releaseExecutionGuard_(poStateGuardKey);
        return handlePurchaseOrderReceived({
          action: 'purchase_order.partially_received',
          resource_type: 'purchase_order',
          object: {
            id: poId,
            status: poStatusAfterSettle,
            href: payload.object && payload.object.href ? payload.object.href : ''
          },
          delegated_from: 'purchase_order.updated.open_rows_diff.late',
          override_rows: delayedOpenRowDiff.rows,
          current_open_rows: rowsAll
        });
      }

      var delayedReceiveReadiness = delayedReceiveReadyState.readiness || assessPOReceiveReadiness_(rowsAll, storedBatches, poStatusAfterSettle);
      if (delayedReceiveReadiness.ready) {
        releasedForReceiveFallback = true;
        releaseExecutionGuard_(poStateGuardKey);
        return handlePurchaseOrderReceived({
          action: 'purchase_order.partially_received',
          resource_type: 'purchase_order',
          object: {
            id: poId,
            status: poStatusAfterSettle,
            href: payload.object && payload.object.href ? payload.object.href : ''
          },
          delegated_from: 'purchase_order.updated.late',
          delegated_receivable_count: delayedReceiveReadiness.receivableCount
        });
      }
    }
  }

  if (poData.status === 'NOT_RECEIVED' && getHotfixFlag_('F1_CONFIRM_FULL_REVERT')) {
    var fullRevertConfirm = refetchKatanaEntityWithAdaptiveConfirm_(
      fetchKatanaPO,
      poId,
      function(entity) { return String((entity && entity.status) || '').toUpperCase() === 'NOT_RECEIVED'; },
      hasRecentPOReceive_(poId) ? [2000, 4000, 6000] : [4000]
    );
    var confirmedPO = fullRevertConfirm.entity;
    if (confirmedPO) {
      poData = confirmedPO;
      var confirmRowsDataAll = fetchKatanaPORows(poId);
      if (confirmRowsDataAll) {
        poRowsDataAll = confirmRowsDataAll;
        rowsAll = confirmRowsDataAll.data || confirmRowsDataAll || [];
      }
      refreshedAfterStatusConfirm = true;
      if (!fullRevertConfirm.matched) {
        return { status: 'ignored', reason: 'PO full revert not confirmed for ' + poRef };
      }
      for (var fr = 0; fr < rowsAll.length; fr++) {
        if (isKatanaPOReceivedRow_(rowsAll[fr])) {
          return { status: 'ignored', reason: 'PO still has received rows after full revert confirm for ' + poRef };
        }
      }
    } else {
      return { status: 'skipped', reason: 'F1 full revert confirmation fetch failed for ' + poRef };
    }
  }

  // Diagnostic: log full raw PO rows to Webhook Queue so partial-revert patterns can be analysed
  try {
    var diagRows = [];
    for (var dri = 0; dri < rowsAll.length; dri++) {
      var drow = rowsAll[dri];
      var dbt = drow.batch_transactions || drow.batchTransactions || [];
      var dbtIds = [];
      for (var dbi = 0; dbi < dbt.length; dbi++) {
        dbtIds.push(dbt[dbi].batch_id || dbt[dbi].batchId || null);
      }
      diagRows.push({
        variant_id: drow.variant_id || null,
        qty: drow.quantity || null,
        received_qty: drow.received_quantity || drow.receivedQuantity || drow.received_qty || null,
        row_status: drow.status || null,
        bt_count: dbt.length,
        bt_ids: dbtIds
      });
    }
    var storedDiag = [];
    for (var sdi = 0; sdi < storedBatches.length; sdi++) {
      storedDiag.push({ sku: storedBatches[sdi].sku, batchId: storedBatches[sdi].batchId });
    }
    logWebhookQueue(
      { action: 'po_updated.diag', poRef: poRef, poStatus: poData.status, rows: diagRows, stored: storedDiag },
      { status: 'diag' }
    );
  } catch (diagErr) { Logger.log('PO diag log error: ' + diagErr.message); }

  if (poReceiveState.hasBlockStorage) {
    var blockPlan = planPOReceiveBlockReverts_(storedBlocks, rowsAll);
    if (blockPlan.revertedBlocks.length === 0) {
      persistPOOpenRowsSnapshot_(poRef, rowsAll);
      return { status: 'ignored', reason: 'No reverted receipt blocks detected for ' + poRef };
    }

    var poLocationIdBlocks = poData.location_id || null;
    var waspDestBlocks = { site: CONFIG.WASP_SITE, location: FLOWS.PO_RECEIVING_LOCATION };
    if (poLocationIdBlocks) {
      var poLocationBlocks = fetchKatanaLocation(poLocationIdBlocks);
      var locNameBlocks = poLocationBlocks ? (poLocationBlocks.name || '') : '';
      if (locNameBlocks && KATANA_LOCATION_TO_WASP[locNameBlocks]) {
        waspDestBlocks = KATANA_LOCATION_TO_WASP[locNameBlocks];
      }
    }

    var totalRevertedCount = 0;
    var rvUomCache = {};
    for (var rb = 0; rb < blockPlan.revertedBlocks.length; rb++) {
      var revertedBlock = blockPlan.revertedBlocks[rb];
      var blockItems = revertedBlock.items || [];
      if (!blockItems.length) continue;

      var rvResultsBlock = [];
      for (var rbi = 0; rbi < blockItems.length; rbi++) {
        var rvItem = blockItems[rbi];
        var rvSite = rvItem.site || waspDestBlocks.site;
        var rvLoc = rvItem.location || waspDestBlocks.location;
        markSyncedToWasp(rvItem.sku, rvLoc, 'remove');
        var rvResultBlock = waspRemoveInventoryWithLot(
          rvItem.sku, rvItem.qty, rvLoc, rvItem.lot || '',
          'PO Reverted: ' + poRef, rvSite, rvItem.expiry || ''
        );
        rvResultsBlock.push({
          sku: rvItem.sku,
          qty: rvItem.qty,
          uom: rvItem.uom || resolveSkuUom(rvItem.sku, rvUomCache),
          lot: rvItem.lot,
          expiry: rvItem.expiry,
          location: rvLoc,
          result: rvResultBlock
        });
      }

      var rvSuccessBlock = 0;
      var rvFailBlock = 0;
      var rvSubItemsBlock = [];
      for (var rba = 0; rba < rvResultsBlock.length; rba++) {
        var rvBlockItem = rvResultsBlock[rba];
        var rvBlockOk = rvBlockItem.result && rvBlockItem.result.success;
        if (rvBlockOk) rvSuccessBlock++; else rvFailBlock++;
        rvSubItemsBlock.push({
          sku: rvBlockItem.sku,
          qty: rvBlockItem.qty,
          uom: normalizeUom(rvBlockItem.uom || ''),
          success: rvBlockOk,
          status: rvBlockOk ? 'Removed' : 'Failed',
          error: rvBlockOk ? '' : (rvBlockItem.result ? (rvBlockItem.result.error || parseWaspError(rvBlockItem.result.response, 'Remove', rvBlockItem.sku)) : ''),
          action: buildActivityCompactMeta_(waspDestBlocks.site, rvBlockItem.location, rvBlockItem.lot, rvBlockItem.expiry),
          qtyColor: 'red'
        });
      }

      var rvStatusBlock = rvFailBlock === 0 ? 'reverted' : rvSuccessBlock === 0 ? 'failed' : 'partial';
      var rvDetailBlock = joinActivitySegments_([
        revertedBlock.label || extractCanonicalActivityRef_(poRef, 'PO-', poId),
        buildActivityCountSummary_(blockItems.length, 'line', 'lines', 'reverted')
      ]);

      logActivity('F1', rvDetailBlock, rvStatusBlock, buildActivitySourceActionContext_('Katana', 'receive revert', waspDestBlocks.site), rvSubItemsBlock, {
        text: poRef,
        url: getKatanaWebUrl('po', poId)
      });

      totalRevertedCount += blockItems.length;
    }

    persistPOReceiveState_(poRef, blockPlan.remainingBlocks);
    persistPOOpenRowsSnapshot_(poRef, rowsAll);

    var stateKeys = getPOReceiveStateKeys_(poRef);
    CacheService.getScriptCache().remove('po_received_' + poRef);
    if (blockPlan.remainingBlocks.length === 0) {
      PropertiesService.getScriptProperties().deleteProperty(stateKeys.partialKey);
      PropertiesService.getScriptProperties().deleteProperty(stateKeys.countKey);
    } else if (String(poData.status || '').toUpperCase() === 'PARTIALLY_RECEIVED') {
      PropertiesService.getScriptProperties().setProperty(stateKeys.partialKey, 'true');
    } else {
      PropertiesService.getScriptProperties().deleteProperty(stateKeys.partialKey);
    }

    return { status: 'reverted', count: totalRevertedCount, poRef: poRef, blocks: blockPlan.revertedBlocks.length };
  }

  if (poData.status === 'NOT_RECEIVED') {
    // Full un-receive: every stored item was removed from Katana
    reverted = storedBatches.slice();
  } else {
    // PO still received — fall back to batch-ID comparison for lot-tracked partial reverts
    var poRowsData = poRowsDataAll;
    var rows = rowsAll;
    var currentBatchIds = {};
    var currentNonBatchReceivedQtyByRow = buildCurrentPOReceivedQtyByRow_(rows, poData.status);
    for (var i = 0; i < rows.length; i++) {
      var bt = rows[i].batch_transactions || rows[i].batchTransactions || [];
      for (var j = 0; j < bt.length; j++) {
        var bid = bt[j].batch_id || bt[j].batchId || null;
        if (bid) currentBatchIds[String(bid)] = true;
      }
    }
    var storedNonBatchByRow = {};
    for (var s = 0; s < storedBatches.length; s++) {
      var stored = storedBatches[s];
      if (stored.batchId) {
        // Only flag as reverted if it had a batchId that is now gone
        if (!currentBatchIds[String(stored.batchId)]) {
          reverted.push(stored);
        } else {
          remaining.push(stored);
        }
        continue;
      }

      if ((typeof HOTFIX_FLAGS !== 'undefined' && HOTFIX_FLAGS.F1_PARTIAL_NON_BATCH_DELTA) && stored.poRowId) {
        var storedRowKey = String(stored.poRowId);
        if (!storedNonBatchByRow[storedRowKey]) storedNonBatchByRow[storedRowKey] = [];
        storedNonBatchByRow[storedRowKey].push(stored);
        continue;
      }

      remaining.push(stored);
    }

    var shouldConfirmNonBatchRevert = !!(
      getHotfixFlag_('F1_PARTIAL_NON_BATCH_DELTA') &&
      getHotfixFlag_('F1_CONFIRM_NON_BATCH_REVERT')
    );
    var hasNonBatchRevertCandidate = false;
    if (shouldConfirmNonBatchRevert) {
      for (var candidateRowKey in storedNonBatchByRow) {
        if (!storedNonBatchByRow.hasOwnProperty(candidateRowKey)) continue;
        var storedRowQty = 0;
        for (var cs = 0; cs < storedNonBatchByRow[candidateRowKey].length; cs++) {
          storedRowQty += parseFloat(storedNonBatchByRow[candidateRowKey][cs].qty || 0) || 0;
        }
        var currentRowQty = parseFloat(currentNonBatchReceivedQtyByRow[candidateRowKey] || 0) || 0;
        if (currentRowQty < storedRowQty) {
          hasNonBatchRevertCandidate = true;
          break;
        }
      }
    }

    if (hasNonBatchRevertCandidate && !refreshedAfterStatusConfirm) {
      var confirmRows = refetchKatanaEntityAfterConfirmDelay_(fetchKatanaPORows, poId);
      confirmRows = confirmRows && confirmRows.data ? confirmRows.data : (confirmRows || []);
      if (confirmRows.length > 0) {
        currentNonBatchReceivedQtyByRow = buildCurrentPOReceivedQtyByRow_(confirmRows, poData.status);
      }
    }

    for (var rowKey in storedNonBatchByRow) {
      if (!storedNonBatchByRow.hasOwnProperty(rowKey)) continue;
      var rowEntries = storedNonBatchByRow[rowKey];
      var qtyToKeep = parseFloat(currentNonBatchReceivedQtyByRow[rowKey] || 0) || 0;

      for (var re = 0; re < rowEntries.length; re++) {
        var storedEntry = rowEntries[re];
        var storedQty = parseFloat(storedEntry.qty || 0) || 0;

        if (qtyToKeep >= storedQty) {
          remaining.push(storedEntry);
          qtyToKeep -= storedQty;
          continue;
        }

        if (qtyToKeep > 0) {
          var keepEntry = JSON.parse(JSON.stringify(storedEntry));
          keepEntry.qty = qtyToKeep;
          remaining.push(keepEntry);

          var revertEntry = JSON.parse(JSON.stringify(storedEntry));
          revertEntry.qty = storedQty - qtyToKeep;
          reverted.push(revertEntry);
          qtyToKeep = 0;
          continue;
        }

        reverted.push(storedEntry);
      }
    }
  }

  if (reverted.length === 0) {
    persistPOOpenRowsSnapshot_(poRef, rowsAll);
    return { status: 'ignored', reason: 'No reverted batches detected for ' + poRef };
  }

  // Resolve WASP destination (same as original F1 receive)
  var poLocationId = poData.location_id || null;
  var waspDest = { site: CONFIG.WASP_SITE, location: FLOWS.PO_RECEIVING_LOCATION };
  if (poLocationId) {
    var poLocation = fetchKatanaLocation(poLocationId);
    var locName = poLocation ? (poLocation.name || '') : '';
    if (locName && KATANA_LOCATION_TO_WASP[locName]) {
      waspDest = KATANA_LOCATION_TO_WASP[locName];
    }
  }

  // Remove each reverted batch from WASP
  var rvResults = [];
  var rvUomCache = {};
  for (var rv = 0; rv < reverted.length; rv++) {
    var rvItem = reverted[rv];
    var rvSite = rvItem.site || waspDest.site;
    var rvLoc = rvItem.location || waspDest.location;
    markSyncedToWasp(rvItem.sku, rvLoc, 'remove');
    var rvResult = waspRemoveInventoryWithLot(
      rvItem.sku, rvItem.qty, rvLoc, rvItem.lot || '',
      'PO Reverted: ' + poRef, rvSite, rvItem.expiry || ''
    );
    rvResults.push({
      sku: rvItem.sku,
      qty: rvItem.qty,
      uom: rvItem.uom || resolveSkuUom(rvItem.sku, rvUomCache),
      lot: rvItem.lot,
      expiry: rvItem.expiry,
      location: rvLoc,
      result: rvResult
    });
  }

  // Update stored data: remove the reverted batches (delete key entirely if nothing remains)
  if (remaining.length > 0) {
    PropertiesService.getScriptProperties().setProperty(f1SaveKey, JSON.stringify(remaining));
  } else {
    PropertiesService.getScriptProperties().deleteProperty(f1SaveKey);
  }
  persistPOOpenRowsSnapshot_(poRef, rowsAll);

  // Activity Log — F1 Revert
  var rvSuccess = 0;
  var rvFail = 0;
  var rvSubItems = [];
  for (var ra = 0; ra < rvResults.length; ra++) {
    var ra_item = rvResults[ra];
    var raOk = ra_item.result && ra_item.result.success;
    if (raOk) rvSuccess++; else rvFail++;
    rvSubItems.push({
      sku: ra_item.sku,
      qty: ra_item.qty,
      uom: normalizeUom(ra_item.uom || ''),
      success: raOk,
      status: raOk ? 'Removed' : 'Failed',
      error: raOk ? '' : (ra_item.result ? (ra_item.result.error || parseWaspError(ra_item.result.response, 'Remove', ra_item.sku)) : ''),
      action: (ra_item.lot ? 'lot:' + ra_item.lot + '  ' : '') + (ra_item.expiry ? 'exp:' + ra_item.expiry + '  ' : '') + '← ' + ra_item.location,
      qtyColor: 'red'
    });
    rvSubItems[rvSubItems.length - 1].action = buildActivityCompactMeta_(waspDest.site, ra_item.location, ra_item.lot, ra_item.expiry);
  }
  var rvStatus = rvFail === 0 ? 'reverted' : rvSuccess === 0 ? 'failed' : 'partial';
  var rvDetail = joinActivitySegments_([
    extractCanonicalActivityRef_(poRef, 'PO-', poId),
    buildActivityCountSummary_(reverted.length, 'line', 'lines', 'reverted')
  ]);

  logActivity('F1', rvDetail, rvStatus, buildActivitySourceActionContext_('Katana', 'receive revert', waspDest.site), rvSubItems, {
    text: poRef,
    url: getKatanaWebUrl('po', poId)
  });

  // Clear dedup cache and partial/count flags so a fresh receive can go through immediately
  var rPoKey = poRef.replace(/[^a-zA-Z0-9_]/g, '');
  CacheService.getScriptCache().remove('po_received_' + poRef);
  PropertiesService.getScriptProperties().deleteProperty('f1_recv_partial_' + rPoKey);
  PropertiesService.getScriptProperties().deleteProperty('f1_count_' + rPoKey);

  return { status: 'reverted', count: reverted.length, poRef: poRef };
  } finally {
    if (!releasedForReceiveFallback) {
      releaseExecutionGuard_(poStateGuardKey);
    }
  }
}

// ============================================
// F5: SALES ORDER CREATED (Order Fulfillment)
// ============================================

/**
 * Handle Sales Order Created event
 * Creates WASP Pick Order from SHIPPING-DOCK
 */
function handleSalesOrderCreated(payload) {
  var soId = payload.object ? payload.object.id : null;

  if (!soId) {
    logToSheet('SO_CREATED_ERROR', payload, { error: 'No SO ID found' });
    return { status: 'error', message: 'No SO ID in webhook' };
  }

  // Dedup: skip if this SO was already processed recently (300s window)
  var cache = CacheService.getScriptCache();
  var soCrDedupKey = 'so_created_' + soId;
  if (cache.get(soCrDedupKey)) {
    return { status: 'skipped', reason: 'SO already processed', soId: soId };
  }
  cache.put(soCrDedupKey, 'true', 300);

  // Fetch SO details from Katana
  var soData = fetchKatanaSalesOrderFull(soId);
  if (!soData) {
    logToSheet('SO_CREATED_ERROR', { soId: soId }, { error: 'Failed to fetch SO details' });
    return { status: 'error', message: 'Failed to fetch SO details' };
  }

  var so = soData.data ? soData.data : soData;
  var orderNumber = so.order_no || ('SO-' + soId);

  // Get line items from embedded data
  var rows = so.sales_order_rows || [];

  // Get customer name from addresses
  var addresses = so.addresses || [];
  var shippingAddress = null;
  for (var a = 0; a < addresses.length; a++) {
    if (addresses[a].entity_type === 'shipping') {
      shippingAddress = addresses[a];
      break;
    }
  }
  if (!shippingAddress && addresses.length > 0) {
    shippingAddress = addresses[0];
  }

  var customerName = shippingAddress
    ? ((shippingAddress.first_name || '') + ' ' + (shippingAddress.last_name || '')).trim()
    : 'Shopify Customer';

  // Processing SO details

  if (rows.length === 0) {
    logToSheet('SO_CREATED_ERROR', { soId: soId }, { error: 'No line items found' });
    return { status: 'error', message: 'No line items in SO' };
  }

  // Build pick order lines
  var pickOrderLines = [];

  for (var i = 0; i < rows.length; i++) {
    var row = rows[i];
    var variantId = row.variant_id;
    var quantity = row.quantity || 0;

    // Get SKU from variant
    var sku = '';
    if (variantId) {
      var variant = fetchKatanaVariant(variantId);
      var v = variant && variant.data ? variant.data : variant;
      sku = v ? (v.sku || '') : '';
    }

    if (sku && quantity > 0 && !isSkippedSku(sku)) {
      pickOrderLines.push({
        ItemNumber: sku,
        Quantity: quantity,
        UomName: 'Each',
        LocationCode: FLOWS.PICK_FROM_LOCATION  // SHIPPING-DOCK
      });
      // Line added
    } else {
      // Skipped — missing SKU, zero qty, or in SKIP_SKUS
    }
  }

  if (pickOrderLines.length === 0) {
    logToSheet('SO_CREATED_ERROR', { soId: soId }, { error: 'No valid line items' });
    return { status: 'error', message: 'No valid items to pick' };
  }

  // Create Pick Order in WASP
  var pickOrderPayload = [{
    PickOrderNumber: orderNumber,
    CustomerNumber: 'SHOPIFY',
    CustomerName: customerName,
    SiteName: CONFIG.WASP_SITE,
    OrderDate: new Date().toISOString().slice(0, 10),
    IssueOrder: true,
    ReferenceNumber: 'Katana-' + soId,
    Notes: 'Auto-created from Katana SO: ' + orderNumber,
    PickOrderLines: pickOrderLines
  }];

  logToSheet('WASP_PICK_ORDER_PAYLOAD', {
    orderNumber: orderNumber,
    lineCount: pickOrderLines.length,
    location: FLOWS.PICK_FROM_LOCATION
  }, pickOrderPayload);

  var result = waspCreatePickOrder(pickOrderPayload);

  // Handle duplicate order (-70010) — treat as success since pick order already exists
  if (!result.success && result.response && result.response.indexOf('-70010') >= 0) {
    logToSheet('WASP_PICK_ORDER_DUPLICATE', { orderNumber: orderNumber }, result);
    result.success = true;
    result.duplicate = true;
  }

  if (result.success) {
    // Store mapping for later lookup
    storePickOrderMapping(orderNumber, soId);

    logToSheet('WASP_PICK_ORDER_CREATED', {
      orderNumber: orderNumber,
      katanaSoId: soId
    }, result);

    sendSlackNotification(
      'Pick Order Created\n' +
      'Order: ' + orderNumber + '\n' +
      'Items: ' + pickOrderLines.length + '\n' +
      'Location: ' + FLOWS.PICK_FROM_LOCATION
    );
  } else {
    logToSheet('WASP_PICK_ORDER_ERROR', { orderNumber: orderNumber }, result);
    sendSlackNotification(
      'Pick Order Failed\n' +
      'Order: ' + orderNumber + '\n' +
      'Error: ' + (result.response || 'Unknown')
    );
  }

  // No Activity log here — F5 logs when SO is actually delivered (handleSalesOrderDelivered)
  // Pick order creation is logged in Debug tab only

  return {
    status: result.success ? 'processed' : 'failed',
    soId: soId,
    orderNumber: orderNumber,
    pickOrderCreated: result.success,
    itemCount: pickOrderLines.length,
    location: FLOWS.PICK_FROM_LOCATION
  };
}

// ============================================
// F3/F5/F6: SALES ORDER DELIVERED
// ============================================
// F5 (Shopify): SO has pick order mapping → skip WASP removal if pick completed
// F6 (Amazon US): exact Amazon US customer match → remove SHIPPING-DOCK + add AMAZON-FBA-USA
// F3 (other): No pick order mapping, not Amazon FBA → remove from AMAZON_TRANSFER_LOCATION
// ============================================

function isAmazonUSSalesOrder_(so, payloadObject) {
  var ids = FLOWS.AMAZON_CUSTOMER_IDS || [];
  var idMap = {};
  for (var i = 0; i < ids.length; i++) {
    if (ids[i] !== null && ids[i] !== undefined && ids[i] !== '') {
      idMap[String(ids[i]).trim()] = true;
    }
  }

  function extractTrailingId_(value) {
    var text = String(value || '').trim();
    if (!text) return '';
    var directMatch = text.match(/^\d+$/);
    if (directMatch) return directMatch[0];
    var urlMatch = text.match(/\/(\d+)(?:\?.*)?$/);
    if (urlMatch) return urlMatch[1];
    return '';
  }

  function pushCandidateId_(target, value) {
    if (value === null || value === undefined || value === '') return;
    var raw = String(value).trim();
    if (!raw) return;
    target.push(raw);
    var extracted = extractTrailingId_(raw);
    if (extracted && extracted !== raw) target.push(extracted);
  }

  function collectIdCandidates_(entity, target) {
    if (!entity) return;
    pushCandidateId_(target, entity.customer_id);
    pushCandidateId_(target, entity.customerId);
    pushCandidateId_(target, entity.customer_href);
    pushCandidateId_(target, entity.customer_reference_id);
    if (entity.customer && typeof entity.customer === 'object') {
      pushCandidateId_(target, entity.customer.id);
      pushCandidateId_(target, entity.customer.customer_id);
      pushCandidateId_(target, entity.customer.customerId);
      pushCandidateId_(target, entity.customer.href);
      pushCandidateId_(target, entity.customer.url);
      pushCandidateId_(target, entity.customer.reference_id);
    } else if (String(entity.customer || '').trim()) {
      pushCandidateId_(target, entity.customer);
    }
  }

  var candidateIds = [];
  collectIdCandidates_(so, candidateIds);
  collectIdCandidates_(payloadObject, candidateIds);
  for (var ci = 0; ci < candidateIds.length; ci++) {
    var candidateId = candidateIds[ci];
    if (candidateId !== null && candidateId !== undefined && candidateId !== '') {
      if (idMap[String(candidateId).trim()]) return true;
    }
  }

  function collectNameCandidates_(entity, target) {
    if (!entity) return;
    target.push(entity.name || '');
    target.push(entity.order_no || '');
    target.push(entity.customer_name || '');
    target.push(entity.customer_display_name || '');
    target.push(entity.customer_company || '');
    target.push(entity.customer_reference || '');
    target.push(entity.customer_reference_number || '');
    target.push(entity.customer_reference_id || '');
    target.push(entity.ship_to_name || '');
    target.push(entity.bill_to_name || '');
    target.push(entity.shipping_name || '');
    target.push(entity.billing_name || '');
    target.push(entity.ship_to_company || '');
    target.push(entity.bill_to_company || '');
    target.push(entity.shipping_company || '');
    target.push(entity.billing_company || '');
    target.push(entity.ship_to_city || '');
    target.push(entity.bill_to_city || '');
    target.push(entity.shipping_city || '');
    target.push(entity.billing_city || '');
    target.push(entity.ship_to_country || '');
    target.push(entity.bill_to_country || '');
    target.push(entity.shipping_country || '');
    target.push(entity.billing_country || '');
    if (entity.ship_to && typeof entity.ship_to === 'object') {
      target.push(entity.ship_to.name || '');
      target.push(entity.ship_to.company || '');
      target.push(entity.ship_to.city || '');
      target.push(entity.ship_to.country || '');
      target.push(entity.ship_to.address || '');
      target.push(entity.ship_to.address_line_1 || '');
      target.push(entity.ship_to.address_line_2 || '');
    }
    if (entity.bill_to && typeof entity.bill_to === 'object') {
      target.push(entity.bill_to.name || '');
      target.push(entity.bill_to.company || '');
      target.push(entity.bill_to.city || '');
      target.push(entity.bill_to.country || '');
      target.push(entity.bill_to.address || '');
      target.push(entity.bill_to.address_line_1 || '');
      target.push(entity.bill_to.address_line_2 || '');
    }
    if (Array.isArray(entity.addresses)) {
      for (var ai = 0; ai < entity.addresses.length; ai++) {
        var address = entity.addresses[ai] || {};
        target.push(address.name || '');
        target.push(address.company || '');
        target.push(address.city || '');
        target.push(address.state || '');
        target.push(address.country || '');
        target.push(address.address || '');
        target.push(address.line_1 || '');
        target.push(address.line_2 || '');
        target.push(address.zip || '');
        target.push(((address.first_name || '') + ' ' + (address.last_name || '')).trim());
        target.push(address.entity_type || '');
      }
    }
    if (entity.customer && typeof entity.customer === 'object') {
      target.push(entity.customer.name || '');
      target.push(entity.customer.display_name || '');
      target.push(entity.customer.company || '');
      target.push(entity.customer.full_name || '');
      target.push(((entity.customer.first_name || '') + ' ' + (entity.customer.last_name || '')).trim());
    } else {
      target.push(entity.customer || '');
    }
  }

  function normalizeAmazonNameCandidate_(value) {
    return String(value || '')
      .trim()
      .toLowerCase()
      .replace(/[^a-z0-9]+/g, ' ')
      .replace(/\s+/g, ' ')
      .trim();
  }

  function hasAmazonUSName_(names) {
    for (var ni = 0; ni < names.length; ni++) {
      var candidateName = normalizeAmazonNameCandidate_(names[ni]);
      if (!candidateName) continue;
      var compact = candidateName.replace(/\s+/g, '');
      if (candidateName === 'amazon us') return true;
      if (candidateName.indexOf('amazon us ') === 0) return true;
      if (candidateName.indexOf(' amazon us ') >= 0) return true;
      if (compact === 'amazonus' || compact.indexOf('amazonus') === 0) return true;
    }
    return false;
  }

  function hasAmazonLikeName_(names) {
    for (var ai = 0; ai < names.length; ai++) {
      var candidateName = normalizeAmazonNameCandidate_(names[ai]);
      if (!candidateName) continue;
      var compact = candidateName.replace(/\s+/g, '');
      if (candidateName.indexOf('amazon') >= 0 || compact.indexOf('amazon') >= 0) return true;
    }
    return false;
  }

  function hasUSAddressMarker_(names) {
    for (var ui = 0; ui < names.length; ui++) {
      var candidateName = normalizeAmazonNameCandidate_(names[ui]);
      if (!candidateName) continue;
      if (candidateName === 'usa' || candidateName === 'us' || candidateName === 'united states') return true;
      if (candidateName.indexOf(' usa ') >= 0 || candidateName.indexOf(' united states ') >= 0) return true;
      if (candidateName.indexOf(' usa') === candidateName.length - 4) return true;
      if (candidateName.indexOf(' united states') === candidateName.length - 14) return true;
    }
    return false;
  }

  var candidateNames = [];
  collectNameCandidates_(so, candidateNames);
  collectNameCandidates_(payloadObject, candidateNames);
  if (hasAmazonUSName_(candidateNames)) return true;
  if (hasAmazonLikeName_(candidateNames) && hasUSAddressMarker_(candidateNames)) return true;

  for (var fi = 0; fi < candidateIds.length; fi++) {
    var fetchId = candidateIds[fi];
    if (fetchId === null || fetchId === undefined || fetchId === '') continue;
    var customerData = fetchKatanaCustomer(fetchId);
    var customer = customerData && customerData.data ? customerData.data : customerData;
    if (!customer) continue;
    var fetchedNames = [];
    fetchedNames.push(customer.name || '');
    fetchedNames.push(customer.display_name || '');
    fetchedNames.push(customer.company || '');
    fetchedNames.push(customer.full_name || '');
    fetchedNames.push(((customer.first_name || '') + ' ' + (customer.last_name || '')).trim());
    if (hasAmazonUSName_(fetchedNames)) return true;
    if (hasAmazonLikeName_(fetchedNames) && hasUSAddressMarker_(candidateNames.concat(fetchedNames))) return true;
  }

  return false;
}

function buildAmazonUSSalesOrderMatchDebug_(so, payloadObject) {
  function pushUnique_(target, value) {
    var text = String(value || '').trim();
    if (!text) return;
    if (target.indexOf(text) < 0) target.push(text);
  }

  function extractTrailingIdForDebug_(value) {
    var text = String(value || '').trim();
    if (!text) return '';
    var directMatch = text.match(/^\d+$/);
    if (directMatch) return directMatch[0];
    var urlMatch = text.match(/\/(\d+)(?:\?.*)?$/);
    if (urlMatch) return urlMatch[1];
    return '';
  }

  function collectIds_(entity, target) {
    if (!entity) return;
    var values = [
      entity.customer_id,
      entity.customerId,
      entity.customer_href,
      entity.customer_reference_id
    ];
    if (entity.customer && typeof entity.customer === 'object') {
      values = values.concat([
        entity.customer.id,
        entity.customer.customer_id,
        entity.customer.customerId,
        entity.customer.href,
        entity.customer.url,
        entity.customer.reference_id
      ]);
    } else {
      values.push(entity.customer);
    }
    for (var i = 0; i < values.length; i++) {
      var raw = String(values[i] || '').trim();
      if (!raw) continue;
      pushUnique_(target, raw);
      var extracted = extractTrailingIdForDebug_(raw);
      if (extracted) pushUnique_(target, extracted);
    }
  }

  function collectNames_(entity, target) {
    if (!entity) return;
    var values = [
      entity.name,
      entity.order_no,
      entity.customer_name,
      entity.customer_display_name,
      entity.customer_company,
      entity.customer_reference,
      entity.customer_reference_number,
      entity.ship_to_name,
      entity.bill_to_name,
      entity.shipping_name,
      entity.billing_name,
      entity.ship_to_company,
      entity.bill_to_company,
      entity.shipping_company,
      entity.billing_company,
      entity.ship_to_city,
      entity.bill_to_city,
      entity.shipping_city,
      entity.billing_city,
      entity.ship_to_country,
      entity.bill_to_country,
      entity.shipping_country,
      entity.billing_country
    ];
    if (entity.ship_to && typeof entity.ship_to === 'object') {
      values = values.concat([
        entity.ship_to.name,
        entity.ship_to.company,
        entity.ship_to.city,
        entity.ship_to.country,
        entity.ship_to.address,
        entity.ship_to.address_line_1,
        entity.ship_to.address_line_2
      ]);
    }
    if (entity.bill_to && typeof entity.bill_to === 'object') {
      values = values.concat([
        entity.bill_to.name,
        entity.bill_to.company,
        entity.bill_to.city,
        entity.bill_to.country,
        entity.bill_to.address,
        entity.bill_to.address_line_1,
        entity.bill_to.address_line_2
      ]);
    }
    if (entity.customer && typeof entity.customer === 'object') {
      values = values.concat([
        entity.customer.name,
        entity.customer.display_name,
        entity.customer.company,
        entity.customer.full_name,
        ((entity.customer.first_name || '') + ' ' + (entity.customer.last_name || '')).trim()
      ]);
    } else {
      values.push(entity.customer);
    }
    for (var i = 0; i < values.length; i++) {
      pushUnique_(target, values[i]);
    }
  }

  var ids = [];
  var names = [];
  collectIds_(so, ids);
  collectIds_(payloadObject, ids);
  collectNames_(so, names);
  collectNames_(payloadObject, names);

  return {
    soId: so && so.id ? so.id : '',
    order_no: so && so.order_no ? so.order_no : '',
    customer_id: so && so.customer_id ? so.customer_id : '',
    customer_name: so && so.customer_name ? so.customer_name : '',
    payload_status: payloadObject && payloadObject.status ? payloadObject.status : '',
    candidateIds: ids.slice(0, 12),
    candidateNames: names.slice(0, 16)
  };
}

function summarizeAmazonUSSalesOrderMatchDebug_(diag) {
  diag = diag || {};
  var ids = Array.isArray(diag.candidateIds) ? diag.candidateIds.slice(0, 4).join(',') : '';
  var names = Array.isArray(diag.candidateNames) ? diag.candidateNames.slice(0, 3).join(' | ') : '';
  var parts = [];
  if (diag.customer_id) parts.push('customer_id=' + diag.customer_id);
  if (diag.customer_name) parts.push('customer_name=' + diag.customer_name);
  if (ids) parts.push('ids=' + ids);
  if (names) parts.push('names=' + names);
  return parts.join(' ; ');
}

function enrichSalesOrderWithAddresses_(soId, so) {
  if (!soId || !so) return so || null;
  if (Array.isArray(so.addresses) && so.addresses.length > 0) return so;
  var addrData = fetchKatanaSalesOrderAddresses(soId);
  var addresses = addrData && addrData.data ? addrData.data : (addrData || []);
  if (!addresses || !addresses.length) return so;
  var enriched = {};
  for (var key in so) {
    if (Object.prototype.hasOwnProperty.call(so, key)) enriched[key] = so[key];
  }
  enriched.addresses = addresses;
  return enriched;
}

function getSODeliveryRowKey_(item) {
  item = item || {};
  return String(
    item.id ||
    item.sales_order_row_id ||
    item.salesOrderRowId ||
    item.row_id ||
    item.line_id ||
    item.variant_id ||
    ''
  ).trim();
}

function buildSODeliverySignature_(items) {
  items = items || [];
  var parts = [];
  for (var i = 0; i < items.length; i++) {
    var item = items[i] || {};
    var rowKey = getSODeliveryRowKey_(item);
    var deliveredQty = (item.delivered_quantity != null) ? item.delivered_quantity : '';
    parts.push([rowKey, deliveredQty, item.variant_id || ''].join(':'));
  }
  parts.sort();
  var raw = parts.join('|') || 'empty';
  var digest = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, raw, Utilities.Charset.UTF_8);
  var hex = '';
  for (var j = 0; j < digest.length; j++) {
    var byteValue = digest[j];
    if (byteValue < 0) byteValue += 256;
    var hexPart = byteValue.toString(16);
    if (hexPart.length < 2) hexPart = '0' + hexPart;
    hex += hexPart;
  }
  return hex || 'empty';
}

function loadF6DeliveryState_(soId, orderNo) {
  var raw = PropertiesService.getScriptProperties().getProperty('f6_delivered_' + soId);
  if (!raw) {
    return { orderNo: orderNo || ('SO-' + soId), items: [], rowProcessedQty: {} };
  }
  try {
    var parsed = JSON.parse(raw);
    return {
      orderNo: parsed.orderNo || orderNo || ('SO-' + soId),
      items: Array.isArray(parsed.items) ? parsed.items : [],
      rowProcessedQty: parsed.rowProcessedQty && typeof parsed.rowProcessedQty === 'object' ? parsed.rowProcessedQty : {}
    };
  } catch (e) {
    Logger.log('loadF6DeliveryState error: ' + e.message);
    return { orderNo: orderNo || ('SO-' + soId), items: [], rowProcessedQty: {} };
  }
}

function persistF6DeliveryState_(soId, state) {
  state = state || {};
  PropertiesService.getScriptProperties().setProperty(
    'f6_delivered_' + soId,
    JSON.stringify({
      orderNo: state.orderNo || ('SO-' + soId),
      items: state.items || [],
      rowProcessedQty: state.rowProcessedQty || {}
    })
  );
}

/**
 * Handle Sales Order Delivered event
 * Detects F5 (Shopify with pick order) vs F6 (Amazon FBA pallet SO) vs F3 (other)
 */
function handleSalesOrderDelivered(payload) {
  var soId = payload.object ? payload.object.id : null;

  if (!soId) {
    logToSheet('SO_DELIVERED_ERROR', payload, { error: 'No SO ID found' });
    return { status: 'error', message: 'No SO ID in webhook' };
  }

  // Fetch SO details
  var soData = fetchKatanaSalesOrder(soId);
  if (!soData) {
    return { status: 'error', message: 'Failed to fetch SO details' };
  }

  // Unwrap Katana API data wrapper
  var so = soData.data ? soData.data : soData;
  so = enrichSalesOrderWithAddresses_(soId, so);

  // ---- DEBUG LOGGING: SO object structure (remove once root cause confirmed) ----
  try {
    var soTopKeys = Object.keys(so || {});
    Logger.log('[SO_DELIVERED_DEBUG] soId=' + soId + ' | top-level keys: ' + JSON.stringify(soTopKeys));
    Logger.log('[SO_DELIVERED_DEBUG] order_no=' + JSON.stringify(so.order_no) +
      ' | name=' + JSON.stringify(so.name) +
      ' | customer_name=' + JSON.stringify(so.customer_name) +
      ' | customer=' + JSON.stringify(so.customer) +
      ' | customer_id=' + JSON.stringify(so.customer_id) +
      ' | location_id=' + JSON.stringify(so.location_id));
    var debugEmbeddedRows = so.sales_order_rows || so.rows || [];
    Logger.log('[SO_DELIVERED_DEBUG] embedded rows count=' + debugEmbeddedRows.length);
    if (debugEmbeddedRows.length > 0) {
      for (var dbgI = 0; dbgI < debugEmbeddedRows.length; dbgI++) {
        var dbgRow = debugEmbeddedRows[dbgI];
        Logger.log('[SO_DELIVERED_DEBUG] row[' + dbgI + '] variant_id=' + JSON.stringify(dbgRow.variant_id) +
          ' | delivered_quantity=' + JSON.stringify(dbgRow.delivered_quantity) +
          ' | quantity=' + JSON.stringify(dbgRow.quantity));
      }
    }
  } catch (dbgErr) {
    Logger.log('[SO_DELIVERED_DEBUG] Error during debug log: ' + dbgErr.message);
  }
  // ---- END DEBUG LOGGING ----

  // Detect Amazon FBA: SO ships from 'Amazon USA' Katana location
  var soLocationId = so.location_id || null;
  var isAmazonFBA = false;
  var soLocName = '';
  if (soLocationId) {
    var soLoc = fetchKatanaLocation(soLocationId);
    soLocName = soLoc ? (soLoc.name || (soLoc.data && soLoc.data.name) || '') : '';
    isAmazonFBA = (soLocName === FLOWS.AMAZON_FBA_KATANA_LOCATION);
  }

  // F5 guard: Shopify SOs are handled entirely by ShipStation.
  // Ignore on Katana delivery — ShipStation already deducted WASP inventory.
  // so.source is checked case-insensitively; log the value so we can confirm it.
  var soSource = (so.source || '').toLowerCase();
  Logger.log('[SO_DELIVERED] soId=' + soId + ' source=' + JSON.stringify(so.source));
  if (soSource === 'shopify') {
    return { status: 'ignored', reason: 'Shopify SO — inventory handled by ShipStation (F5)', soId: soId };
  }

  // Detect F6 (Amazon US FBA) vs F3 (other)
  // F6: customer must be EXACTLY 'Amazon US' (Katana customer name).
  // Detection uses customer_id (Katana API omits customer_name) as primary source,
  // with exact name match as fallback if name is ever present.
  // Deliberately NOT using broad indexOf('amazon') to prevent matching
  // other Amazon variants (Amazon CA, Amazon UK, etc.).
  var soOrderStr = (so.order_no || so.name || '').toLowerCase();
  var soCustomerStr = (so.customer_name || (so.customer && typeof so.customer === 'object' ? so.customer.name || '' : so.customer) || '').toLowerCase();
  var soCustomerId = so.customer_id || (payload.object ? payload.object.customer_id : 0) || 0;
  var isAmazonUSById = isAmazonUSSalesOrder_({ customer_id: soCustomerId }, null);
  var isAmazonUSByName = isAmazonUSSalesOrder_(so, payload.object || {});
  var isF6 = !isAmazonFBA && (isAmazonUSById || isAmazonUSByName);
  var isF5 = false; // Shopify deliveries exit earlier via source guard; keep false to avoid stray F5 branch crashes.

  // Dedup: skip if this SO was already processed recently.
  // Placed AFTER F6 detection so detection bugs don't permanently block retries.
  // F6 partial deliveries dedup by delivered-quantity signature, not by SO only.
  var cache = CacheService.getScriptCache();
  var flowLabel = isF6 ? 'F6' : 'F3';
  var removeLocation = isAmazonFBA ? FLOWS.AMAZON_FBA_WASP_LOCATION : FLOWS.AMAZON_TRANSFER_LOCATION;
  var removeSite = isAmazonFBA ? FLOWS.AMAZON_FBA_WASP_SITE : null; // null = default site (MMH Kelowna)

  var items = so.sales_order_rows || so.rows || [];

  // If no embedded rows, fetch separately
  if (items.length === 0) {
    var rowsData = fetchKatanaSalesOrderRows(soId);
    items = rowsData && rowsData.data ? rowsData.data : (rowsData || []);
  }

  // F6: Katana SO rows don't have delivered_quantity or done fields.
  // Fetch fulfillment records to find which items were actually delivered.
  if (isF6 && items.length > 0) {
    try {
      var fulfillData = katanaApiCall('sales_order_fulfillments?sales_order_id=' + soId);
      var fulfillments = fulfillData && fulfillData.data ? fulfillData.data : (fulfillData || []);
      if (Array.isArray(fulfillments) && fulfillments.length > 0) {
        // Sort fulfillments by id descending — highest id = newest delivery
        fulfillments.sort(function(a, b) { return (b.id || 0) - (a.id || 0); });
        // Use ONLY the newest fulfillment — this is the delivery that triggered the webhook
        var newestFulfillment = fulfillments[0];
        var newestRows = newestFulfillment.sales_order_fulfillment_rows || newestFulfillment.rows || [];
        var newestRowIds = {};
        for (var fr = 0; fr < newestRows.length; fr++) {
          var fRow = newestRows[fr];
          var fRowId = String(fRow.sales_order_row_id || fRow.id || '');
          var fQty = parseFloat(fRow.quantity) || 0;
          if (fRowId && fQty > 0) {
            newestRowIds[fRowId] = fQty;
          }
        }
        // Filter: only include items from the newest fulfillment
        var f6Delivered = [];
        for (var dqf = 0; dqf < items.length; dqf++) {
          var dqItem = items[dqf];
          var dqRowId = String(dqItem.id || '');
          if (newestRowIds[dqRowId] && newestRowIds[dqRowId] > 0) {
            dqItem.delivered_quantity = newestRowIds[dqRowId];
            f6Delivered.push(dqItem);
          }
        }
        logWebhookQueue(
          { action: 'f6_fulfillment_debug' },
          { status: 'diag', message: 'Fulfillments=' + fulfillments.length +
            ' | newestId=' + newestFulfillment.id +
            ' | newestRowIds=' + JSON.stringify(newestRowIds) +
            ' | filtered=' + f6Delivered.length + '/' + items.length }
        );
        if (f6Delivered.length > 0) {
          items = f6Delivered;
        }
      } else {
        // No fulfillment records found — try rows endpoint as fallback
        logWebhookQueue(
          { action: 'f6_fulfillment_debug' },
          { status: 'diag', message: 'No fulfillment records found for SO ' + soId + ' — processing all rows' }
        );
      }
    } catch (fulfillErr) {
      logWebhookQueue(
        { action: 'f6_fulfillment_error' },
        { status: 'diag', message: 'Fulfillment fetch error: ' + fulfillErr.message }
      );
    }
  }

  var dedupKey = 'so_delivered_' + soId;
  if (isF6) {
    dedupKey += '_' + buildSODeliverySignature_(items);
  }
  if (cache.get(dedupKey)) {
    return { status: 'skipped', reason: 'SO already processed', soId: soId };
  }
  if (isF6 || isAmazonFBA) {
    cache.put(dedupKey, 'true', 300); // 5-minute dedup window (only for actionable flows)
  }

  // ---- DEBUG LOGGING: flow detection + items (remove once root cause confirmed) ----
  try {
    Logger.log('[SO_DELIVERED_DEBUG] soOrderStr=' + JSON.stringify(soOrderStr) +
      ' | soCustomerStr=' + JSON.stringify(soCustomerStr) +
      ' | soCustomerId=' + soCustomerId +
      ' | soLocName=' + JSON.stringify(soLocName) +
      ' | isAmazonUSById=' + isAmazonUSById + ' | isAmazonUSByName=' + isAmazonUSByName +
      ' | isF5=' + isF5 + ' | isF6=' + isF6 + ' | isAmazonFBA=' + isAmazonFBA +
      ' | flowLabel=' + flowLabel + ' | items.length=' + items.length);
    for (var dbgJ = 0; dbgJ < items.length; dbgJ++) {
      var dbgItem = items[dbgJ];
      var dbgVariant = fetchKatanaVariant(dbgItem.variant_id);
      var dbgSku = dbgVariant ? dbgVariant.sku : null;
      var dbgQty = (dbgItem.delivered_quantity != null) ? dbgItem.delivered_quantity : (dbgItem.quantity || 0);
      var dbgSkipped = dbgSku ? isSkippedSku(dbgSku) : 'no-sku';
      Logger.log('[SO_DELIVERED_DEBUG] item[' + dbgJ + '] variant_id=' + JSON.stringify(dbgItem.variant_id) +
        ' | sku=' + JSON.stringify(dbgSku) +
        ' | delivered_quantity=' + JSON.stringify(dbgItem.delivered_quantity) +
        ' | quantity=' + JSON.stringify(dbgItem.quantity) +
        ' | resolvedQty=' + dbgQty +
        ' | isSkippedSku=' + dbgSkipped);
    }
  } catch (dbgErr2) {
    Logger.log('[SO_DELIVERED_DEBUG] Error during flow/item debug log: ' + dbgErr2.message);
  }
  // ---- END DEBUG LOGGING ----

  var results = [];
  var f6State = isF6 ? loadF6DeliveryState_(soId, so.order_no || ('SO-' + soId)) : null;
  var f6NewItems = [];
  var f6UpdatedRowProcessedQty = f6State ? JSON.parse(JSON.stringify(f6State.rowProcessedQty || {})) : null;

  for (var i = 0; i < items.length; i++) {
    var item = items[i];
    var variantId = item.variant_id;
    var quantity = (item.delivered_quantity != null) ? item.delivered_quantity : (item.quantity || 0);

    // Get SKU from variant
    var variant = fetchKatanaVariant(variantId);
    var sku = variant ? variant.sku : null;

    if (sku && quantity > 0 && !isSkippedSku(sku)) {
      if (isF6) {
        var f6RowKey = getSODeliveryRowKey_(item);
        var f6PrevQty = parseFloat((f6UpdatedRowProcessedQty && f6UpdatedRowProcessedQty[f6RowKey]) || 0) || 0;
        var f6CurrentQty = parseFloat(quantity || 0) || 0;
        quantity = f6CurrentQty - f6PrevQty;
        if (quantity <= 0) {
          continue;
        }

        // F6: SO from MMH Kelowna → FBA shipment
        // Step 1: Remove from PRODUCTION (lot-aware)
        // Step 2: Add to AMAZON-FBA-USA with same lot if applicable
        markSyncedToWasp(sku, removeLocation, 'remove', removeSite || CONFIG.WASP_SITE);
        var f6LotInfo = waspLookupItemLotAndDate(sku, removeLocation, null);
        var f6RemoveResult;
        if (f6LotInfo && f6LotInfo.lot) {
          f6RemoveResult = waspRemoveInventoryWithLot(
            sku, quantity, removeLocation, f6LotInfo.lot,
            'F6 FBA Ship: ' + (so.order_no || soId), null, f6LotInfo.dateCode
          );
        } else {
          f6RemoveResult = waspRemoveInventory(
            sku, quantity, removeLocation,
            'F6 FBA Ship: ' + (so.order_no || soId)
          );
        }
        var f6AddResult = null;
        var f6UsedLot = (f6LotInfo && f6LotInfo.lot) ? f6LotInfo.lot : '';
        var f6UsedExpiry = normalizeBusinessDate_((f6LotInfo && f6LotInfo.dateCode) ? f6LotInfo.dateCode : '');
        if (f6RemoveResult && f6RemoveResult.success) {
          markSyncedToWasp(sku, FLOWS.AMAZON_FBA_WASP_LOCATION, 'add', FLOWS.AMAZON_FBA_WASP_SITE);
          if (f6LotInfo && f6LotInfo.lot) {
            f6AddResult = waspAddInventoryWithLot(
              sku, quantity, FLOWS.AMAZON_FBA_WASP_LOCATION, f6LotInfo.lot, f6LotInfo.dateCode,
              'F6 FBA Ship: ' + (so.order_no || soId), FLOWS.AMAZON_FBA_WASP_SITE
            );
          } else {
            f6AddResult = waspAddInventory(
              sku, quantity, FLOWS.AMAZON_FBA_WASP_LOCATION,
              'F6 FBA Ship: ' + (so.order_no || soId),
              FLOWS.AMAZON_FBA_WASP_SITE
            );
          }
        }
        var f6Ok = !!(f6RemoveResult && f6RemoveResult.success && f6AddResult && f6AddResult.success);
        results.push({
          sku: sku,
          quantity: quantity,
          result: { success: f6Ok },
          f6RemoveResult: f6RemoveResult,
          f6AddResult: f6AddResult,
          uom: resolveVariantUom(variant),
          lot: f6UsedLot,
          expiry: f6UsedExpiry,
          isF6: true,
          rowKey: f6RowKey
        });
        if (f6Ok) {
          if (f6UpdatedRowProcessedQty) f6UpdatedRowProcessedQty[f6RowKey] = f6PrevQty + quantity;
          f6NewItems.push({
            sku: sku,
            qty: quantity,
            uom: resolveVariantUom(variant) || '',
            lot: f6UsedLot,
            expiry: normalizeBusinessDate_(f6UsedExpiry || '')
          });
        }
      } else {
        // Mark sync BEFORE WASP call — prevents F2 echo when WASP callout fires
        markSyncedToWasp(sku, removeLocation, 'remove', removeSite || CONFIG.WASP_SITE);

        var result;
        if (isAmazonFBA) {
          // Amazon FBA: try lot lookup — some items are lot-tracked at Amazon
          var amazonLotInfo = waspLookupItemLotAndDate(sku, removeLocation, removeSite);
          if (amazonLotInfo && amazonLotInfo.lot) {
            result = waspRemoveInventoryWithLot(
              sku, quantity, removeLocation, amazonLotInfo.lot,
              'SO Delivered: ' + (so.order_no || soId), removeSite, amazonLotInfo.dateCode
            );
          } else {
            result = waspRemoveInventory(
              sku, quantity, removeLocation,
              'SO Delivered: ' + (so.order_no || soId), removeSite
            );
          }
        } else if (isF5) {
          // F5 pick not yet deducted — remove from SHOPIFY
          result = waspRemoveInventory(
            sku, quantity, removeLocation,
            'SO Delivered: ' + (so.order_no || soId)
          );
          results.push({ sku: sku, quantity: quantity, result: result, uom: resolveVariantUom(variant) });
        }
        // else: unrecognised SO (not F5, F6, or Amazon FBA) — skip, no WASP action needed
      }
    }
  }

  var soRef = so.order_no || ('SO-' + soId);

  // Store F6 data in ScriptProperties for potential revert
  if (isF6 && results.length > 0) {
    var useExactF6RevertState = !!(typeof HOTFIX_FLAGS !== 'undefined' && HOTFIX_FLAGS.F6_REVERT_PRESERVE_EXACT_LOT);
    var mergedF6Items = f6State && Array.isArray(f6State.items) ? f6State.items.slice() : [];
    for (var f6ni = 0; f6ni < f6NewItems.length; f6ni++) {
      var addItem = f6NewItems[f6ni];
      if (!useExactF6RevertState) {
        mergedF6Items.push({ sku: addItem.sku, qty: addItem.qty, uom: addItem.uom || '' });
      } else {
        mergedF6Items.push({
          sku: addItem.sku,
          qty: addItem.qty,
          uom: addItem.uom || '',
          lot: addItem.lot || '',
          expiry: normalizeBusinessDate_(addItem.expiry || '')
        });
      }
    }
    if (mergedF6Items.length > 0) {
      try {
        persistF6DeliveryState_(soId, {
          orderNo: soRef,
          items: mergedF6Items,
          rowProcessedQty: f6UpdatedRowProcessedQty || {}
        });
      } catch (f6StoreErr) {
        Logger.log('F6 store error: ' + f6StoreErr.message);
      }
    }
  }

  // Skip Activity log if no real items were processed (e.g., only OP-* protection items delivered)
  if (results.length === 0) {
    return {
      status: 'skipped',
      soId: soId,
      flow: flowLabel,
      reason: 'No trackable items delivered'
    };
  }

  // Activity Log
  var sdSuccess = 0;
  var sdFail = 0;
  var sdSubItems = [];
  for (var d = 0; d < results.length; d++) {
    var dRes = results[d];
    var dOk = dRes.result && dRes.result.success;
    if (dOk) sdSuccess++; else sdFail++;
    var actionText = dRes.isF6
      ? buildActivityCompactMeta_(FLOWS.AMAZON_FBA_WASP_SITE, FLOWS.AMAZON_FBA_WASP_LOCATION, dRes.lot, dRes.expiry)
      : buildActivityCompactMeta_(CONFIG.WASP_SITE, removeLocation, dRes.lot, dRes.expiry);
    var itemError = '';
    if (!dOk) {
      if (dRes.isF6) {
        var f6rOk = dRes.f6RemoveResult && dRes.f6RemoveResult.success;
        itemError = f6rOk
          ? (dRes.f6AddResult ? (dRes.f6AddResult.error || parseWaspError(dRes.f6AddResult.response, 'Add', dRes.sku)) : 'Add skipped')
          : (dRes.f6RemoveResult ? (dRes.f6RemoveResult.error || parseWaspError(dRes.f6RemoveResult.response, 'Remove', dRes.sku)) : '');
      } else {
        itemError = dRes.result ? (dRes.result.error || parseWaspError(dRes.result.response, 'Remove', dRes.sku)) : '';
      }
    }
    var sdItemStatus = dOk ? (dRes.isF6 ? 'Complete' : 'Deducted') : (dRes.isF6 ? 'Skipped' : 'Failed');
    if (dRes.skipped) sdItemStatus = 'Skipped';
    sdSubItems.push({
      sku: dRes.sku,
      qty: dRes.quantity,
      uom: dRes.uom || '',
      success: dOk || sdItemStatus === 'Skipped',
      status: sdItemStatus,
      error: (sdItemStatus === 'Skipped') ? '' : itemError,
      action: actionText,
      qtyColor: dRes.isF6 ? 'green' : 'red'
    });
  }
  var sdStatus = sdFail === 0 ? 'success' : sdSuccess === 0 ? 'failed' : 'partial';
  var soRefToken = extractCanonicalActivityRef_(soRef, 'SO-', soId);
  var sdDetail = joinActivitySegments_([
    soRefToken,
    isF6
      ? buildActivityCountSummary_(results.length, 'line', 'lines', 'staged')
      : buildActivityCountSummary_(results.length, 'line', 'lines', 'shipped')
  ]);
  var sdLocationLabel = isF6 ? getActivityDisplayLocation_(FLOWS.AMAZON_FBA_WASP_LOCATION) : getActivityDisplayLocation_(removeLocation);
  var sdContext = isF6
    ? buildActivityTransferContext_('Katana', 'deliver', CONFIG.WASP_SITE, FLOWS.AMAZON_FBA_WASP_SITE)
    : buildActivitySourceActionContext_('ShipStation', 'ship deduct', CONFIG.WASP_SITE);
  var sdExecId = logActivity(flowLabel, sdDetail, sdStatus, sdContext, (isF6 || sdSubItems.length > 1) ? sdSubItems : null, {
    text: soRefToken,
    url: getKatanaWebUrl('so', soId)
  });

  var sdFlowItems = [];
  for (var fl = 0; fl < results.length; fl++) {
    var dItem = results[fl];
    var dItemOk = dItem.result && dItem.result.success;
    var dItemDetail = dItemOk ? ('\u2192 ' + (dItem.isF6 ? sdLocationLabel : removeLocation)) : '';
    var dItemError = '';
    if (!dItemOk) {
      if (dItem.isF6) {
        var flf6rOk = dItem.f6RemoveResult && dItem.f6RemoveResult.success;
        dItemError = flf6rOk
          ? (dItem.f6AddResult ? (dItem.f6AddResult.error || parseWaspError(dItem.f6AddResult.response, 'Add', dItem.sku)) : 'Add skipped')
          : (dItem.f6RemoveResult ? (dItem.f6RemoveResult.error || parseWaspError(dItem.f6RemoveResult.response, 'Remove', dItem.sku)) : '');
      } else {
        dItemError = dItem.result ? (dItem.result.error || parseWaspError(dItem.result.response, 'Remove', dItem.sku)) : '';
      }
    }
    sdFlowItems.push({
      sku: dItem.sku,
      qty: dItem.quantity,
      uom: dItem.uom || '',
      detail: dItemDetail,
      status: dItem.skipped ? 'Picked' : (dItemOk ? 'Complete' : 'Failed'),
      error: dItemError,
      qtyColor: dItem.isF6 ? 'orange' : 'red'
    });
  }
  logFlowDetail(flowLabel, sdExecId, {
    ref: soRef,
    detail: results.length + ' item' + (results.length !== 1 ? 's' : '') + ' \u2192 ' + sdLocationLabel,
    status: sdStatus === 'success' ? 'Complete' : sdStatus === 'failed' ? 'Failed' : 'Partial',
    linkText: soRef,
    linkUrl: getKatanaWebUrl('so', soId)
  }, sdFlowItems);

  sendSlackNotification(
    'SO Delivered: ' + soRef + '\n' +
    'Flow: ' + flowLabel + '\n' +
    'Items: ' + results.length + '\n' +
    'Location: ' + sdLocationLabel
  );

  return {
    status: 'processed',
    soId: soId,
    flow: flowLabel,
    location: removeLocation,
    results: results
  };
}

// ============================================
// F4: MANUFACTURING ORDER DONE
// ============================================

/**
 * Extract batch/lot number from a Katana MO ingredient (recipe row)
 * Checks multiple possible field structures for batch allocation data
 * Returns batch number string, or null if not batch-tracked
 */
function extractKatanaBatchNumber_(node) {
  if (!node) return '';
  return node.batch_number || node.batchNumber || node.batch_nr || node.batchNr ||
    node.lot_number || node.lotNumber || node.number || node.nr || '';
}

function extractKatanaExpiryDate_(node) {
  if (!node) return '';
  return normalizeBusinessDate_(
    node.expiration_date || node.expiry_date || node.best_before_date ||
    node.expirationDate || node.expiryDate || node.bestBeforeDate ||
    node.date_code || node.dateCode || ''
  );
}

function extractIngredientBatchNumber(ing) {
  // Check batch_transactions array (Katana batch allocation)
  var bt = ing.batch_transactions || ing.batchTransactions || [];
  if (bt.length > 0) {
    var batchNum = extractKatanaBatchNumber_(bt[0]);
    if (batchNum) return batchNum;

    // If batch_transactions has batch_id but no number, try batch_stock reference
    var batchStock = bt[0].batch_stock || bt[0].batchStock || null;
    if (batchStock) {
      var bsNum = extractKatanaBatchNumber_(batchStock);
      if (bsNum) return bsNum;
    }

    // Katana API returns batch_id + quantity only — resolve via batch_stocks endpoint
    var batchId = bt[0].batch_id || bt[0].batchId || null;
    if (batchId) {
      var batchInfo = fetchKatanaBatchStock(batchId);
      if (batchInfo) {
        var resolvedNum = extractKatanaBatchNumber_(batchInfo);
        if (resolvedNum) return resolvedNum;
      }
    }
  }

  // Check direct batch fields on the ingredient row
  var directBatch = extractKatanaBatchNumber_(ing);
  if (directBatch) return directBatch;

  // Check stock_allocations or allocations
  var allocs = ing.stock_allocations || ing.allocations || [];
  for (var a = 0; a < allocs.length; a++) {
    var allocBatch = extractKatanaBatchNumber_(allocs[a]) ||
      extractKatanaBatchNumber_(allocs[a].batch_stock || allocs[a].batchStock || allocs[a].batch || null);
    if (allocBatch) return allocBatch;
  }

  // Check picked_batches (if Katana tracks which batches were picked)
  var picked = ing.picked_batches || ing.pickedBatches || [];
  for (var p = 0; p < picked.length; p++) {
    var pickedBatch = extractKatanaBatchNumber_(picked[p]) ||
      extractKatanaBatchNumber_(picked[p].batch_stock || picked[p].batchStock || picked[p].batch || null);
    if (pickedBatch) return pickedBatch;
  }

  return null;
}

/**
 * Extract expiry/best-before date from Katana MO ingredient batch_transactions.
 * Checks: direct fields, then batch_transactions array, then stock_allocations.
 * Returns ISO date string or null.
 */
function extractIngredientExpiryDate(ing) {
  // Direct fields
  var direct = extractKatanaExpiryDate_(ing);
  if (direct) return direct;

  // Check batch_transactions (Katana MO ingredient include)
  var bt = ing.batch_transactions || ing.batchTransactions || [];
  for (var b = 0; b < bt.length; b++) {
    var btExpiry = extractKatanaExpiryDate_(bt[b]) ||
      extractKatanaExpiryDate_(bt[b].batch_stock || bt[b].batchStock || null);
    if (btExpiry) return btExpiry;

    // Katana API returns batch_id only — resolve expiry via batch_stocks endpoint
    var btId = bt[b].batch_id || bt[b].batchId || null;
    if (btId) {
      var btInfo = fetchKatanaBatchStock(btId);
      if (btInfo) {
        var resolvedExpiry = extractKatanaExpiryDate_(btInfo);
        if (resolvedExpiry) return resolvedExpiry;
      }
    }
  }

  // Check stock_allocations
  var allocs = ing.stock_allocations || ing.allocations || [];
  for (var a = 0; a < allocs.length; a++) {
    var allocExpiry = extractKatanaExpiryDate_(allocs[a]) ||
      extractKatanaExpiryDate_(allocs[a].batch_stock || allocs[a].batchStock || allocs[a].batch || null);
    if (allocExpiry) return allocExpiry;
  }

  // Check picked_batches
  var picked = ing.picked_batches || ing.pickedBatches || [];
  for (var p = 0; p < picked.length; p++) {
    var pickedExpiry = extractKatanaExpiryDate_(picked[p]) ||
      extractKatanaExpiryDate_(picked[p].batch_stock || picked[p].batchStock || picked[p].batch || null);
    if (pickedExpiry) return pickedExpiry;
  }

  return null;
}

function resolveMOOutputBatchData_(mo, variantId, moOrderNo) {
  mo = mo || {};
  var outputBatchNode = mo.output_batch || mo.outputBatch || null;
  var headerBatchTransactions = mo.batch_transactions || mo.batchTransactions || [];
  var resolved = {
    outputBatchNode: outputBatchNode,
    headerBatchTransactions: headerBatchTransactions,
    lot: '',
    expiry: '',
    batchId: null,
    source: ''
  };

  resolved.lot =
    mo.batch_number ||
    mo.output_batch_number ||
    mo.lot_number ||
    mo.batch ||
    extractKatanaBatchNumber_(outputBatchNode) ||
    '';
  resolved.expiry = extractKatanaExpiryDate_(outputBatchNode || mo) || '';
  resolved.batchId = outputBatchNode
    ? (outputBatchNode.batch_id || outputBatchNode.batchId || outputBatchNode.id || null)
    : null;
  if (resolved.lot || resolved.expiry || resolved.batchId) resolved.source = 'mo_header';

  if ((!resolved.lot || !resolved.expiry || !resolved.batchId) && headerBatchTransactions.length > 0) {
    var headerBt = headerBatchTransactions[0];
    var headerBtNode = headerBt.batch_stock || headerBt.batchStock || headerBt.batch || null;
    resolved.lot = resolved.lot || extractKatanaBatchNumber_(headerBt) || extractKatanaBatchNumber_(headerBtNode) || '';
    resolved.expiry = resolved.expiry || extractKatanaExpiryDate_(headerBt) || extractKatanaExpiryDate_(headerBtNode) || '';
    resolved.batchId = resolved.batchId || headerBt.batch_id || headerBt.batchId || headerBt.id || null;
    if (resolved.lot || resolved.expiry || resolved.batchId) resolved.source = resolved.source || 'header_batch_transactions';
  }

  if (resolved.batchId && (!resolved.lot || !resolved.expiry)) {
    var outputBatchRef = fetchKatanaBatchStock(resolved.batchId);
    if (outputBatchRef) {
      resolved.lot = resolved.lot || extractKatanaBatchNumber_(outputBatchRef) || '';
      resolved.expiry = resolved.expiry || extractKatanaExpiryDate_(outputBatchRef) || '';
      if (resolved.lot || resolved.expiry) resolved.source = 'batch_stocks_lookup';
    }
  }

  if (!resolved.lot && moOrderNo) {
    var batchMatch = String(moOrderNo).match(/[\(\[]([^\)\]]+)[\)\]]/);
    if (batchMatch) {
      var parts = batchMatch[1].trim().split(/\s+/);
      resolved.lot = parts[0] || '';
      if (parts[1] && !resolved.expiry) {
        var expParts = parts[1].split('/');
        if (expParts.length === 2) {
          var expMonth = expParts[0];
          var expYear = expParts[1].length === 2 ? '20' + expParts[1] : expParts[1];
          var lastDay = new Date(Number(expYear), Number(expMonth), 0).getDate();
          resolved.expiry = expYear + '-' + expMonth + '-' + ('0' + lastDay).slice(-2);
        }
      }
      if (resolved.lot) resolved.source = resolved.source || 'order_no_fallback';
    }
  }

  if (variantId && resolved.lot) {
    var batchInfo = resolveKatanaVariantBatchByLot_(variantId, resolved.lot);
    if (batchInfo) {
      resolved.batchId = resolved.batchId || batchInfo.id || null;
      if (batchInfo.expiry) {
        resolved.expiry = batchInfo.expiry;
        resolved.source = 'variant_batch_lookup';
      }
    }
  }

  resolved.expiry = normalizeBusinessDate_(resolved.expiry || '');
  return resolved;
}

function parseKatanaBooleanFlag_(value) {
  if (value === true || value === false) return value;
  if (value === 1 || value === 0) return !!value;
  if (typeof value !== 'string') return null;
  var text = value.trim().toLowerCase();
  if (!text) return null;
  if (text === 'true' || text === 'yes' || text === '1') return true;
  if (text === 'false' || text === 'no' || text === '0') return false;
  return null;
}

function getKatanaBatchTrackingFlag_(node) {
  if (!node || typeof node !== 'object') return null;
  var fields = [
    node.batch_tracking,
    node.batchTracking,
    node.batch_tracked,
    node.batchTracked,
    node.is_batch_tracked,
    node.isBatchTracked,
    node.track_batches,
    node.trackBatches
  ];
  for (var i = 0; i < fields.length; i++) {
    var parsed = parseKatanaBooleanFlag_(fields[i]);
    if (parsed !== null) return parsed;
  }
  return null;
}

function isKatanaBatchTrackedVariant_(variant) {
  var v = variant && variant.data ? variant.data : variant;
  if (!v) return false;

  var directFlag = getKatanaBatchTrackingFlag_(v);
  if (directFlag !== null) return directFlag;

  if (v.product) {
    var embeddedFlag = getKatanaBatchTrackingFlag_(v.product);
    if (embeddedFlag !== null) return embeddedFlag;
  }

  if (v.product_id) {
    var productData = fetchKatanaProduct(v.product_id);
    var product = productData && productData.data ? productData.data : productData;
    var productFlag = getKatanaBatchTrackingFlag_(product);
    if (productFlag !== null) return productFlag;
  }

  if (v.material) {
    var embeddedMaterialFlag = getKatanaBatchTrackingFlag_(v.material);
    if (embeddedMaterialFlag !== null) return embeddedMaterialFlag;
  }

  if (v.material_id) {
    var materialData = fetchKatanaMaterial(v.material_id);
    var material = materialData && materialData.data ? materialData.data : materialData;
    var materialFlag = getKatanaBatchTrackingFlag_(material);
    if (materialFlag !== null) return materialFlag;
  }

  return false;
}

function resolveKatanaVariantBatchByLot_(variantId, lotNumber) {
  if (!variantId || !lotNumber) return null;
  var result = katanaApiCall('batch_stocks?variant_id=' + variantId + '&include_deleted=true');
  if (!result) return null;
  var batches = result.data || result || [];
  var wantedLot = String(lotNumber).trim().toUpperCase();

  for (var i = 0; i < batches.length; i++) {
    var batch = batches[i];
    var batchLot = batch.batch_number || batch.batch_nr || batch.nr || batch.number || '';
    if (String(batchLot).trim().toUpperCase() !== wantedLot) continue;
    return {
      id: batch.batch_id || batch.id || null,
      lot: batchLot || '',
      expiry: normalizeBusinessDate_(extractKatanaExpiryDate_(batch) || '')
    };
  }

  return null;
}

function getCachedKatanaVariant_(variantId, variantCache) {
  if (!variantId) return null;
  variantCache = variantCache || {};
  if (variantCache.hasOwnProperty(variantId)) return variantCache[variantId];
  var variantData = fetchKatanaVariant(variantId);
  var variant = variantData && variantData.data ? variantData.data : variantData;
  variantCache[variantId] = variant || null;
  return variantCache[variantId];
}

function assessMOCompletionReadiness_(mo, ingredients, variantId, moOrderNo) {
  var variantCache = {};
  var pendingBatchIngredientCount = 0;
  var batchTrackedIngredientCount = 0;

  ingredients = ingredients || [];
  for (var i = 0; i < ingredients.length; i++) {
    var ing = ingredients[i] || {};
    var ingQty = parseFloat(
      ing.total_consumed_quantity ||
      ing.consumed_quantity ||
      ing.actual_quantity ||
      ing.quantity || 0
    ) || 0;
    if (ingQty <= 0) continue;

    var bt = ing.batch_transactions || ing.batchTransactions || [];
    var ingLot = extractIngredientBatchNumber(ing);
    var ingExpiry = extractIngredientExpiryDate(ing);
    var batchTracked = bt.length > 0 || !!ingLot || !!ingExpiry;

    if (!batchTracked && ing.variant_id) {
      var ingVariant = getCachedKatanaVariant_(ing.variant_id, variantCache);
      batchTracked = isKatanaBatchTrackedVariant_(ingVariant);
    }

    if (!batchTracked) continue;

    batchTrackedIngredientCount++;
    if (bt.length === 0 && !ingLot && !ingExpiry) {
      pendingBatchIngredientCount++;
    }
  }

  var outputBatchTracked = false;
  if (variantId) {
    var outputVariant = getCachedKatanaVariant_(variantId, variantCache);
    outputBatchTracked = isKatanaBatchTrackedVariant_(outputVariant);
  }

  var outputBatchData = resolveMOOutputBatchData_(mo, variantId, moOrderNo);
  var outputReady = !outputBatchTracked || !!(outputBatchData.lot || outputBatchData.batchId);

  return {
    ready: outputReady && pendingBatchIngredientCount === 0,
    outputReady: outputReady,
    outputBatchTracked: outputBatchTracked,
    outputBatchData: outputBatchData,
    batchTrackedIngredientCount: batchTrackedIngredientCount,
    pendingBatchIngredientCount: pendingBatchIngredientCount
  };
}

function waitForMOCompletionReadiness_(moId, initialMO, variantId, initialOrderNo) {
  var maxWaitMs = 15000;
  var pollMs = 2500;
  var waitedMs = 0;
  var mo = initialMO || null;
  var moOrderNo = initialOrderNo || (mo ? (mo.order_no || '') : '');
  var ingredients = [];
  var readiness = null;

  while (true) {
    if (!mo) {
      var moData = fetchKatanaMO(moId);
      mo = moData && moData.data ? moData.data : moData;
      if (!mo) break;
    }

    moOrderNo = mo.order_no || moOrderNo || ('MO-' + moId);
    var ingredientsData = fetchKatanaMOIngredients(moId);
    ingredients = ingredientsData && ingredientsData.data ? ingredientsData.data : (ingredientsData || []);
    readiness = assessMOCompletionReadiness_(mo, ingredients, variantId, moOrderNo);

    if (readiness.ready || waitedMs >= maxWaitMs) break;

    Utilities.sleep(pollMs);
    waitedMs += pollMs;

    var refreshedMOData = fetchKatanaMO(moId);
    mo = refreshedMOData && refreshedMOData.data ? refreshedMOData.data : refreshedMOData;
    if (!mo) break;
  }

  return {
    mo: mo,
    ingredients: ingredients,
    waitedMs: waitedMs,
    readiness: readiness || {
      ready: false,
      outputReady: false,
      outputBatchTracked: false,
      outputBatchData: resolveMOOutputBatchData_(mo || {}, variantId, moOrderNo),
      batchTrackedIngredientCount: 0,
      pendingBatchIngredientCount: 0
    }
  };
}

/**
 * Check if an MO already has a completed F4 entry in Activity.
 * Partial/failed rows are intentionally reprocessable.
 */
function isMOAlreadyCompleted(moRef) {
  try {
    var rawMoRef = String(moRef || '').trim();
    var searchRef = extractCanonicalActivityRef_(rawMoRef) || rawMoRef;
    if (!searchRef) return false;

    var activitySheet = getActivitySheet();
    var lastRow = activitySheet.getLastRow();
    if (lastRow <= 3) return false;

    // Read A:E so status can distinguish Complete vs Partial/Failed/Reverted.
    var data = activitySheet.getRange(4, 1, lastRow - 3, 5).getValues();

    for (var i = data.length - 1; i >= 0; i--) {
      if (!data[i][0]) continue;
      if (String(data[i][2]) !== 'F4 Manufacturing') continue;
      var details = String(data[i][3]);
      if (details.indexOf(searchRef) < 0) continue;
      var statusText = String(data[i][4] || '').trim();
      if (details.indexOf('reverted') >= 0 || statusText === 'Reverted') return false;
      if (statusText === 'Complete') return true;
    }
    return false;
  } catch (e) {
    Logger.log('isMOAlreadyCompleted error: ' + e.message);
    return false;
  }
}

function buildMORefCandidates_(moRef, moId) {
  var props = PropertiesService.getScriptProperties();
  var candidates = [];
  var indexedRef = moId ? String(props.getProperty('mo_id_ref_' + moId) || '').trim() : '';
  var rawRef = String(moRef || '').trim();
  var canonicalRef = extractCanonicalActivityRef_(rawRef, 'MO-', moId || '') || '';

  if (indexedRef) candidates.push(indexedRef);
  if (canonicalRef && candidates.indexOf(canonicalRef) < 0) candidates.push(canonicalRef);
  if (rawRef && candidates.indexOf(rawRef) < 0) candidates.push(rawRef);
  if (!candidates.length && moId) candidates.push('MO-' + moId);
  return candidates;
}

function resolveMOStoredRef_(moRef, moId) {
  var candidates = buildMORefCandidates_(moRef, moId);
  return candidates.length ? candidates[0] : String(moRef || '').trim();
}

function getMOSnapshotRaw_(moRef, moId) {
  try {
    var props = PropertiesService.getScriptProperties();
    var candidates = buildMORefCandidates_(moRef, moId);
    for (var i = 0; i < candidates.length; i++) {
      var ref = candidates[i];
      var raw = props.getProperty('mo_snapshot_' + ref);
      if (raw) return { ref: ref, raw: raw };
    }
    // Fallback: search all snapshots for one whose moId matches.
    // Handles the case where mo_id_ref_{moId} was not saved.
    if (moId) {
      var all = props.getProperties();
      for (var k in all) {
        if (k.indexOf('mo_snapshot_') === 0 && all[k]) {
          try {
            var obj = JSON.parse(all[k]);
            if (obj && String(obj.moId) === String(moId)) {
              var foundRef = k.replace('mo_snapshot_', '');
              // Also save the id_ref now so future lookups are fast
              props.setProperty('mo_id_ref_' + moId, foundRef);
              return { ref: foundRef, raw: all[k] };
            }
          } catch (pe) {}
        }
      }
    }
    return { ref: candidates.length ? candidates[0] : String(moRef || '').trim(), raw: '' };
  } catch (e) {
    Logger.log('getMOSnapshotRaw error: ' + e.message);
    return { ref: String(moRef || '').trim(), raw: '' };
  }
}

/**
 * Get map of ingredient SKUs already consumed for an MO on a previous run.
 * Used for smart retry — prevents double-consumption when MO is reprocessed.
 * Returns object { sku: count, ... } or empty object.
 */
function getMOConsumedMap(moRef, moId) {
  try {
    var props = PropertiesService.getScriptProperties();
    var keyRef = resolveMOStoredRef_(moRef, moId);
    var raw = props.getProperty('mo_consumed_' + keyRef);
    if (!raw) return {};
    return JSON.parse(raw);
  } catch (e) {
    Logger.log('getMOConsumedMap error: ' + e.message);
    return {};
  }
}

/**
 * Save the map of successfully consumed ingredient SKUs for an MO.
 * Persists across retries so partial MOs don't double-consume on reprocess.
 * Pass empty object to delete the tracking record.
 */
function saveMOConsumedMap(moRef, consumedMap, moId) {
  try {
    var props = PropertiesService.getScriptProperties();
    var keyRef = resolveMOStoredRef_(moRef, moId);
    var hasAny = false;
    for (var k in consumedMap) {
      if (consumedMap.hasOwnProperty(k) && consumedMap[k] > 0) {
        hasAny = true;
        break;
      }
    }
    if (!hasAny) {
      props.deleteProperty('mo_consumed_' + keyRef);
    } else {
      props.setProperty('mo_consumed_' + keyRef, JSON.stringify(consumedMap));
    }
  } catch (e) {
    Logger.log('saveMOConsumedMap error: ' + e.message);
  }
}

function getMOSnapshot(moRef, moId) {
  try {
    var snapshotLookup = getMOSnapshotRaw_(moRef, moId);
    var raw = snapshotLookup.raw;
    if (!raw) return null;
    var parsed = JSON.parse(raw);
    if (!parsed || typeof parsed !== 'object') return null;
    if (!Array.isArray(parsed.ingredients)) parsed.ingredients = [];
    return parsed;
  } catch (e) {
    Logger.log('getMOSnapshot error: ' + e.message);
    return null;
  }
}

function saveMOSnapshot(moRef, moId, snapshotObj) {
  try {
    var props = PropertiesService.getScriptProperties();
    var hasIngredients = snapshotObj && Array.isArray(snapshotObj.ingredients) && snapshotObj.ingredients.length > 0;
    var hasOutput = snapshotObj && snapshotObj.output && snapshotObj.output.sku;
    if (!hasIngredients && !hasOutput) {
      logWebhookQueue({ action: 'mo_snapshot.diag', moRef: moRef, moId: moId }, { status: 'diag', message: 'saveMOSnapshot: SKIPPED SAVE — no ingredients and no output in snapshot' });
      props.deleteProperty('mo_snapshot_' + moRef);
      props.deleteProperty('mo_id_ref_' + (moId || ''));
      return;
    }
    props.setProperty('mo_snapshot_' + moRef, JSON.stringify(snapshotObj));
    // Save id→ref mapping so NOT_STARTED webhook can find snapshot by Katana ID.
    // Use moId from param, fall back to moId stored in snapshot object.
    var saveId = moId || (snapshotObj && snapshotObj.moId) || null;
    if (saveId) props.setProperty('mo_id_ref_' + saveId, moRef);
    logWebhookQueue({ action: 'mo_snapshot.diag', moRef: moRef, moId: moId }, { status: 'diag', message: 'saveMOSnapshot: SAVED — ingredients:' + (snapshotObj.ingredients || []).length + ' output:' + (snapshotObj.output ? snapshotObj.output.sku : 'none') + ' key=mo_snapshot_' + moRef });
  } catch (e) {
    logWebhookQueue({ action: 'mo_snapshot.diag', moRef: moRef, moId: moId }, { status: 'error', message: 'saveMOSnapshot ERROR: ' + e.message });
  }
}

function acquireExecutionGuard_(guardKey, ttlMs) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (e) {
    // Lock timeout — another execution is holding it (likely F2 batch pre-fix,
    // or another handler). Treat as if guard is already held; caller will skip.
    Logger.log('acquireExecutionGuard_ lock timeout for ' + guardKey + ': ' + e.message);
    return false;
  }
  try {
    var props = PropertiesService.getScriptProperties();
    var existingRaw = props.getProperty(guardKey);
    if (existingRaw) {
      try {
        var existing = JSON.parse(existingRaw);
        var startedAt = parseInt(existing.startedAt || '0', 10) || 0;
        if (startedAt && (Date.now() - startedAt) < ttlMs) {
          return false;
        }
      } catch (parseErr) {
        // Stale/corrupt guard — replace it.
      }
    }
    props.setProperty(guardKey, JSON.stringify({ startedAt: Date.now() }));
    return true;
  } finally {
    lock.releaseLock();
  }
}

function releaseExecutionGuard_(guardKey) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (e) {
    // Lock timeout — guard will expire via its own TTL.
    Logger.log('releaseExecutionGuard_ lock timeout for ' + guardKey + ': ' + e.message);
    return;
  }
  try {
    PropertiesService.getScriptProperties().deleteProperty(guardKey);
  } finally {
    lock.releaseLock();
  }
}

function hasExecutionGuard_(guardKey) {
  try {
    return !!PropertiesService.getScriptProperties().getProperty(guardKey);
  } catch (e) {
    Logger.log('hasExecutionGuard error: ' + e.message);
    return false;
  }
}

function waitForExecutionGuardToClear_(guardKey, delayStepsMs) {
  var delays = delayStepsMs || [0, 1000, 2000];
  var waitedMs = 0;
  var cleared = !hasExecutionGuard_(guardKey);
  if (cleared) return { waitedMs: waitedMs, cleared: true };

  for (var i = 0; i < delays.length; i++) {
    var delayMs = parseInt(delays[i], 10) || 0;
    if (delayMs > 0) {
      Utilities.sleep(delayMs);
      waitedMs += delayMs;
    }
    if (!hasExecutionGuard_(guardKey)) {
      cleared = true;
      break;
    }
  }

  return { waitedMs: waitedMs, cleared: cleared };
}

function getPOReceiveRecentCacheKey_(poId) {
  return 'po_receive_recent_' + poId;
}

function markPOReceiveRecent_(poId, ttlSeconds) {
  try {
    CacheService.getScriptCache().put(getPOReceiveRecentCacheKey_(poId), '1', ttlSeconds || 20);
  } catch (e) {
    Logger.log('markPOReceiveRecent error: ' + e.message);
  }
}

function hasRecentPOReceive_(poId) {
  try {
    return !!CacheService.getScriptCache().get(getPOReceiveRecentCacheKey_(poId));
  } catch (e) {
    Logger.log('hasRecentPOReceive error: ' + e.message);
    return false;
  }
}

function stabilizePOStateAfterRecentReceive_(poId, poData, rowsData, delayStepsMs) {
  var delays = delayStepsMs || [0, 2000, 4000];
  var waitedMs = 0;
  var latestPO = poData;
  var latestRows = rowsData;

  for (var i = 0; i < delays.length; i++) {
    var delayMs = parseInt(delays[i], 10) || 0;
    if (delayMs > 0) {
      Utilities.sleep(delayMs);
      waitedMs += delayMs;
    }

    var freshPO = fetchKatanaPO(poId);
    if (freshPO) latestPO = freshPO.data ? freshPO.data : freshPO;

    var freshRows = fetchKatanaPORows(poId);
    if (freshRows) latestRows = freshRows;

    if (!hasRecentPOReceive_(poId) && delayMs > 0) break;
  }

  return { poData: latestPO, rowsData: latestRows, waitedMs: waitedMs };
}

function extractPORowReceivedDate_(row) {
  row = row || {};
  return row.received_date || row.receivedDate || row.receive_date || row.receiveDate || '';
}

function isKatanaPOReceivedRow_(row) {
  row = row || {};
  if (extractPORowReceivedDate_(row)) return true;

  var bt = row.batch_transactions || row.batchTransactions || [];
  for (var i = 0; i < bt.length; i++) {
    if ((parseFloat((bt[i] && bt[i].quantity) || 0) || 0) > 0) return true;
  }

  var rowStatus = String(row.status || row.row_status || '').toUpperCase();
  return rowStatus === 'RECEIVED';
}

function getPORowReceivedStockQty_(row, poStatus, options) {
  row = row || {};
  options = options || {};
  var allowPoStatusFallback = options.allowPoStatusFallback !== false;
  var rowPuomRate = parseFloat(row.purchase_uom_conversion_rate) || 1;
  var rowReceivedRaw = row.received_quantity;
  if (rowReceivedRaw === undefined || rowReceivedRaw === null || rowReceivedRaw === '') rowReceivedRaw = row.receivedQuantity;
  if (rowReceivedRaw === undefined || rowReceivedRaw === null || rowReceivedRaw === '') rowReceivedRaw = row.received_qty;

  var rowReceivedQty = parseFloat(rowReceivedRaw);
  if (isNaN(rowReceivedQty)) {
    rowReceivedQty = 0;
  }

  var poStatusText = String(poStatus || '').toUpperCase();
  if ((rowReceivedRaw === undefined || rowReceivedRaw === null || rowReceivedRaw === '') && rowReceivedQty <= 0) {
    var rowStatus = String(row.status || row.row_status || '').toUpperCase();
    if ((allowPoStatusFallback && poStatusText === 'RECEIVED') || rowStatus === 'RECEIVED' || extractPORowReceivedDate_(row)) {
      rowReceivedQty = parseFloat(row.quantity || row.qty || 0) || 0;
    }
  }

  return rowReceivedQty * rowPuomRate;
}

function assessPOReceiveReadiness_(rows, previouslySaved, poStatus) {
  rows = rows || [];
  previouslySaved = previouslySaved || [];

  var previousBatchIds = {};
  var previousNonBatchRowIds = {};
  var previousNonBatchQtyByRow = {};
  for (var i = 0; i < previouslySaved.length; i++) {
    var prev = previouslySaved[i] || {};
    if (prev.batchId) previousBatchIds[String(prev.batchId)] = true;
    if (!prev.batchId && prev.poRowId) {
      var prevRowKey = String(prev.poRowId);
      previousNonBatchRowIds[prevRowKey] = true;
      previousNonBatchQtyByRow[prevRowKey] = (previousNonBatchQtyByRow[prevRowKey] || 0) + (parseFloat(prev.qty || 0) || 0);
    }
  }

  var receivableCount = 0;
  for (var r = 0; r < rows.length; r++) {
    var row = rows[r] || {};
    var bt = row.batch_transactions || row.batchTransactions || [];
    var rowId = row.id || row.purchase_order_row_id || row.po_row_id || null;

    if (bt.length > 0) {
      for (var b = 0; b < bt.length; b++) {
        var btEntry = bt[b] || {};
        var batchId = btEntry.batch_id || btEntry.batchId || null;
        var btQty = parseFloat(btEntry.quantity || 0) || 0;
        if (btQty > 0 && (!batchId || !previousBatchIds[String(batchId)])) {
          receivableCount++;
          break;
        }
      }
      continue;
    }

    if (!rowId) continue;
    if (isKatanaPOReceivedRow_(row) && !previousNonBatchRowIds[String(rowId)]) {
      receivableCount++;
      continue;
    }

    var cumulativeReceivedStockQty = getPORowReceivedStockQty_(row, poStatus, { allowPoStatusFallback: false });
    var previousQty = previousNonBatchQtyByRow[String(rowId)] || 0;
    if (cumulativeReceivedStockQty - previousQty > 0) {
      receivableCount++;
    }
  }

  return { ready: receivableCount > 0, receivableCount: receivableCount };
}

function waitForPOReceiveReadiness_(poId, initialPOData, initialRowsData, previouslySaved, delayStepsMs) {
  var delays = delayStepsMs || [0, 1500, 3000, 4500];
  var waitedMs = 0;
  var latestPO = initialPOData;
  var latestRowsData = initialRowsData;
  var latestRows = latestRowsData && latestRowsData.data ? latestRowsData.data : (latestRowsData || []);
  var readiness = assessPOReceiveReadiness_(latestRows, previouslySaved, latestPO && latestPO.status);

  if (readiness.ready) {
    return { poData: latestPO, rowsData: latestRowsData, waitedMs: waitedMs, readiness: readiness };
  }

  for (var i = 0; i < delays.length; i++) {
    var delayMs = parseInt(delays[i], 10) || 0;
    if (delayMs > 0) {
      Utilities.sleep(delayMs);
      waitedMs += delayMs;
    }

    var freshPOData = fetchKatanaPO(poId);
    var freshPO = freshPOData && freshPOData.data ? freshPOData.data : freshPOData;
    if (freshPO) latestPO = freshPO;

    var freshRowsData = fetchKatanaPORows(poId);
    if (freshRowsData) latestRowsData = freshRowsData;
    latestRows = latestRowsData && latestRowsData.data ? latestRowsData.data : (latestRowsData || []);

    readiness = assessPOReceiveReadiness_(latestRows, previouslySaved, latestPO && latestPO.status);
    if (readiness.ready) {
      break;
    }
  }

  return { poData: latestPO, rowsData: latestRowsData, waitedMs: waitedMs, readiness: readiness };
}

function getPOReceiveStateKeys_(poRef) {
  var keyRef = String(poRef || '').replace(/[^a-zA-Z0-9_]/g, '');
  return {
    keyRef: keyRef,
    flatKey: 'f1_recv_' + keyRef,
    blocksKey: 'f1_recv_blocks_' + keyRef,
    openRowsKey: 'f1_open_rows_' + keyRef,
    partialKey: 'f1_recv_partial_' + keyRef,
    countKey: 'f1_count_' + keyRef
  };
}

function flattenPOReceiveBlocks_(blocks) {
  blocks = blocks || [];
  var flat = [];
  for (var i = 0; i < blocks.length; i++) {
    var block = blocks[i] || {};
    var items = block.items || [];
    for (var j = 0; j < items.length; j++) {
      var item = items[j] || {};
      flat.push(item);
    }
  }
  return flat;
}

function loadPOReceiveState_(poRef) {
  var keys = getPOReceiveStateKeys_(poRef);
  var props = PropertiesService.getScriptProperties();
  var blocks = [];
  var hasBlockStorage = false;
  var openRows = {};

  var blocksRaw = props.getProperty(keys.blocksKey);
  if (blocksRaw) {
    try {
      var parsedBlocks = JSON.parse(blocksRaw);
      if (Array.isArray(parsedBlocks)) {
        blocks = parsedBlocks;
        hasBlockStorage = true;
      }
    } catch (e) {
      Logger.log('loadPOReceiveState blocks parse error: ' + e.message);
    }
  }

  if (!blocks.length) {
    var flatRaw = props.getProperty(keys.flatKey);
    if (flatRaw && flatRaw.length > 2) {
      try {
        var legacyItems = JSON.parse(flatRaw);
        if (Array.isArray(legacyItems) && legacyItems.length > 0) {
          blocks = [{
            id: 'legacy_' + keys.keyRef,
            label: String(poRef || ''),
            createdAt: '',
            receiptDates: [],
            groupIds: [],
            items: legacyItems,
            legacy: true
          }];
        }
      } catch (e2) {
        Logger.log('loadPOReceiveState flat parse error: ' + e2.message);
      }
    }
  }

  var openRowsRaw = props.getProperty(keys.openRowsKey);
  if (openRowsRaw) {
    try {
      var parsedOpenRows = JSON.parse(openRowsRaw);
      if (parsedOpenRows && typeof parsedOpenRows === 'object') {
        openRows = parsedOpenRows;
      }
    } catch (e3) {
      Logger.log('loadPOReceiveState open rows parse error: ' + e3.message);
    }
  }

  return {
    keys: keys,
    blocks: blocks,
    flat: flattenPOReceiveBlocks_(blocks),
    hasBlockStorage: hasBlockStorage,
    openRows: openRows
  };
}

function persistPOReceiveState_(poRef, blocks) {
  var keys = getPOReceiveStateKeys_(poRef);
  var props = PropertiesService.getScriptProperties();
  blocks = blocks || [];

  if (blocks.length > 0) {
    props.setProperty(keys.blocksKey, JSON.stringify(blocks));
    props.setProperty(keys.flatKey, JSON.stringify(flattenPOReceiveBlocks_(blocks)));
  } else {
    props.deleteProperty(keys.blocksKey);
    props.deleteProperty(keys.flatKey);
  }
}

function buildPOOpenRowsSnapshot_(rows) {
  rows = rows || [];
  var snapshot = {};
  for (var i = 0; i < rows.length; i++) {
    var row = rows[i] || {};
    var rowId = row.id || row.purchase_order_row_id || row.po_row_id || null;
    if (!rowId || isKatanaPOReceivedRow_(row)) continue;

    snapshot[String(rowId)] = {
      id: rowId,
      variant_id: row.variant_id || row.variantId || null,
      quantity: parseFloat(row.quantity || row.qty || 0) || 0,
      purchase_uom_conversion_rate: parseFloat(row.purchase_uom_conversion_rate) || 1,
      purchase_uom: row.purchase_uom || row.purchaseUom || '',
      uom: row.uom || '',
      internal_barcode: row.internal_barcode || row.internalBarcode || '',
      supplier_item_code: row.supplier_item_code || row.supplierItemCode || '',
      batch_transactions: row.batch_transactions || row.batchTransactions || []
    };
  }
  return snapshot;
}

function buildPOOpenRowsSignature_(rows) {
  var snapshot = buildPOOpenRowsSnapshot_(rows);
  var keys = Object.keys(snapshot).sort();
  var parts = [];
  for (var i = 0; i < keys.length; i++) {
    var row = snapshot[keys[i]] || {};
    parts.push([
      keys[i],
      row.variant_id || '',
      row.quantity || 0,
      row.purchase_uom_conversion_rate || 1
    ].join(':'));
  }
  var raw = parts.join('|') || 'empty';
  var digest = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, raw, Utilities.Charset.UTF_8);
  var hex = '';
  for (var j = 0; j < digest.length; j++) {
    var byteValue = digest[j];
    if (byteValue < 0) byteValue += 256;
    var hexPart = byteValue.toString(16);
    if (hexPart.length < 2) hexPart = '0' + hexPart;
    hex += hexPart;
  }
  return hex || 'empty';
}

function persistPOOpenRowsSnapshot_(poRef, rows) {
  var keys = getPOReceiveStateKeys_(poRef);
  var props = PropertiesService.getScriptProperties();
  var snapshot = buildPOOpenRowsSnapshot_(rows);
  if (Object.keys(snapshot).length > 0) {
    props.setProperty(keys.openRowsKey, JSON.stringify(snapshot));
  } else {
    props.deleteProperty(keys.openRowsKey);
  }
}

function buildSyntheticPOReceiveRowsFromOpenDiff_(previousOpenRows, currentRows, previouslySaved) {
  previousOpenRows = previousOpenRows || {};
  currentRows = currentRows || [];
  previouslySaved = previouslySaved || [];

  var currentOpenRows = buildPOOpenRowsSnapshot_(currentRows);
  var previousReceivedQtyByRow = {};
  for (var i = 0; i < previouslySaved.length; i++) {
    var saved = previouslySaved[i] || {};
    if (saved.batchId || !saved.poRowId) continue;
    var savedRowKey = String(saved.poRowId);
    previousReceivedQtyByRow[savedRowKey] = (previousReceivedQtyByRow[savedRowKey] || 0) + (parseFloat(saved.qty || 0) || 0);
  }

  var syntheticRows = [];
  for (var rowKey in previousOpenRows) {
    if (!previousOpenRows.hasOwnProperty(rowKey)) continue;
    var prevRow = previousOpenRows[rowKey] || {};
    var currRow = currentOpenRows[rowKey] || null;
    var prevOpenPurchaseQty = parseFloat(prevRow.quantity || 0) || 0;
    var currOpenPurchaseQty = currRow ? (parseFloat(currRow.quantity || 0) || 0) : 0;
    if (prevOpenPurchaseQty <= currOpenPurchaseQty) continue;

    var puomRate = parseFloat(prevRow.purchase_uom_conversion_rate) || 1;
    var deltaPurchaseQty = prevOpenPurchaseQty - currOpenPurchaseQty;
    var prevReceivedStockQty = previousReceivedQtyByRow[rowKey] || 0;
    var cumulativeStockQty = prevReceivedStockQty + (deltaPurchaseQty * puomRate);

    syntheticRows.push({
      id: prevRow.id,
      variant_id: prevRow.variant_id || null,
      quantity: prevOpenPurchaseQty,
      purchase_uom_conversion_rate: puomRate,
      purchase_uom: prevRow.purchase_uom || '',
      uom: prevRow.uom || '',
      internal_barcode: prevRow.internal_barcode || '',
      supplier_item_code: prevRow.supplier_item_code || '',
      received_quantity: cumulativeStockQty / puomRate,
      batch_transactions: prevRow.batch_transactions || []
    });
  }

  return {
    rows: syntheticRows,
    currentOpenRows: currentOpenRows
  };
}

function buildPOReceiptBlockFromResults_(poRefLabel, poId, results, rows, sourceAction) {
  rows = rows || [];
  results = results || [];
  var rowLookup = {};
  for (var i = 0; i < rows.length; i++) {
    var row = rows[i] || {};
    var rowId = row.id || row.purchase_order_row_id || row.po_row_id || null;
    if (rowId) rowLookup[String(rowId)] = row;
  }

  var receiptDatesMap = {};
  var groupIdsMap = {};
  var items = [];
  for (var r = 0; r < results.length; r++) {
    var res = results[r] || {};
    if (!(res.result && res.result.success)) continue;

    var sourceRow = res.poRowId ? rowLookup[String(res.poRowId)] : null;
    var receivedDate = sourceRow ? extractPORowReceivedDate_(sourceRow) : '';
    var groupId = sourceRow ? (sourceRow.group_id || sourceRow.groupId || '') : '';
    if (receivedDate) receiptDatesMap[String(receivedDate)] = true;
    if (groupId !== '' && groupId !== null && groupId !== undefined) groupIdsMap[String(groupId)] = true;

    items.push({
      sku: res.sku,
      variantId: res.variantId || null,
      poRowId: res.poRowId || null,
      batchId: res.batchId || null,
      lot: res.lot || '',
      qty: res.quantity,
      uom: res.uom || '',
      expiry: normalizeBusinessDate_(res.expiry || ''),
      location: res.location,
      site: res.site,
      receivedDate: receivedDate,
      groupId: groupId
    });
  }

  return {
    id: Utilities.getUuid(),
    label: poRefLabel,
    poId: poId,
    createdAt: new Date().toISOString(),
    sourceAction: sourceAction || '',
    receiptDates: Object.keys(receiptDatesMap),
    groupIds: Object.keys(groupIdsMap),
    items: items
  };
}

function buildCurrentPOReceiptPresence_(rows) {
  rows = rows || [];
  var currentReceivedRowIds = {};
  var currentBatchIds = {};

  for (var i = 0; i < rows.length; i++) {
    var row = rows[i] || {};
    var rowId = row.id || row.purchase_order_row_id || row.po_row_id || null;
    if (rowId && isKatanaPOReceivedRow_(row)) {
      currentReceivedRowIds[String(rowId)] = true;
    }

    var bt = row.batch_transactions || row.batchTransactions || [];
    for (var j = 0; j < bt.length; j++) {
      var batchId = bt[j].batch_id || bt[j].batchId || null;
      if (batchId) currentBatchIds[String(batchId)] = true;
    }
  }

  return {
    rowIds: currentReceivedRowIds,
    batchIds: currentBatchIds
  };
}

function planPOReceiveBlockReverts_(storedBlocks, rows) {
  storedBlocks = storedBlocks || [];
  var presence = buildCurrentPOReceiptPresence_(rows);
  var revertedBlocks = [];
  var remainingBlocks = [];

  for (var i = 0; i < storedBlocks.length; i++) {
    var block = storedBlocks[i] || {};
    var items = block.items || [];
    var blockRemaining = [];
    var blockReverted = [];

    for (var j = 0; j < items.length; j++) {
      var item = items[j] || {};
      var present = false;

      if (item.batchId && presence.batchIds[String(item.batchId)]) {
        present = true;
      } else if (item.poRowId && presence.rowIds[String(item.poRowId)]) {
        present = true;
      }

      if (present) blockRemaining.push(item);
      else blockReverted.push(item);
    }

    if (blockReverted.length > 0) {
      revertedBlocks.push({
        id: block.id || '',
        label: block.label || '',
        createdAt: block.createdAt || '',
        receiptDates: block.receiptDates || [],
        groupIds: block.groupIds || [],
        items: blockReverted
      });
    }

    if (blockRemaining.length > 0) {
      remainingBlocks.push({
        id: block.id || '',
        label: block.label || '',
        createdAt: block.createdAt || '',
        receiptDates: block.receiptDates || [],
        groupIds: block.groupIds || [],
        sourceAction: block.sourceAction || '',
        legacy: !!block.legacy,
        items: blockRemaining
      });
    }
  }

  return {
    revertedBlocks: revertedBlocks,
    remainingBlocks: remainingBlocks,
    presence: presence
  };
}

function inferPOReceiveMode_(rows, poStatus) {
  rows = rows || [];
  var receivedRows = 0;
  var pendingRows = 0;

  for (var i = 0; i < rows.length; i++) {
    var row = rows[i] || {};
    var puomRate = parseFloat(row.purchase_uom_conversion_rate) || 1;
    var orderedQty = (parseFloat(row.quantity || row.qty || 0) || 0) * puomRate;
    if (orderedQty <= 0) continue;

    var receivedQty = 0;
    if (isKatanaPOReceivedRow_(row)) {
      receivedQty = orderedQty;
    } else {
      receivedQty = getPORowReceivedStockQty_(row, poStatus, { allowPoStatusFallback: false });
    }

    if (receivedQty > 0) receivedRows++;
    if (receivedQty < orderedQty) pendingRows++;
  }

  return {
    isPartial: pendingRows > 0 && receivedRows > 0,
    receivedRows: receivedRows,
    pendingRows: pendingRows
  };
}

function waitForMOSnapshotForRevert_(moRef, moId, delayStepsMs) {
  var delays = delayStepsMs || [0, 1000, 2000];
  var waitedMs = 0;
  var lookup = getMOSnapshotRaw_(moRef, moId);
  if (lookup.raw) {
    return { lookup: lookup, waitedMs: waitedMs, found: true };
  }

  var guardKey = 'mo_done_guard_' + moId;
  for (var i = 0; i < delays.length; i++) {
    var delayMs = parseInt(delays[i], 10) || 0;
    if (delayMs > 0) {
      Utilities.sleep(delayMs);
      waitedMs += delayMs;
    }

    lookup = getMOSnapshotRaw_(moRef, moId);
    if (lookup.raw) {
      return { lookup: lookup, waitedMs: waitedMs, found: true };
    }

    if (!hasExecutionGuard_(guardKey) && delayMs > 0) {
      break;
    }
  }

  return { lookup: lookup, waitedMs: waitedMs, found: false };
}

function buildCurrentPOReceivedQtyByRow_(rows, poStatus) {
  rows = rows || [];
  var receivedQtyByRow = {};

  for (var i = 0; i < rows.length; i++) {
    var row = rows[i] || {};
    var bt = row.batch_transactions || row.batchTransactions || [];
    var rowId = row.id || row.purchase_order_row_id || row.po_row_id || null;
    if (!rowId || bt.length > 0) continue;

    if (isKatanaPOReceivedRow_(row)) {
      var orderedQty = parseFloat(row.quantity || row.qty || 0) || 0;
      var orderedPuomRate = parseFloat(row.purchase_uom_conversion_rate) || 1;
      receivedQtyByRow[String(rowId)] = orderedQty * orderedPuomRate;
      continue;
    }

    receivedQtyByRow[String(rowId)] = getPORowReceivedStockQty_(row, poStatus, { allowPoStatusFallback: false });
  }

  return receivedQtyByRow;
}

function getHotfixFlag_(flagName) {
  return !!(typeof HOTFIX_FLAGS !== 'undefined' && HOTFIX_FLAGS[flagName]);
}

function getHotfixConfirmDelayMs_() {
  return 4000;
}

function refetchKatanaEntityAfterConfirmDelay_(fetchFn, entityId) {
  Utilities.sleep(getHotfixConfirmDelayMs_());
  var data = fetchFn(entityId);
  return data && data.data ? data.data : data;
}

function refetchKatanaEntityWithAdaptiveConfirm_(fetchFn, entityId, acceptFn, delayStepsMs) {
  var delays = delayStepsMs || [0, 1000, 2000];
  var totalWaitMs = 0;
  var lastEntity = null;

  for (var i = 0; i < delays.length; i++) {
    var waitMs = parseInt(delays[i], 10) || 0;
    if (waitMs > 0) {
      Utilities.sleep(waitMs);
      totalWaitMs += waitMs;
    }

    var data = fetchFn(entityId);
    var entity = data && data.data ? data.data : data;
    if (!entity) continue;

    lastEntity = entity;
    if (!acceptFn || acceptFn(entity)) {
      return { entity: entity, waitedMs: totalWaitMs, matched: true };
    }
  }

  return { entity: lastEntity, waitedMs: totalWaitMs, matched: false };
}

function hasAnyMOConsumedEntries_(consumedMap) {
  consumedMap = consumedMap || {};
  for (var sku in consumedMap) {
    if (consumedMap.hasOwnProperty(sku) && (parseFloat(consumedMap[sku]) || 0) > 0) {
      return true;
    }
  }
  return false;
}

function hasAnyMOSnapshotEntries_(snapshotObj) {
  if (!snapshotObj || typeof snapshotObj !== 'object') return false;
  var hasIngredients = Array.isArray(snapshotObj.ingredients) && snapshotObj.ingredients.length > 0;
  var hasOutput = !!(snapshotObj.output && snapshotObj.output.sku);
  return hasIngredients || hasOutput;
}

function hasAnyF4InventoryChange_(results) {
  results = results || {};
  var removed = results.ingredientsRemoved || [];
  for (var i = 0; i < removed.length; i++) {
    var ing = removed[i];
    if (ing && ing.result && ing.result.success && !ing.skippedRetry) {
      return true;
    }
  }

  return !!(results.outputAdded && results.outputAdded.success && !results.outputAdded.skippedRetry);
}

/**
 * Check if a PO was already processed (Activity log idempotency).
 * Blocks ANY existing F1 entry for this PO — Received, Partial, or Failed.
 * Prevents duplicate processing when Katana sends repeat webhooks.
 */
function isPOAlreadyReceived(poRef) {
  try {
    var rawPoRef = String(poRef || '').trim();
    var searchRef = extractCanonicalActivityRef_(rawPoRef) || rawPoRef;

    // If a partial receive was recorded, allow re-receiving (receive all or next partial)
    var partialKey = 'f1_recv_partial_' + rawPoRef.replace(/[^a-zA-Z0-9_]/g, '');
    if (PropertiesService.getScriptProperties().getProperty(partialKey) === 'true') return false;
    if (!searchRef) return false;

    var activitySheet = getActivitySheet();
    var lastRow = activitySheet.getLastRow();
    if (lastRow <= 3) return false;

    // Read 5 columns: A(execId), B(time), C(flow), D(details), E(status)
    var data = activitySheet.getRange(4, 1, lastRow - 3, 5).getValues();

    for (var i = data.length - 1; i >= 0; i--) {
      if (!data[i][0]) continue;
      if (String(data[i][2]) !== 'F1 Receiving') continue;
      var details = String(data[i][3]);
      if (details.indexOf(searchRef) < 0) continue;
      // Most-recent F1 entry for this PO found.
      // If it was a revert, allow re-processing (PO can be re-received).
      if (details.indexOf('reverted') >= 0) return false;
      // If it failed, allow retry — nothing was added to WASP.
      var entryStatus = String(data[i][4] || '').trim();
      if (entryStatus === 'Failed') return false;
      return true;
    }
    return false;
  } catch (e) {
    Logger.log('isPOAlreadyReceived error: ' + e.message);
    return false;
  }
}

/**
 * Determine item type from Katana variant data.
 * Returns 'product', 'material', 'intermediate', or 'unknown'.
 * Checks embedded product.type first, fetches product if needed.
 */
function getKatanaItemType(variant) {
  var v = variant && variant.data ? variant.data : variant;
  if (!v) return 'unknown';

  if (v.product && v.product.type) return v.product.type;

  if (v.product_id) {
    var prod = fetchKatanaProduct(v.product_id);
    var p = prod && prod.data ? prod.data : prod;
    if (p) return p.type || p.product_type || p.category || 'unknown';
  }

  return 'unknown';
}

/**
 * Handle Manufacturing Order Done event
 * - Removes ingredients from PRODUCTION
 * - Adds finished goods to PROD-RECEIVING
 */
function handleManufacturingOrderDone(payload) {
  var moId = payload.object ? payload.object.id : null;

  if (!moId) {
    logToSheet('MO_DONE_ERROR', payload, { error: 'No MO ID found' });
    return { status: 'error', message: 'No MO ID in webhook' };
  }

  var guardKey = 'mo_done_guard_' + moId;
  if (!acquireExecutionGuard_(guardKey, 120000)) {
    return { status: 'skipped', reason: 'MO completion already in progress', moId: moId };
  }

  try {
    // Dedup: skip if this MO was already processed recently
    var cache = CacheService.getScriptCache();
    var dedupKey = 'mo_done_' + moId;
    if (cache.get(dedupKey)) {
      return { status: 'skipped', reason: 'MO already processed', moId: moId };
    }
    cache.put(dedupKey, 'true', 300); // 5-minute dedup window

    // Fetch MO header
    var moData = fetchKatanaMO(moId);
    // Removed verbose MO logging

    if (!moData) {
      return { status: 'error', message: 'Failed to fetch MO header' };
    }

    var mo = moData.data ? moData.data : moData;

    // Idempotency: check Activity log for existing completed F4 entry
    // Allows reprocessing of failed/partial MOs but blocks duplicate completions
    var moOrderNo = mo.order_no || ('MO-' + moId);
    var moRef = resolveMOStoredRef_(moOrderNo, moId);
    var existingSnapshot = getMOSnapshot(moRef, moId);
    // Block re-processing if snapshot exists (MO was completed, even partially).
    // Snapshot is only cleared by reverseMOSnapshot() — so if it's still here,
    // WASP already has the changes applied. Revert first, then re-complete.
    if (existingSnapshot) {
      logToSheet('MO_DONE_DUPLICATE', { moId: moId, moRef: moRef }, { reason: 'Snapshot exists — MO already processed, not yet reverted' });
      return { status: 'skipped', reason: 'MO already processed (snapshot exists)', moId: moId, moRef: moRef };
    }
    if (isMOAlreadyCompleted(moRef)) {
      logToSheet('MO_DONE_DUPLICATE', { moId: moId, moRef: moRef }, { reason: 'Already completed in Activity log' });
      return { status: 'skipped', reason: 'MO already completed', moId: moId, moRef: moRef };
    }

    // Skip auto-assembly (ASM) MOs — bundle products (VB-) assemble finished goods,
    // not raw ingredients at PRODUCTION. F5 (ShipStation) handles all VB- deductions
    // at shipment time, so F4 processing these creates noise with no inventory effect.
    if (moOrderNo.indexOf(' ASM ') >= 0) {
      return { status: 'ignored', reason: 'Auto-assembly MO skipped: ' + moOrderNo };
    }

    var moLocationName = String(mo.location_name || mo.manufacturing_location_name || '').trim();
    if (!moLocationName && mo.location_id && typeof fetchKatanaLocation === 'function') {
      var moLocationData = fetchKatanaLocation(mo.location_id);
      var moLocation = moLocationData && moLocationData.data ? moLocationData.data : moLocationData;
      moLocationName = moLocation ? String(moLocation.name || '').trim() : '';
    }
    if (String(moLocationName || '').toLowerCase() === 'shopify') {
      return { status: 'ignored', reason: 'Shopify-location MO skipped: ' + moOrderNo };
    }

    // Get product details
    var variantId = mo.variant_id;
    var outputSku = '';
    var productName = '';
    var productCategory = '';

    if (variantId) {
      var variantData = fetchKatanaVariant(variantId);
      var variant = variantData && variantData.data ? variantData.data : variantData;
      outputSku = variant ? (variant.sku || '') : '';
      productName = variant ? (variant.product ? variant.product.name : variant.name) : '';
      if (!productName) productName = mo.product_name || '';
      // Get category from inline product object first, then fetch separately if needed
      productCategory = variant ? (variant.product ? (variant.product.category || variant.product.category_name || '') : '') : '';
      if (!productCategory && variant && variant.product_id) {
        var productData = fetchKatanaProduct(variant.product_id);
        var product = productData && productData.data ? productData.data : productData;
        productCategory = product ? (product.category || product.category_name || '') : '';
      }
    }

    var outputUom = resolveVariantUom(variant);
    var outputQuantity = mo.actual_quantity || mo.quantity || 0;
    var outputBatchData = resolveMOOutputBatchData_(mo, variantId, moOrderNo);
    var lotNumber = outputBatchData.lot || '';
    var expiryDate = outputBatchData.expiry || '';

  // If Katana gave an output batch reference but not the embedded lot/expiry, resolve it directly.

  // Fallback: parse batch/expiry from MO order_no if API fields are empty
  // MO name patterns:
  //   "MO-7246 (UFC410B 02/29) // FEB 24"  → batch = UFC410B, expiry = 02/29 (MM/YY)
  //   "MO-7230 [UFC412A] // FEB 26"        → batch = UFC412A, expiry = +3 years
  if (!lotNumber && moOrderNo) {
    var batchMatch = moOrderNo.match(/[\(\[]([^\)\]]+)[\)\]]/);
    if (batchMatch) {
      var parts = batchMatch[1].trim().split(/\s+/);
      lotNumber = parts[0] || '';  // e.g. "UFC410B" or "UFC412A"
      if (parts[1] && !expiryDate) {
        // Parse MM/YY → YYYY-MM-DD (DD = last day of month)
        var expParts = parts[1].split('/');
        if (expParts.length === 2) {
          var expMonth = expParts[0];
          var expYear = expParts[1].length === 2 ? '20' + expParts[1] : expParts[1];
          // Last day of the expiry month
          var lastDay = new Date(Number(expYear), Number(expMonth), 0).getDate();
          expiryDate = expYear + '-' + expMonth + '-' + ('0' + lastDay).slice(-2);
        }
      }
      Logger.log('F4 batch from MO name: lot=' + lotNumber + ' exp=' + expiryDate);
    }
  }

  // Prefer the exact Katana batch_stocks expiry for this finished/output lot when available.
  if (variantId && lotNumber) {
    var outputBatchInfo = resolveKatanaVariantBatchByLot_(variantId, lotNumber);
    if (outputBatchInfo && outputBatchInfo.expiry) {
      expiryDate = outputBatchInfo.expiry;
    }
  }

  // Default expiry: +3 years from MO completion date when lot exists but no expiry
  if (lotNumber && !expiryDate) {
    var completionSrc = mo.completed_at || mo.done_at || mo.updated_at || null;
    var baseDate = completionSrc ? new Date(completionSrc) : new Date();
    var expDate = new Date(baseDate.getFullYear() + 3, baseDate.getMonth(), baseDate.getDate());
    var ey = expDate.getFullYear();
    var em = ('0' + (expDate.getMonth() + 1)).slice(-2);
    var ed = ('0' + expDate.getDate()).slice(-2);
    expiryDate = ey + '-' + em + '-' + ed;
    Logger.log('F4 expiry defaulted +3yr: ' + expiryDate);
  }

  // Removed verbose MO logging

  // Detect stage
  var stage = detectMOStage(productName, outputSku, productCategory);
  // Removed verbose MO logging

  // Wait only as long as Katana still needs to expose batch allocations.
  // This keeps F4 fast when data is already ready, while preserving the 15s safety cap.
  var completionReady = waitForMOCompletionReadiness_(moId, mo, variantId, moOrderNo);
  var confirmedMO = completionReady.mo;
  if (confirmedMO) {
    mo = confirmedMO;
    moOrderNo = mo.order_no || moOrderNo;
    moRef = resolveMOStoredRef_(moOrderNo, moId);
    outputQuantity = mo.actual_quantity || mo.quantity || outputQuantity;

    outputBatchData = completionReady.readiness && completionReady.readiness.outputBatchData
      ? completionReady.readiness.outputBatchData
      : resolveMOOutputBatchData_(mo, variantId, moOrderNo);
    lotNumber = outputBatchData.lot || lotNumber || '';
    expiryDate = outputBatchData.expiry || expiryDate || '';
  }

  var liveDoneStatus = String((mo && mo.status) || '').toUpperCase();
  if (liveDoneStatus !== 'DONE' && liveDoneStatus !== 'COMPLETED') {
    try { cache.remove(dedupKey); } catch (cacheRemoveErr) {}
    return { status: 'ignored', reason: 'Stale MO done webhook; current status is ' + (liveDoneStatus || 'UNKNOWN') };
  }

  var ingredients = completionReady.ingredients || [];
  if (completionReady.waitedMs > 0) {
    logToSheet('MO_DONE_WAIT', {
      moId: moId,
      moRef: moOrderNo,
      waitedMs: completionReady.waitedMs,
      pendingBatchIngredients: completionReady.readiness ? completionReady.readiness.pendingBatchIngredientCount : '',
      outputReady: completionReady.readiness ? completionReady.readiness.outputReady : ''
    }, 'Adaptive Katana readiness wait before F4 completion');
  }

  // Expiry pre-flight: check batch-tracked ingredients have expiry dates
  var missingExpiry = [];
  for (var ec = 0; ec < ingredients.length; ec++) {
    var ecIng = ingredients[ec];
    var ecLot = extractIngredientBatchNumber(ecIng);
    if (ecLot) {
      var ecExpiry = extractIngredientExpiryDate(ecIng);
      if (!ecExpiry) {
        var ecSku = '';
        if (ecIng.variant_id) {
          var ecVar = fetchKatanaVariant(ecIng.variant_id);
          var ecV = ecVar && ecVar.data ? ecVar.data : ecVar;
          ecSku = ecV ? (ecV.sku || '') : '';
        }
        missingExpiry.push(ecSku || ('variant_' + ecIng.variant_id));
      }
    }
  }
  if (missingExpiry.length > 0) {
    var moOrderNo2 = mo.order_no || ('MO-' + moId);
    logToSheet('MO_MISSING_EXPIRY', { moId: moId, moRef: moOrderNo2, items: missingExpiry.join(', ') });
    Logger.log('F4 WARNING: batch-tracked ingredients missing expiry in ' + moOrderNo2 + ': ' + missingExpiry.join(', '));
    // Continue processing — don't abort. WASP will use blank expiry for these items.
    // Operator should fix lot data in Katana for next time.
  }

  var results = {
    stage: stage,
    ingredientsRemoved: [],
    outputAdded: null
  };

  // Smart retry: load previously consumed SKUs for this MO (prevents double-consumption)
  var consumedMap = getMOConsumedMap(moRef, moId);
  var hadPriorCommittedState = hasAnyMOSnapshotEntries_(existingSnapshot) || hasAnyMOConsumedEntries_(consumedMap);
  // Working copy for skip-checking (decremented as we skip items)
  var retrySkipMap = {};
  for (var rk in consumedMap) {
    if (consumedMap.hasOwnProperty(rk)) retrySkipMap[rk] = consumedMap[rk];
  }

  // Remove each ingredient from PRODUCTION
  for (var i = 0; i < ingredients.length; i++) {
    var ing = ingredients[i];
    var ingSku = '';
    var ingUom = '';
    if (ing.variant_id) {
      var ingVariantData = fetchKatanaVariant(ing.variant_id);
      var ingVariant = ingVariantData && ingVariantData.data ? ingVariantData.data : ingVariantData;
      ingSku = ingVariant ? (ingVariant.sku || '') : '';
      ingUom = resolveVariantUom(ingVariant);
    }

    var ingQty = ing.total_consumed_quantity ||
                 ing.consumed_quantity ||
                 ing.actual_quantity ||
                 ing.quantity || 0;

    if (ingSku && ingQty > 0) {
      // Smart retry: skip if this ingredient was already consumed on a prior run
      if (retrySkipMap[ingSku] && retrySkipMap[ingSku] > 0) {
        retrySkipMap[ingSku]--;
        var skipLot = extractIngredientBatchNumber(ing) || '';
        var skipExpiry = normalizeBusinessDate_(extractIngredientExpiryDate(ing) || '');
        results.ingredientsRemoved.push({
          sku: ingSku,
          quantity: ingQty,
          lot: skipLot,
          expiry: skipExpiry,
          uom: ingUom,
          result: { success: true },
          skippedRetry: true
        });
        continue;
      }

      // Check for multi-lot batch_transactions (ingredient split across 2+ batches)
      var bt = ing.batch_transactions || ing.batchTransactions || [];

      if (bt.length > 1) {
        // MULTI-LOT: Process each batch_transaction separately with its own qty and lot
        var allBatchesOk = true;
        var btVarBatchCache = null; // cache variant-level lookup — shared across all btEntry iterations

        // Mark sync BEFORE WASP calls — prevents F2 echo
        markSyncedToWasp(ingSku, FLOWS.MO_INGREDIENT_LOCATION, 'remove');

        for (var btIdx = 0; btIdx < bt.length; btIdx++) {
          var btEntry = bt[btIdx];
          var btQty = btEntry.quantity || 0;
          if (btQty <= 0) continue;

          // Resolve lot and expiry from this batch_transaction
          var btLot = extractKatanaBatchNumber_(btEntry);
          var btExpiry = extractKatanaExpiryDate_(btEntry);

          // If missing, resolve via batch_stock embedded object or batch_stocks endpoint
          if (!btLot && btEntry.batch_stock) {
            btLot = extractKatanaBatchNumber_(btEntry.batch_stock);
          }
          var btBatchId = btEntry.batch_id || btEntry.batchId || null;
          if (btBatchId && (!btLot || !btExpiry)) {
            var btBatchInfo = fetchKatanaBatchStock(btBatchId);
            if (btBatchInfo) {
              btLot = btLot || extractKatanaBatchNumber_(btBatchInfo);
              btExpiry = btExpiry || extractKatanaExpiryDate_(btBatchInfo);
            }
          }

          // Variant-level fallback: batch_stocks?variant_id=X&include_deleted=true
          // Depleted batches return 404 from /batch_stocks/{id} but still appear here.
          // This mirrors the single-lot path (line ~1553) — fixes wrong-lot deductions when
          // fetchKatanaBatchStock returns null and WASP fallback picks the wrong lot.
          if ((!btLot || !btExpiry) && btBatchId && ing.variant_id) {
            if (!btVarBatchCache) {
              var btVarResult = katanaApiCall('batch_stocks?variant_id=' + ing.variant_id + '&include_deleted=true');
              btVarBatchCache = (btVarResult && (btVarResult.data || btVarResult)) || [];
            }
            for (var btvi = 0; btvi < btVarBatchCache.length; btvi++) {
              var btvBatch = btVarBatchCache[btvi];
              var btvId = btvBatch.batch_id || btvBatch.id;
              if (String(btvId) === String(btBatchId)) {
                btLot = btLot || btvBatch.batch_number || btvBatch.nr || btvBatch.number || '';
                if (!btExpiry) {
                  var btvExp = extractKatanaExpiryDate_(btvBatch);
                  btExpiry = btvExp;
                }
                logToSheet('F4_MULTI_BATCH_VAR_FALLBACK', {
                  sku: ingSku, moRef: moOrderNo, lot: btLot, expiry: btExpiry,
                  batchId: btBatchId, btIdx: btIdx
                }, 'Resolved via variant-level batch_stocks (depleted batch, multi-lot path)');
                break;
              }
            }
          }

          btExpiry = normalizeBusinessDate_(btExpiry);

          // If Katana lot unresolved, try WASP fallback: query all lots at PRODUCTION
          // for this SKU. Only use WASP lot when unambiguous (exactly one candidate after
          // filtering out lots already consumed by other batches in this MO).
          if (!btLot) {
            var btWaspFallbackLots = waspLookupAllLots_(ingSku, FLOWS.MO_INGREDIENT_LOCATION, CONFIG.WASP_SITE);
            // Filter out lots already consumed by prior batches in this same MO
            var btUsedLots = {};
            for (var btPrior = 0; btPrior < results.ingredientsRemoved.length; btPrior++) {
              var priorItem = results.ingredientsRemoved[btPrior];
              if (priorItem.sku === ingSku && priorItem.lot && priorItem.result && priorItem.result.success) {
                btUsedLots[priorItem.lot] = true;
              }
            }
            var btCandidates = [];
            for (var btfi = 0; btfi < btWaspFallbackLots.length; btfi++) {
              if (!btUsedLots[btWaspFallbackLots[btfi].lot]) {
                btCandidates.push(btWaspFallbackLots[btfi]);
              }
            }
            if (btCandidates.length === 1) {
              btLot = btCandidates[0].lot;
              btExpiry = btCandidates[0].expiry || btExpiry;
              logToSheet('F4_WASP_LOT_FALLBACK', {
                sku: ingSku, qty: btQty, moRef: moOrderNo, lot: btLot, expiry: btExpiry,
                batchId: btBatchId || '', btIdx: btIdx, source: 'wasp_single_candidate'
              }, 'Katana lot unresolved — WASP has exactly 1 remaining lot at PRODUCTION');
            } else if (btCandidates.length > 1) {
              var btQtyMatch = [];
              for (var bqm = 0; bqm < btCandidates.length; bqm++) {
                if (btCandidates[bqm].quantity >= btQty) btQtyMatch.push(btCandidates[bqm]);
              }
              if (btQtyMatch.length === 1) {
                btLot = btQtyMatch[0].lot;
                btExpiry = btQtyMatch[0].expiry || btExpiry;
                logToSheet('F4_WASP_LOT_FALLBACK', {
                  sku: ingSku, qty: btQty, moRef: moOrderNo, lot: btLot, expiry: btExpiry,
                  batchId: btBatchId || '', btIdx: btIdx, source: 'wasp_qty_match',
                  totalCandidates: btCandidates.length
                }, 'Katana lot unresolved — 1 lot with sufficient qty (' + btCandidates.length + ' total)');
              }
            }
          }
          if (!btLot) {
            var btUnresolvedReason = 'Exact Katana lot unresolved' + (btBatchId ? ' (batch_id:' + btBatchId + ')' : '');
            allBatchesOk = false;
            logToSheet('F4_BATCH_UNRESOLVED', {
              sku: ingSku, qty: btQty, moRef: moOrderNo, batchId: btBatchId || '', btIdx: btIdx
            }, 'Katana MO ingredient batch could not be resolved exactly — skipped to avoid wrong lot removal.');
            results.ingredientsRemoved.push({
              sku: ingSku, quantity: btQty, lot: '', expiry: btExpiry || '',
              batchId: btBatchId || '',
              uom: ingUom, result: { success: false, error: btUnresolvedReason }, multiBatch: true,
              multiBatchTotal: ingQty, multiBatchFirst: (btIdx === 0),
              multiBatchLast: (btIdx === bt.length - 1), strictBatchSkip: true,
              strictBatchSkipReason: btUnresolvedReason
            });
            continue;
          }

          var lotOnlyBtWasp = waspLookupLotForRemoval(
            ingSku,
            FLOWS.MO_INGREDIENT_LOCATION,
            CONFIG.WASP_SITE,
            btLot,
            btQty
          );

          if (!lotOnlyBtWasp) {
            var btSkipReason = 'Katana lot not in WASP';
            allBatchesOk = false;
            logToSheet('F4_BATCH_NOT_IN_WASP', {
              sku: ingSku, qty: btQty, katanaLot: btLot, expiry: btExpiry || '', moRef: moOrderNo, btIdx: btIdx
            }, btSkipReason + ' - skipped. No random lot fallback allowed.');
            results.ingredientsRemoved.push({
              sku: ingSku, quantity: btQty, lot: btLot, expiry: btExpiry || '',
              uom: ingUom, result: { success: false, error: btSkipReason }, multiBatch: true,
              multiBatchTotal: ingQty, multiBatchFirst: (btIdx === 0),
              multiBatchLast: (btIdx === bt.length - 1), strictBatchSkip: true,
              strictBatchSkipReason: btSkipReason
            });
            continue;
          }

          if (lotOnlyBtWasp.rowCount > 1) {
            logToSheet('F4_LOT_ONLY_MATCH', {
              sku: ingSku,
              qty: btQty,
              katanaLot: btLot,
              katanaExpiry: btExpiry || '',
              waspDateCode: lotOnlyBtWasp.dateCode || '',
              rowCount: lotOnlyBtWasp.rowCount,
              moRef: moOrderNo,
              btIdx: btIdx
            }, 'MO deduction matched by lot only; expiry ignored.');
          }
          btExpiry = lotOnlyBtWasp.dateCode || btExpiry || '';

          var btRemoveResult = waspRemoveInventoryWithLot(
            ingSku, btQty, FLOWS.MO_INGREDIENT_LOCATION, btLot,
            '[MO-SYNC] MO ' + moId + ' consumed: ' + productName + ' (lot ' + btLot + ')',
            null, btExpiry
          );

          if (!btRemoveResult.success) {
            allBatchesOk = false;
            logToSheet('F4_INGREDIENT_FAIL', {
              sku: ingSku, qty: btQty, lot: btLot || 'none',
              batchIdx: btIdx, totalBatches: bt.length
            }, JSON.stringify(btEntry).substring(0, 400));
          }

          results.ingredientsRemoved.push({
            sku: ingSku,
            quantity: btQty,
            lot: btLot || '',
            expiry: btExpiry || '',
            uom: ingUom,
            result: btRemoveResult,
            multiBatch: true,
            multiBatchTotal: ingQty,
            multiBatchFirst: (btIdx === 0),
            multiBatchLast: (btIdx === bt.length - 1)
          });
        }

        // Track successful consumption for smart retry (only if all batches succeeded)
        if (allBatchesOk) {
          consumedMap[ingSku] = (consumedMap[ingSku] || 0) + 1;
        }

      } else {
        // SINGLE-LOT or NO-LOT: existing behavior
        var ingLot = extractIngredientBatchNumber(ing);
        var ingExpiry = normalizeBusinessDate_(extractIngredientExpiryDate(ing));
        var ingBatchTracked = bt.length > 0 || !!ingLot || !!ingExpiry || isKatanaBatchTrackedVariant_(ingVariant);

        // If Katana batch resolution failed but batch_transactions exist, item IS batch-tracked.
        // Only Katana-based resolution is allowed here. Never choose a random WASP lot.
        if (!ingLot) {
          if (bt.length > 0) {
            var ingBtId = bt[0].batch_id || bt[0].batchId || null;

            // 1) Katana variant-level: batch_stocks?variant_id=X — scan for matching batch_id.
            // Depleted batches appear here with include_deleted=true even when /batch_stocks/{id} returns 404.
            if (ingBtId && ing.variant_id) {
              var ingVarResult = katanaApiCall('batch_stocks?variant_id=' + ing.variant_id + '&include_deleted=true');
              if (ingVarResult) {
                var ingVarList = ingVarResult.data || ingVarResult || [];
                for (var ivbi = 0; ivbi < ingVarList.length; ivbi++) {
                  var ivb = ingVarList[ivbi];
                  var ivbId = ivb.batch_id || ivb.id;
                  if (String(ivbId) === String(ingBtId)) {
                    ingLot = ivb.batch_number || ivb.nr || ivb.number || '';
                    if (!ingExpiry) {
                      var ivbExp = extractKatanaExpiryDate_(ivb);
                      ingExpiry = ivbExp;
                    }
                    logToSheet('F4_BATCH_KATANA_VAR_FALLBACK', {
                      sku: ingSku, moRef: moOrderNo, lot: ingLot, expiry: ingExpiry, batchId: ingBtId
                    }, 'Resolved via variant-level batch_stocks (depleted batch)');
                    break;
                  }
                }
              }
            }
          }
        }

        // WASP lot fallback for single-batch: if Katana couldn't resolve the lot,
        // query WASP for all lots at PRODUCTION. If exactly ONE lot exists, use it.
        if (ingBatchTracked && !ingLot) {
          var sbWaspFallbackLots = waspLookupAllLots_(ingSku, FLOWS.MO_INGREDIENT_LOCATION, CONFIG.WASP_SITE);
          if (sbWaspFallbackLots.length === 1) {
            ingLot = sbWaspFallbackLots[0].lot;
            ingExpiry = sbWaspFallbackLots[0].expiry || ingExpiry;
            logToSheet('F4_WASP_LOT_FALLBACK', {
              sku: ingSku, qty: ingQty, moRef: moOrderNo, lot: ingLot, expiry: ingExpiry,
              batchId: (bt.length > 0 ? (bt[0].batch_id || bt[0].batchId || '') : ''),
              source: 'wasp_single_lot'
            }, 'Katana lot unresolved — WASP has exactly 1 lot at PRODUCTION');
          } else if (sbWaspFallbackLots.length > 1) {
            var sbQtyMatch = [];
            for (var sqm = 0; sqm < sbWaspFallbackLots.length; sqm++) {
              if (sbWaspFallbackLots[sqm].quantity >= ingQty) sbQtyMatch.push(sbWaspFallbackLots[sqm]);
            }
            if (sbQtyMatch.length === 1) {
              ingLot = sbQtyMatch[0].lot;
              ingExpiry = sbQtyMatch[0].expiry || ingExpiry;
              logToSheet('F4_WASP_LOT_FALLBACK', {
                sku: ingSku, qty: ingQty, moRef: moOrderNo, lot: ingLot, expiry: ingExpiry,
                batchId: (bt.length > 0 ? (bt[0].batch_id || bt[0].batchId || '') : ''),
                source: 'wasp_qty_match',
                totalCandidates: sbWaspFallbackLots.length
              }, 'Katana lot unresolved — 1 lot with sufficient qty (' + sbWaspFallbackLots.length + ' total)');
            }
          }
        }
        if (ingBatchTracked && !ingLot) {
          var unresolvedIngBatchId = (bt.length > 0 ? (bt[0].batch_id || bt[0].batchId || '') : '');
          var ingMissingReason = 'Katana batch missing on MO ingredient' + (unresolvedIngBatchId ? ' (batch_id:' + unresolvedIngBatchId + ')' : '');
          logToSheet('F4_BATCH_UNRESOLVED', {
            sku: ingSku,
            qty: ingQty,
            moRef: moOrderNo,
            batchId: unresolvedIngBatchId,
            batchTracked: true
          }, JSON.stringify(ing).substring(0, 500));
          results.ingredientsRemoved.push({
            sku: ingSku, quantity: ingQty, lot: '', expiry: ingExpiry || '',
            batchId: unresolvedIngBatchId,
            uom: ingUom, result: { success: false, error: ingMissingReason },
            strictBatchSkip: true, strictBatchSkipReason: ingMissingReason
          });
          continue;
        }

        var lotOnlyIngWasp = null;
        if (ingLot) {
          lotOnlyIngWasp = waspLookupLotForRemoval(
            ingSku,
            FLOWS.MO_INGREDIENT_LOCATION,
            CONFIG.WASP_SITE,
            ingLot,
            ingQty
          );

          if (!lotOnlyIngWasp) {
            var ingSkipReason = 'Katana lot not in WASP';
            logToSheet('F4_BATCH_NOT_IN_WASP', {
              sku: ingSku, qty: ingQty, katanaLot: ingLot, expiry: ingExpiry || '', moRef: moOrderNo
            }, ingSkipReason + ' - skipped. No random lot fallback allowed.');
            results.ingredientsRemoved.push({
              sku: ingSku, quantity: ingQty, lot: ingLot, expiry: ingExpiry || '',
              uom: ingUom, result: { success: false, error: ingSkipReason },
              strictBatchSkip: true, strictBatchSkipReason: ingSkipReason
            });
            continue;
          }

          if (lotOnlyIngWasp.rowCount > 1) {
            logToSheet('F4_LOT_ONLY_MATCH', {
              sku: ingSku,
              qty: ingQty,
              katanaLot: ingLot,
              katanaExpiry: ingExpiry || '',
              waspDateCode: lotOnlyIngWasp.dateCode || '',
              rowCount: lotOnlyIngWasp.rowCount,
              moRef: moOrderNo
            }, 'MO deduction matched by lot only; expiry ignored.');
          }
          ingExpiry = lotOnlyIngWasp.dateCode || ingExpiry || '';
        }

        var removeResult;

        // Mark sync BEFORE WASP call — prevents F2 echo
        markSyncedToWasp(ingSku, FLOWS.MO_INGREDIENT_LOCATION, 'remove');

        if (ingLot) {
          removeResult = waspRemoveInventoryWithLot(
            ingSku,
            ingQty,
            FLOWS.MO_INGREDIENT_LOCATION,
            ingLot,
            '[MO-SYNC] MO ' + moId + ' consumed: ' + productName,
            null,
            ingExpiry
          );
        } else {
          removeResult = waspRemoveInventory(
            ingSku,
            ingQty,
            FLOWS.MO_INGREDIENT_LOCATION,
            '[MO-SYNC] MO ' + moId + ' consumed: ' + productName
          );
        }

        if (!removeResult.success) {
          logToSheet('F4_INGREDIENT_FAIL', {
            sku: ingSku, qty: ingQty, lot: ingLot || 'none',
            ingKeys: Object.keys(ing).join(',')
          }, JSON.stringify(ing).substring(0, 800));
        }

        results.ingredientsRemoved.push({
          sku: ingSku,
          quantity: ingQty,
          lot: ingLot || '',
          expiry: ingExpiry || '',
          uom: ingUom,
          result: removeResult
        });

        if (removeResult.success) {
          consumedMap[ingSku] = (consumedMap[ingSku] || 0) + 1;
        }
      }

    } else {
      // Skipped - missing SKU or zero qty
    }
  }

  // Batch mismatch retry: if any items were skipped due to Katana lot propagation delay,
  // wait 20 s, re-fetch fresh ingredient data, and retry just those items.
  var batchMismatchSkips = [];
  for (var bmi = 0; bmi < results.ingredientsRemoved.length; bmi++) {
    if (results.ingredientsRemoved[bmi].batchMismatchSkip) {
      batchMismatchSkips.push(results.ingredientsRemoved[bmi]);
    }
  }
  if (batchMismatchSkips.length > 0) {
    logToSheet('F4_BATCH_RETRY_START', {
      moRef: moOrderNo, skippedCount: batchMismatchSkips.length
    }, 'Waiting 20s then re-fetching MO ingredients for batch retry');

    Utilities.sleep(20000);

    var retryIngrData = fetchKatanaMOIngredients(moId);
    var retryIngrs = retryIngrData && retryIngrData.data ? retryIngrData.data : (retryIngrData || []);

    // Build skipped-SKU lookup: { sku -> skippedEntry }
    var skippedSkuMap = {};
    for (var smi = 0; smi < batchMismatchSkips.length; smi++) {
      skippedSkuMap[batchMismatchSkips[smi].sku] = batchMismatchSkips[smi];
    }

    for (var rri = 0; rri < retryIngrs.length; rri++) {
      var retryIng = retryIngrs[rri];
      var retrySku = '';
      if (retryIng.variant_id) {
        var retryVarData = fetchKatanaVariant(retryIng.variant_id);
        var retryVar = retryVarData && retryVarData.data ? retryVarData.data : retryVarData;
        retrySku = retryVar ? (retryVar.sku || '') : '';
      }

      if (!retrySku || !skippedSkuMap[retrySku]) continue;

      var skippedEntry = skippedSkuMap[retrySku];
      var retryLot = extractIngredientBatchNumber(retryIng);
      var retryExpiry = normalizeBusinessDate_(extractIngredientExpiryDate(retryIng));

      if (!retryLot) {
        logToSheet('F4_BATCH_RETRY_NO_LOT', {
          sku: retrySku, moRef: moOrderNo
        }, 'Retry: still no lot in fresh Katana data — staying skipped');
        continue;
      }

      markSyncedToWasp(retrySku, FLOWS.MO_INGREDIENT_LOCATION, 'remove');

      var retryLotWasp = waspLookupLotForRemoval(
        retrySku,
        FLOWS.MO_INGREDIENT_LOCATION,
        CONFIG.WASP_SITE,
        retryLot,
        skippedEntry.quantity
      );
      if (retryLotWasp && retryLotWasp.rowCount > 1) {
        logToSheet('F4_LOT_ONLY_MATCH', {
          sku: retrySku,
          qty: skippedEntry.quantity,
          katanaLot: retryLot,
          katanaExpiry: retryExpiry || '',
          waspDateCode: retryLotWasp.dateCode || '',
          rowCount: retryLotWasp.rowCount,
          moRef: moOrderNo
        }, 'MO retry deduction matched by lot only; expiry ignored.');
      }
      retryExpiry = (retryLotWasp && retryLotWasp.dateCode) ? retryLotWasp.dateCode : (retryExpiry || '');

      var retryRemoveResult = waspRemoveInventoryWithLot(
        retrySku,
        skippedEntry.quantity,
        FLOWS.MO_INGREDIENT_LOCATION,
        retryLot,
        '[MO-SYNC] MO ' + moId + ' consumed (batch retry): ' + productName,
        null,
        retryExpiry || ''
      );

      // Update the skipped entry in-place so activity log reflects the retry outcome
      for (var rei = 0; rei < results.ingredientsRemoved.length; rei++) {
        var rEntry = results.ingredientsRemoved[rei];
        if (rEntry.batchMismatchSkip && rEntry.sku === retrySku) {
          rEntry.lot = retryLot;
          rEntry.expiry = retryExpiry || '';
          rEntry.result = retryRemoveResult;
          rEntry.batchMismatchSkip = false;
          rEntry.batchMismatchRetried = true;
          break;
        }
      }

      if (retryRemoveResult.success) {
        consumedMap[retrySku] = (consumedMap[retrySku] || 0) + 1;
        logToSheet('F4_BATCH_RETRY_OK', {
          sku: retrySku, qty: skippedEntry.quantity, oldLot: skippedEntry.lot, newLot: retryLot, moRef: moOrderNo
        }, 'Batch retry succeeded with updated lot from Katana');
      } else {
        logToSheet('F4_BATCH_RETRY_FAIL', {
          sku: retrySku, qty: skippedEntry.quantity, lot: retryLot, moRef: moOrderNo
        }, retryRemoveResult.response ? retryRemoveResult.response.substring(0, 400) : 'unknown error');
      }
    }
  }

  // Save consumed map for smart retry on future partial reprocessing
  saveMOConsumedMap(moRef, consumedMap, moId);

  // Add output to WASP
  var outputLocation = (stage === 'FINISHED')
    ? FLOWS.MO_OUTPUT_LOCATION
    : FLOWS.MO_INGREDIENT_LOCATION;
  var existingOutputSnapshot = existingSnapshot && existingSnapshot.output ? existingSnapshot.output : null;

  if (outputSku && outputQuantity > 0) {
    if (existingOutputSnapshot && existingOutputSnapshot.sku) {
      var existingQty = parseFloat(existingOutputSnapshot.qty || 0) || 0;
      var currentQty = parseFloat(outputQuantity || 0) || 0;
      var sameOutputShape =
        String(existingOutputSnapshot.sku || '') === String(outputSku || '') &&
        String(existingOutputSnapshot.location || '') === String(outputLocation || '') &&
        String(existingOutputSnapshot.lot || '') === String(lotNumber || '') &&
        String(existingOutputSnapshot.expiry || '') === String(expiryDate || '') &&
        existingQty === currentQty;

      logToSheet(
        sameOutputShape ? 'MO_OUTPUT_ALREADY_COMMITTED' : 'MO_OUTPUT_ALREADY_COMMITTED_MISMATCH',
        {
          moId: moId,
          moRef: moOrderNo,
          sku: outputSku,
          qty: outputQuantity,
          location: outputLocation,
          existingQty: existingOutputSnapshot.qty || '',
          existingLocation: existingOutputSnapshot.location || ''
        },
        sameOutputShape
          ? 'Skipping duplicate output add on MO retry'
          : 'Skipping output add because prior snapshot already committed a different output shape'
      );

      results.outputAdded = {
        success: true,
        skippedRetry: true,
        lot: existingOutputSnapshot.lot || lotNumber || '',
        expiry: existingOutputSnapshot.expiry || expiryDate || '',
        location: existingOutputSnapshot.location || outputLocation
      };
    } else {
      // Mark sync BEFORE WASP call — prevents F2 echo
      markSyncedToWasp(outputSku, outputLocation, 'add');

      if (lotNumber || expiryDate) {
        results.outputAdded = waspAddInventoryWithLot(
          outputSku,
          outputQuantity,
          outputLocation,
          lotNumber,
          expiryDate,
          '[MO-SYNC] MO ' + moId + ' output: ' + productName
        );
      } else {
        results.outputAdded = waspAddInventory(
          outputSku,
          outputQuantity,
          outputLocation,
          '[MO-SYNC] MO ' + moId + ' output: ' + productName
        );
      }
    }
  }

  if (hadPriorCommittedState && !hasAnyF4InventoryChange_(results)) {
    logToSheet('MO_DONE_NOOP_RETRY', {
      moId: moId,
      moRef: moOrderNo,
      outputSku: outputSku,
      outputQty: outputQuantity,
      stage: stage
    }, 'Repeat MO webhook caused no new WASP changes; suppressing duplicate F4 activity row');
    return {
      status: 'skipped',
      reason: 'MO retry caused no new WASP changes',
      moId: moId,
      moRef: moOrderNo,
      stage: stage,
      outputSku: outputSku,
      outputQuantity: outputQuantity,
      outputLocation: outputLocation,
      ingredientLocation: FLOWS.MO_INGREDIENT_LOCATION,
      ingredientsRemoved: results.ingredientsRemoved.length,
      results: results,
      noopRetry: true
    };
  }

  // Activity Log — F4 Manufacturing
  var f4Success = 0;
  var f4Fail = 0;
  var f4Skip = 0;
  var f4SubItems = [];

  // Ingredient sub-items
  for (var m = 0; m < results.ingredientsRemoved.length; m++) {
    var ing = results.ingredientsRemoved[m];
    var ingOk = ing.result && ing.result.success;

    // Treat -46002 (insufficient quantity) as skip, not failure
    // Katana already consumed these — WASP just doesn't have stock at PRODUCTION
    var isInsufficientQty = !ingOk && ing.result && ing.result.response &&
      ing.result.response.indexOf('-46002') >= 0;
    var isBatchMismatchSkip = !!ing.batchMismatchSkip;
    var isStrictBatchSkip = !!ing.strictBatchSkip;

    if (ingOk) {
      f4Success++;
    } else if (isInsufficientQty || isBatchMismatchSkip || isStrictBatchSkip) {
      f4Success++; // Count as success — Katana is source of truth
      f4Skip++;
    } else {
      f4Fail++;
    }

    var ingAction = '';
    var ingStatus = '';
    var ingErrorMsg = '';

    if (ing.skippedRetry) {
      ingAction = buildActivityActionText_('consume from ' + getActivityDisplayLocation_(FLOWS.MO_INGREDIENT_LOCATION), ing.lot, ing.expiry);
      ingStatus = 'Skipped';
      ingErrorMsg = 'Already consumed';
    } else if (isStrictBatchSkip) {
      ingAction = joinActivitySegments_([
        'consume from ' + getActivityDisplayLocation_(FLOWS.MO_INGREDIENT_LOCATION),
        ing.batchId && !ing.lot ? 'batch_id:' + ing.batchId : '',
        buildActivityBatchMeta_(ing.lot, ing.expiry)
      ]);
      ingStatus = 'Skipped';
      ingErrorMsg = ing.strictBatchSkipReason || 'Exact Katana lot skipped';
    } else if (isBatchMismatchSkip) {
      ingAction = 'consume from ' + getActivityDisplayLocation_(FLOWS.MO_INGREDIENT_LOCATION);
      ingStatus = 'Skipped';
      ingErrorMsg = 'Batch mismatch (stale lot)';
    } else if (isInsufficientQty) {
      ingAction = buildActivityActionText_('consume from ' + getActivityDisplayLocation_(FLOWS.MO_INGREDIENT_LOCATION), ing.lot, ing.expiry);
      ingStatus = 'Skipped';
      ingErrorMsg = 'Not enough in ' + FLOWS.MO_INGREDIENT_LOCATION;
    } else if (ingOk) {
      ingAction = buildActivityActionText_('consume from ' + getActivityDisplayLocation_(FLOWS.MO_INGREDIENT_LOCATION), ing.lot, ing.expiry);
      ingStatus = 'Consumed';
    } else {
      ingAction = buildActivityActionText_('consume from ' + getActivityDisplayLocation_(FLOWS.MO_INGREDIENT_LOCATION), ing.lot, ing.expiry);
      ingStatus = 'Failed';
      if (ing.result) {
        ingErrorMsg = ing.result.error || (ing.result.response ? parseWaspError(ing.result.response, 'Remove', ing.sku, FLOWS.MO_INGREDIENT_LOCATION) : '');
      }
    }
    var ingCompactAction = buildActivityCompactMeta_(
      CONFIG.WASP_SITE,
      FLOWS.MO_INGREDIENT_LOCATION,
      ing.lot,
      ing.expiry,
      [ing.batchId && !ing.lot ? 'batch_id:' + ing.batchId : '']
    );

    if (ing.multiBatch && ing.multiBatchFirst) {
      // Multi-batch parent row: grey qty showing total, "(N batches)" suffix
      // Only show if at least one sub-batch was not skipped (item not in WASP → hide entirely)
      var batchCount = 0;
      var anySubBatchShown = false;
      for (var bc = m; bc < results.ingredientsRemoved.length; bc++) {
        if (results.ingredientsRemoved[bc].sku === ing.sku && results.ingredientsRemoved[bc].multiBatch) {
          batchCount++;
          if (!results.ingredientsRemoved[bc].batchMismatchSkip) anySubBatchShown = true;
        } else if (bc > m) break;
      }
      if (anySubBatchShown) {
        f4SubItems.push({
          sku: ing.sku,
          qty: ing.multiBatchTotal,
          uom: normalizeUom(ing.uom || ''),
          success: true,
          status: '',
          error: '',
          action: buildActivityCompactMeta_(CONFIG.WASP_SITE, FLOWS.MO_INGREDIENT_LOCATION, '', '', []),
          qtyColor: 'grey',
          isParent: true,
          batchCount: batchCount
        });
      }
    }

    if (ing.multiBatch) {
      // Nested batch sub-row: skip silently if item/batch not found in WASP
      if (!isBatchMismatchSkip) {
        var nestedAction = '';
        if (ing.lot) nestedAction += 'lot:' + ing.lot;
        if (ing.expiry) nestedAction += (nestedAction ? '  ' : '') + 'exp:' + normalizeBusinessDate_(ing.expiry);
        f4SubItems.push({
          sku: '',
          qty: ing.quantity,
          uom: normalizeUom(ing.uom || ''),
          success: ingOk || isInsufficientQty,
          status: ingStatus,
          error: ingErrorMsg,
          action: nestedAction,
          qtyColor: 'red',
          nested: true
        });
      }
    } else {
      // Single-lot: skip silently if item not found in WASP
      if (!isBatchMismatchSkip) {
        f4SubItems.push({
          sku: ing.sku,
          qty: ing.quantity,
          uom: normalizeUom(ing.uom || ''),
          success: ingOk || isInsufficientQty,
          status: ingStatus,
          error: ingErrorMsg,
          action: ingCompactAction,
          qtyColor: 'red'
        });
      }
    }
  }

  // Output sub-item
  if (results.outputAdded && outputSku) {
    var outOk = results.outputAdded.success;
    var outSkipOk = !!results.outputAdded.skippedRetry;
    var outLot = results.outputAdded.lot || lotNumber || '';
    var outExpiry = results.outputAdded.expiry || expiryDate || '';
    var outLocation = results.outputAdded.location || outputLocation;
    if (outOk) f4Success++; else f4Fail++;
    var outError = '';
    var outAction = buildActivityCompactMeta_(CONFIG.WASP_SITE, outLocation, outLot, outExpiry);
    if (outSkipOk) {
      outError = 'Already produced';
    } else if (!outOk && results.outputAdded) {
      outError = results.outputAdded.error || (results.outputAdded.response ? parseWaspError(results.outputAdded.response, 'Add', outputSku) : '');
    }
    f4SubItems.push({
      sku: outputSku,
      qty: outputQuantity,
      uom: outputUom,
      success: outOk,
      status: outOk ? (outSkipOk ? 'Skipped' : 'Produced') : 'Failed',
      error: outError,
      action: outAction,
      qtyColor: 'green'
    });
  }

  var f4Status = (f4Fail === 0 && f4Skip === 0) ? 'success' : (f4Success === 0 ? 'failed' : 'partial');
  var f4Detail = joinActivitySegments_([moRef, outputSku + ' x' + outputQuantity]);
  outputLocation = getActivityDisplayLocation_(outputLocation);
  var f4ExecId = logActivity('F4', f4Detail, f4Status, buildActivitySourceActionContext_('Katana', f4Status === 'success' ? 'build' : 'build partial', CONFIG.WASP_SITE), f4SubItems, {
    text: moRef,
    url: getKatanaWebUrl('mo', moId)
  });

  // Log to F4 detail tab
  var f4FlowItems = [];
  for (var fl = 0; fl < results.ingredientsRemoved.length; fl++) {
    var fIng = results.ingredientsRemoved[fl];
    var fIngOk = fIng.result && fIng.result.success;
    // Check for -46002 (insufficient qty) — same logic as Activity tab
    var fIngInsufficient = !fIngOk && fIng.result && fIng.result.response &&
      fIng.result.response.indexOf('-46002') >= 0;
    var fIngStrictSkip = !!fIng.strictBatchSkip;
    var fIngError = '';
    var fIngDetail = '';
    var fIngStatus = '';
    if (fIng.skippedRetry) {
      fIngDetail = FLOWS.MO_INGREDIENT_LOCATION;
      if (fIng.lot) fIngDetail += '  lot:' + fIng.lot;
      if (fIng.expiry) fIngDetail += '  exp:' + normalizeBusinessDate_(fIng.expiry);
      fIngStatus = 'Skip-OK';
      fIngError = 'Already consumed';
    } else if (fIngStrictSkip) {
      fIngDetail = FLOWS.MO_INGREDIENT_LOCATION;
      if (!fIng.lot && fIng.batchId) fIngDetail += '  batch_id:' + fIng.batchId;
      if (fIng.lot) fIngDetail += '  lot:' + fIng.lot;
      if (fIng.expiry) fIngDetail += '  exp:' + fIng.expiry;
      fIngStatus = 'Skipped';
      fIngError = fIng.strictBatchSkipReason || 'Exact Katana lot skipped';
    } else if (fIngInsufficient) {
      fIngDetail = FLOWS.MO_INGREDIENT_LOCATION;
      fIngStatus = 'Skipped';
      fIngError = 'Not enough in ' + FLOWS.MO_INGREDIENT_LOCATION;
    } else if (fIngOk) {
      fIngDetail = FLOWS.MO_INGREDIENT_LOCATION;
      if (fIng.lot) fIngDetail += '  lot:' + fIng.lot;
      if (fIng.expiry) fIngDetail += '  exp:' + fIng.expiry;
      fIngStatus = 'Consumed';
    } else {
      fIngDetail = FLOWS.MO_INGREDIENT_LOCATION;
      fIngStatus = 'Failed';
      if (fIng.result) {
        fIngError = fIng.result.error || (fIng.result.response ? parseWaspError(fIng.result.response, 'Remove', fIng.sku, FLOWS.MO_INGREDIENT_LOCATION) : '');
      }
    }
    if (fIng.multiBatch && fIng.multiBatchFirst) {
      // Multi-batch parent row
      var fBatchCount = 0;
      for (var fbc = fl; fbc < results.ingredientsRemoved.length; fbc++) {
        if (results.ingredientsRemoved[fbc].sku === fIng.sku && results.ingredientsRemoved[fbc].multiBatch) fBatchCount++;
        else if (fbc > fl) break;
      }
      f4FlowItems.push({
        sku: fIng.sku,
        qty: fIng.multiBatchTotal,
        uom: normalizeUom(fIng.uom || ''),
        detail: FLOWS.MO_INGREDIENT_LOCATION,
        status: '',
        error: '',
        qtyColor: 'grey',
        isParent: true,
        batchCount: fBatchCount
      });
    }

    if (fIng.multiBatch) {
      // Nested batch sub-row
      var fNestedDetail = '';
      if (fIng.lot) fNestedDetail += 'lot:' + fIng.lot;
      if (fIng.expiry) fNestedDetail += (fNestedDetail ? '  ' : '') + 'exp:' + fIng.expiry;
      f4FlowItems.push({
        sku: '',
        qty: fIng.quantity,
        uom: normalizeUom(fIng.uom || ''),
        detail: fNestedDetail,
        status: fIngStatus,
        error: fIngError,
        qtyColor: 'red',
        nested: true
      });
    } else {
      f4FlowItems.push({
        sku: fIng.sku,
        qty: fIng.quantity,
        uom: normalizeUom(fIng.uom || ''),
        detail: fIngDetail,
        status: fIngStatus,
        error: fIngError,
        qtyColor: 'red'
      });
    }
  }
  if (results.outputAdded && outputSku) {
    var fOutOk = results.outputAdded.success;
    var fOutSkipOk = !!results.outputAdded.skippedRetry;
    var fOutLot = results.outputAdded.lot || lotNumber || '';
    var fOutExpiry = results.outputAdded.expiry || expiryDate || '';
    var fOutLocation = results.outputAdded.location || outputLocation;
    var fOutError = '';
    var fOutDetail = fOutLocation;
    if (fOutLot) fOutDetail += '  lot:' + fOutLot;
    if (fOutExpiry) fOutDetail += '  exp:' + normalizeBusinessDate_(fOutExpiry);
    if (fOutSkipOk) {
      fOutError = 'Already produced';
    } else if (!fOutOk && results.outputAdded) {
      fOutError = results.outputAdded.error || (results.outputAdded.response ? parseWaspError(results.outputAdded.response, 'Add', outputSku) : '');
    }
    f4FlowItems.push({
      sku: outputSku,
      qty: outputQuantity,
      uom: outputUom,
      detail: fOutDetail,
      status: fOutOk ? (fOutSkipOk ? 'Skip-OK' : 'Produced') : 'Failed',
      error: fOutError,
      qtyColor: 'green'
    });
  }
  // Save reversal snapshot BEFORE Slack — sendSlackNotification can throw and kill
  // the execution, leaving no snapshot for revert. Snapshot must be written first.
  var snapshotIng = existingSnapshot && Array.isArray(existingSnapshot.ingredients)
    ? existingSnapshot.ingredients.slice()
    : [];
  for (var si = 0; si < results.ingredientsRemoved.length; si++) {
    var sIng = results.ingredientsRemoved[si];
    if (sIng.result && sIng.result.success && !sIng.skippedRetry) {
      snapshotIng.push({ sku: sIng.sku, qty: sIng.quantity, uom: sIng.uom || '', lot: sIng.lot || '', expiry: sIng.expiry || '' });
    }
  }
  var snapshotOut = existingOutputSnapshot || null;
  if (!snapshotOut && results.outputAdded && results.outputAdded.success && outputSku && !results.outputAdded.skippedRetry) {
    snapshotOut = { sku: outputSku, qty: outputQuantity, uom: outputUom || '', lot: lotNumber || '', expiry: expiryDate || '', location: outputLocation };
  }
  // TEMP DIAG: write directly to Script Properties to confirm code is reached
  try {
    PropertiesService.getScriptProperties().setProperty(
      'SNAP_DIAG_' + (moRef || 'unknown'),
      'ing:' + snapshotIng.length + '|out:' + (snapshotOut ? snapshotOut.sku : 'null') + '|moId:' + moId + '|t:' + new Date().getTime()
    );
  } catch (diagErr) {}
  if (snapshotIng.length > 0 || snapshotOut) {
    var snapshotObj = { moRef: moRef, moId: moId, stage: stage, ingredients: snapshotIng, output: snapshotOut, completedAt: new Date().getTime() };
    saveMOSnapshot(moRef, moId, snapshotObj);
  }

  // Clean up consumed tracking if MO fully completed (no need for retry)
  if (f4Status === 'success') {
    saveMOConsumedMap(moRef, {});
  }

  logFlowDetail('F4', f4ExecId, {
    ref: moRef,
    detail: outputSku + ' x' + outputQuantity + ' → ' + outputLocation,
    status: f4Status === 'success' ? 'Complete' : f4Status === 'failed' ? 'Failed' : 'Partial',
    linkText: moRef,
    linkUrl: getKatanaWebUrl('mo', moId)
  }, f4FlowItems);

  try {
    sendSlackNotification(
      'MO Done: ' + moRef + '\n' +
      'Stage: ' + stage + '\n' +
      'Output: ' + outputSku + ' x' + outputQuantity + ' → ' + outputLocation + '\n' +
      'Ingredients removed: ' + results.ingredientsRemoved.length
    );
  } catch (slackErr) {
    Logger.log('sendSlackNotification error (non-fatal): ' + slackErr.message);
  }

    return {
      status: 'processed',
      moId: moId,
      stage: stage,
      outputSku: outputSku,
      outputQuantity: outputQuantity,
      outputLocation: outputLocation,
      ingredientLocation: FLOWS.MO_INGREDIENT_LOCATION,
      ingredientsRemoved: results.ingredientsRemoved.length,
      results: results
    };
  } finally {
    releaseExecutionGuard_(guardKey);
  }
}

/**
 * Detect MO stage based on Katana product category.
 * INTERMEDIATE PRODUCT → output stays in PRODUCTION (MO_INGREDIENT_LOCATION)
 * Everything else (FINISHED GOODS, EQUIPMENT, RAW MATERIALS, unknown) → PROD-RECEIVING (MO_OUTPUT_LOCATION)
 */
function detectMOStage(productName, sku, category) {
  var cat = (category || '').toUpperCase().trim();
  if (cat === 'INTERMEDIATE PRODUCT') return 'INTERMEDIATE';
  return 'FINISHED';
}

// ============================================
// F6: MANUFACTURING ORDER IN_PROGRESS (Ingredient Staging)
// ============================================
// When MO moves to IN_PROGRESS, move ingredients from
// RECEIVING-DOCK → PRODUCTION so F4 can remove them on MO DONE.
// ============================================

/**
 * Handle Manufacturing Order Updated event
 * Primary state-change router for MOs.
 * - DONE payloads fall back to F4 completion when Katana does not send .done reliably
 * - Non-DONE payloads with an F4 snapshot trigger F4 revert
 * - IN_PROGRESS without an F4 snapshot continues to the legacy F6 staging path
 */
function handleManufacturingOrderUpdated(payload) {
  var moId = payload.object ? payload.object.id : null;

  if (!moId) {
    return { status: 'error', message: 'No MO ID in webhook' };
  }

  var payloadObj = payload.object || {};
  var payloadStatus = String(payloadObj.status || '').toUpperCase();
  var doneByPayload = (payloadStatus === 'DONE' || payloadStatus === 'COMPLETED');
  var nonDoneByPayload = !!payloadStatus && !doneByPayload;

  // Katana sometimes surfaces MO completion only via manufacturing_order.updated.
  // Treat the webhook payload status as the first source of truth, and let
  // handleManufacturingOrderDone() enforce idempotency if the separate .done event
  // also arrives.
  if (doneByPayload) {
    return handleManufacturingOrderDone(payload);
  }

  // Check for F4 snapshot BEFORE the IN_PROGRESS staging branch.
  // Snapshot means MO was previously completed (Done) and is now being reverted —
  // this applies to ANY non-DONE status including IN_PROGRESS (Work in progress),
  // NOT_STARTED, BLOCKED, PARTIALLY_COMPLETE, etc.
  var moRefRev = resolveMOStoredRef_(payloadObj.order_no || ('MO-' + moId), moId);
  var snapshotLookupRev = getMOSnapshotRaw_(moRefRev, moId);
  if (!snapshotLookupRev.raw && nonDoneByPayload) {
    var pendingSnapshot = waitForMOSnapshotForRevert_(moRefRev, moId, [0, 1000, 2000]);
    if (pendingSnapshot.waitedMs > 0) {
      logToSheet('MO_REVERT_SNAPSHOT_WAIT', {
        moId: moId,
        moRef: moRefRev,
        payloadStatus: payloadStatus,
        waitedMs: pendingSnapshot.waitedMs,
        found: pendingSnapshot.found
      }, 'Adaptive wait for F4 snapshot before revert');
    }
    snapshotLookupRev = pendingSnapshot.lookup || snapshotLookupRev;
  }
  var snapshotStrRev = snapshotLookupRev.raw;
  moRefRev = snapshotLookupRev.ref || moRefRev;
  if (snapshotStrRev && nonDoneByPayload) {
    // Guard against concurrent reverts — same MO getting multiple NOT_STARTED webhooks
    // simultaneously. Without this, all concurrent executions restore ingredients
    // multiple times before any one of them deletes the snapshot.
    var revertGuardKey = 'mo_revert_guard_' + moId;
    if (!acquireExecutionGuard_(revertGuardKey, 60000)) {
      return { status: 'skipped', reason: 'MO revert already in progress', moId: moId };
    }
    // Re-check snapshot after acquiring guard — another execution may have already reverted
    var snapshotRecheckRev = getMOSnapshotRaw_(moRefRev, moId);
    if (!snapshotRecheckRev.raw) {
      releaseExecutionGuard_(revertGuardKey);
      return { status: 'skipped', reason: 'MO revert already completed by concurrent execution', moId: moId };
    }
    snapshotStrRev = snapshotRecheckRev.raw;
    moRefRev = snapshotRecheckRev.ref || moRefRev;

    var triggerStatus = payloadStatus;
    if (getHotfixFlag_('F4_CONFIRM_STATUS_REVERT')) {
      var adaptiveConfirm = refetchKatanaEntityWithAdaptiveConfirm_(
        fetchKatanaMO,
        moId,
        function(entity) {
          var status = String((entity && entity.status) || '').toUpperCase();
          return !!status && status !== 'DONE' && status !== 'COMPLETED';
        },
        [0, 2000, 4000, 8000]
      );
      var confirmedMO = adaptiveConfirm ? adaptiveConfirm.entity : null;
      if (!confirmedMO) {
        var revResult0 = reverseMOSnapshot(moId, moRefRev, snapshotStrRev, 'status_change:' + triggerStatus);
        releaseExecutionGuard_(revertGuardKey);
        return revResult0;
      }

      var confirmedStatus = String(confirmedMO.status || '').toUpperCase();
      if (adaptiveConfirm && adaptiveConfirm.waitedMs > 0) {
        logToSheet('MO_REVERT_CONFIRM_WAIT', {
          moId: moId,
          moRef: moRefRev,
          payloadStatus: triggerStatus,
          confirmedStatus: confirmedStatus,
          waitedMs: adaptiveConfirm.waitedMs
        }, 'Adaptive F4 revert confirmation wait');
      }
      // If Katana API still returns DONE after extended retries but webhook
      // explicitly says non-DONE: Katana API propagation is slow. The webhook
      // payload is authoritative — a snapshot exists proving the MO was completed,
      // and the webhook explicitly carries a non-DONE status. Proceed with revert.
      if (confirmedStatus === 'DONE' || confirmedStatus === 'COMPLETED') {
        logToSheet('MO_REVERT_API_SLOW_PROCEEDING', {
          moId: moId,
          payloadStatus: triggerStatus,
          confirmedStatus: confirmedStatus,
          moRef: moRefRev,
          waitedMs: adaptiveConfirm.waitedMs
        }, 'Katana API still returns DONE after ' + adaptiveConfirm.waitedMs + 'ms but webhook says ' + triggerStatus + ' — trusting webhook, proceeding with revert');
      }

      if (confirmedStatus && confirmedStatus !== triggerStatus) {
        logToSheet('MO_REVERT_STATUS_MISMATCH', {
          moId: moId,
          payloadStatus: triggerStatus,
          confirmedStatus: confirmedStatus,
          moRef: moRefRev
        }, 'Payload status disagrees with MO refetch; using webhook payload for F4 revert branch');
      }

      snapshotLookupRev = getMOSnapshotRaw_(confirmedMO.order_no || moRefRev, moId);
      moRefRev = snapshotLookupRev.ref || moRefRev;
      snapshotStrRev = snapshotLookupRev.raw || snapshotStrRev;
    }
    var revResult = reverseMOSnapshot(moId, moRefRev, snapshotStrRev, 'status_change:' + triggerStatus);
    releaseExecutionGuard_(revertGuardKey);
    return revResult;
  }

  // Fetch the live MO after handling the payload-driven DONE/revert paths.
  var moData = fetchKatanaMO(moId);
  if (!moData) {
    return { status: 'error', message: 'Failed to fetch MO' };
  }

  var mo = moData.data ? moData.data : moData;
  var moStatus = String(mo.status || '').toUpperCase();

  // If payload status was blank or Katana sent an inconsistent update, fall back to
  // completion only when both payload and refetch do not clearly indicate a non-DONE state.
  if ((moStatus === 'DONE' || moStatus === 'COMPLETED') && !nonDoneByPayload) {
    return handleManufacturingOrderDone(payload);
  }

  // Treat IN_PROGRESS the same as NOT_STARTED unless we are reversing a completed MO.
  // Operators often move an MO back to Work in progress to edit qty, ingredients, lot,
  // or expiry before completing it again. That should be a clean pre-start/edit state,
  // not a separate inventory-moving flow.
  if (moStatus === 'IN_PROGRESS') {
    return { status: 'ignored', reason: 'MO IN_PROGRESS treated as pre-start/edit state until DONE' };
  }

  // All other non-DONE states are also no-ops unless an F4 snapshot existed above.
  return { status: 'ignored', reason: 'MO status ' + moStatus + ' has no action unless reverting a completed MO' };
}

// ============================================
// F4 REVERSAL: MO reverted or deleted
// ============================================

/**
 * Handle Manufacturing Order Deleted event.
 * Full reversal of F4 WASP changes using stored snapshot.
 * Triggered by manufacturing_order.deleted webhook.
 */
function handleManufacturingOrderDeleted(payload) {
  var moId = payload.object ? payload.object.id : null;
  if (!moId) {
    return { status: 'error', message: 'No MO ID in webhook' };
  }

  // Deleted MOs cannot be fetched from Katana — use secondary index to find moRef
  var scriptProps = PropertiesService.getScriptProperties();
  var moRef = scriptProps.getProperty('mo_id_ref_' + moId);

  if (!moRef) {
    // Try fetching (may briefly succeed after deletion)
    var moData = fetchKatanaMO(moId);
    if (moData) {
      var mo = moData.data ? moData.data : moData;
      moRef = mo.order_no || ('MO-' + moId);
    } else {
      moRef = 'MO-' + moId;
    }
  }

  var deletedLookup = getMOSnapshotRaw_(moRef, moId);
  moRef = deletedLookup.ref || moRef;
  var snapshotStr = deletedLookup.raw;
  if (!snapshotStr) {
    Logger.log('F4-REVERT: No snapshot for ' + moRef + ' (deleted MO ' + moId + ') — nothing to reverse');
    return { status: 'ignored', reason: 'No F4 snapshot for ' + moRef };
  }

  return reverseMOSnapshot(moId, moRef, snapshotStr, 'deleted');
}

/**
 * Reverse F4 WASP changes using a stored snapshot.
 * Called by handleManufacturingOrderUpdated (status change away from DONE)
 * and handleManufacturingOrderDeleted (full deletion).
 *
 * @param {number|string} moId - Katana MO ID
 * @param {string} moRef - MO order number (e.g. "MO-7309")
 * @param {string} snapshotStr - JSON string of stored snapshot
 * @param {string} trigger - What triggered the reversal (for logging)
 */
var REVERT_WINDOW_DAYS = 14;

function reverseMOSnapshot(moId, moRef, snapshotStr, trigger) {
  var snapshot;
  try {
    snapshot = JSON.parse(snapshotStr);
  } catch (e) {
    return { status: 'error', message: 'Invalid snapshot JSON for ' + moRef };
  }

  // Enforce revert window — block stale reverts where inventory has likely drifted
  if (snapshot.completedAt) {
    var ageMs = new Date().getTime() - snapshot.completedAt;
    var ageDays = ageMs / (1000 * 60 * 60 * 24);
    if (ageDays > REVERT_WINDOW_DAYS) {
      logWebhookQueue(
        { action: 'mo_revert.blocked', moRef: moRef, moId: moId },
        { status: 'skipped', message: 'Revert blocked: MO completed ' + Math.round(ageDays) + ' days ago (limit: ' + REVERT_WINDOW_DAYS + ' days). Adjust WASP manually.' }
      );
      return { status: 'skipped', reason: 'Revert window expired (' + Math.round(ageDays) + ' days old, limit ' + REVERT_WINDOW_DAYS + ')', moRef: moRef };
    }
  }

  var ingredients = snapshot.ingredients || [];
  var output = snapshot.output || null;
  var reversalResults = [];
  var revUomCache = {};
  var revSuccess = 0;
  var revFail = 0;

  // Step 1: Add ingredients back to PRODUCTION
  for (var i = 0; i < ingredients.length; i++) {
    var ing = ingredients[i];
    if (!ing.sku || !ing.qty) continue;

    markSyncedToWasp(ing.sku, FLOWS.MO_INGREDIENT_LOCATION, 'add');

    var addResult;
    if (ing.lot) {
      addResult = waspAddInventoryWithLot(
        ing.sku, ing.qty, FLOWS.MO_INGREDIENT_LOCATION, ing.lot, ing.expiry || '',
        '[F4-REVERT] ' + moRef + ' reverted: ' + trigger
      );
    } else {
      addResult = waspAddInventory(
        ing.sku, ing.qty, FLOWS.MO_INGREDIENT_LOCATION,
        '[F4-REVERT] ' + moRef + ' reverted: ' + trigger
      );
    }

    if (addResult && addResult.success) revSuccess++; else revFail++;
    reversalResults.push({
      sku: ing.sku,
      qty: ing.qty,
      uom: ing.uom || resolveSkuUom(ing.sku, revUomCache),
      lot: ing.lot || '',
      expiry: ing.expiry || '',
      action: 'add_ingredient',
      result: addResult
    });
  }

  // Step 2: Remove output from its stored location
  if (output && output.sku && output.qty) {
    markSyncedToWasp(output.sku, output.location, 'remove');

    var removeResult;
    if (output.lot) {
      removeResult = waspRemoveInventoryWithLot(
        output.sku, output.qty, output.location, output.lot,
        '[F4-REVERT] ' + moRef + ' reverted: ' + trigger,
        null, output.expiry || ''
      );
    } else {
      removeResult = waspRemoveInventory(
        output.sku, output.qty, output.location,
        '[F4-REVERT] ' + moRef + ' reverted: ' + trigger
      );
    }

    if (removeResult && removeResult.success) revSuccess++; else revFail++;
    reversalResults.push({
      sku: output.sku,
      qty: output.qty,
      uom: output.uom || resolveSkuUom(output.sku, revUomCache),
      lot: output.lot || '',
      expiry: output.expiry || '',
      action: 'remove_output',
      result: removeResult
    });
  }

  // Clear snapshot and consumed map — MO is no longer in Done state
  var scriptPropsRev = PropertiesService.getScriptProperties();
  scriptPropsRev.deleteProperty('mo_snapshot_' + moRef);
  scriptPropsRev.deleteProperty('mo_id_ref_' + (moId || snapshot.moId || ''));
  saveMOConsumedMap(moRef, {}, moId || snapshot.moId || '');
  try {
    var moCache = CacheService.getScriptCache();
    var cacheMoId = moId || snapshot.moId || '';
    if (cacheMoId) {
      moCache.remove('mo_done_' + cacheMoId);
      moCache.remove('mo_staging_' + cacheMoId);
      moCache.remove('mo_done_guard_' + cacheMoId);
      moCache.remove('mo_revert_guard_' + cacheMoId);
    }
  } catch (cacheErr) {
    Logger.log('reverseMOSnapshot cache cleanup error: ' + cacheErr.message);
  }

  // Build Activity log sub-items
  var revSubItems = [];
  for (var ri = 0; ri < reversalResults.length; ri++) {
    var rr = reversalResults[ri];
    var rrOk = rr.result && rr.result.success;
    var rrIsAdd = (rr.action === 'add_ingredient');
    var rrLoc = rrIsAdd ? FLOWS.MO_INGREDIENT_LOCATION : (output ? output.location : FLOWS.MO_OUTPUT_LOCATION);
    var rrCanonicalAction = buildActivityCompactMeta_(CONFIG.WASP_SITE, rrLoc, rr.lot, rr.expiry);
    var rrAction = rrIsAdd ? ('→ ' + rrLoc) : ('← ' + rrLoc);
    if (rr.lot) rrAction += '  lot:' + rr.lot;
    revSubItems.push({
      sku: rr.sku,
      qty: rr.qty,
      uom: normalizeUom(rr.uom || ''),
      success: rrOk,
      status: rrOk ? (rrIsAdd ? 'Restored' : 'Removed') : 'Failed',
      error: rrOk ? '' : (rr.result ? (rr.result.error || parseWaspError(rr.result.response, rrIsAdd ? 'Add' : 'Remove', rr.sku)) : ''),
      action: rrCanonicalAction,
      qtyColor: rrIsAdd ? 'green' : 'red'
    });
  }

  var revStatus = revFail === 0 ? 'reverted' : revSuccess === 0 ? 'failed' : 'partial';
  // Clean trigger: "status_change:NOT_STARTED" → "Not Started", "deleted" → "deleted"
  moRef = extractCanonicalActivityRef_(moRef, 'MO-', moId || snapshot.moId || '');
  var revDetail = joinActivitySegments_([
    moRef,
    ingredients.length > 0
      ? buildActivityCountSummary_(ingredients.length, 'ingredient', 'ingredients', 'restored')
      : buildActivityCountSummary_(revSubItems.length, 'line', 'lines', 'reversed')
  ]);

  try {
    logActivity('F4', revDetail, revStatus, buildActivitySourceActionContext_('Katana', 'build revert', CONFIG.WASP_SITE), revSubItems, {
      text: moRef,
      url: getKatanaWebUrl('mo', moId || snapshot.moId || '')
    });
  } catch (logErr) {
    Logger.log('reverseMOSnapshot logActivity error (non-fatal): ' + logErr.message);
  }

  try {
    sendSlackNotification(
      'MO Reverted: ' + moRef + '\n' +
      'Trigger: ' + trigger + '\n' +
      'Ingredients restored: ' + ingredients.length + ', Output removed: ' + (output ? 1 : 0)
    );
  } catch (slackErr) {
    Logger.log('reverseMOSnapshot sendSlackNotification error (non-fatal): ' + slackErr.message);
  }

  return {
    status: revStatus,
    moId: moId,
    moRef: moRef,
    trigger: trigger,
    ingredientsRestored: ingredients.length,
    outputRemoved: (output && output.sku) ? 1 : 0,
    results: reversalResults
  };
}

// ============================================
// F2 REVERT: stock_adjustment.deleted
// ============================================

/**
 * Handle stock_adjustment.deleted webhook.
 * When a Katana SA created by the F2 WASP→Katana flow is deleted,
 * reverse the corresponding WASP quantity change using the saved snapshot.
 *
 * Setup: subscribe to stock_adjustment.deleted in Katana webhook settings.
 */
function handleSADeleted(payload) {
  var saId = payload.object ? payload.object.id : (payload.id || null);
  if (!saId) {
    return { status: 'error', message: 'No SA ID in stock_adjustment.deleted webhook' };
  }

  var snapshotStr = PropertiesService.getScriptProperties().getProperty('sa_snapshot_' + saId);
  if (!snapshotStr) {
    return { status: 'ignored', reason: 'No F2 snapshot for SA ' + saId + ' — not created by this system or already cleared' };
  }

  var snapshot;
  try {
    snapshot = JSON.parse(snapshotStr);
  } catch (e) {
    return { status: 'error', message: 'Invalid snapshot JSON for SA ' + saId };
  }

  var sku = snapshot.sku;
  var qty = snapshot.qty;
  var location = snapshot.location;
  var site = snapshot.site;
  var lot = snapshot.lot || '';
  var expiry = normalizeBusinessDate_(snapshot.expiry || '');
  var originalAction = snapshot.action; // 'add' or 'remove'
  var reverseAction = originalAction === 'add' ? 'remove' : 'add';
  var uom = snapshot.uom || resolveSkuUom(sku);

  markSyncedToWasp(sku, location, reverseAction, site);

  var reverseResult;
  if (reverseAction === 'add') {
    if (lot) {
      reverseResult = waspAddInventoryWithLot(sku, qty, location, lot, expiry, '[SA-REVERT] SA ' + saId + ' deleted', site);
    } else {
      reverseResult = waspAddInventory(sku, qty, location, '[SA-REVERT] SA ' + saId + ' deleted', site);
    }
  } else {
    if (lot) {
      reverseResult = waspRemoveInventoryWithLot(sku, qty, location, lot, '[SA-REVERT] SA ' + saId + ' deleted', site, expiry);
    } else {
      reverseResult = waspRemoveInventory(sku, qty, location, '[SA-REVERT] SA ' + saId + ' deleted', site);
    }
  }

  // Clear the snapshot
  PropertiesService.getScriptProperties().deleteProperty('sa_snapshot_' + saId);

  var revOk = reverseResult && reverseResult.success;
  var revStatus = revOk ? 'reverted' : 'failed';
  var saRef = 'SA-' + saId;
  var revDetail = joinActivitySegments_([saRef, buildActivityCountSummary_(1, 'adjustment', 'adjustments', 'reversed')]);

  var subItemAction = buildActivityCompactMeta_(site, location, lot, expiry);

  logActivity('F2', revDetail, revStatus, buildActivitySourceActionContext_('Katana', 'adjust revert', site), [{
    sku: sku,
    qty: qty,
    uom: uom || '',
    success: revOk,
    status: revOk ? (reverseAction === 'add' ? 'Restored' : 'Removed') : 'Failed',
    error: revOk ? '' : (reverseResult && reverseResult.response ? parseWaspError(reverseResult.response, reverseAction === 'add' ? 'Add' : 'Remove', sku) : ''),
    action: subItemAction,
    qtyColor: reverseAction === 'add' ? 'green' : 'red'
  }], { text: saRef, url: getKatanaWebUrl('sa', saId) });

  return {
    status: revStatus,
    saId: saId,
    sku: sku,
    qty: qty,
    reversed: reverseAction
  };
}

// ============================================
// F5: SALES ORDER CANCELLED (Shopify Cancellation)
// ============================================

/**
 * Handle Sales Order Cancelled event
 * Adds items back to WASP SHIPPING-DOCK
 */
function handleSalesOrderCancelled(payload) {
  var soId = payload.object ? payload.object.id : null;

  if (!soId) {
    logToSheet('SO_CANCEL_ERROR', payload, { error: 'No SO ID found' });
    return { status: 'error', message: 'No SO ID in webhook' };
  }

  // Fetch SO details — deleted SOs return null (Katana API 404)
  var soData = fetchKatanaSalesOrder(soId);
  if (!soData) {
    // Deleted SOs can't be fetched — skip gracefully (ShipStation handles shipping deductions)
    return { status: 'skipped', reason: 'SO not fetchable (likely deleted)', soId: soId };
  }

  // Unwrap Katana API data wrapper
  var so = soData.data ? soData.data : soData;
  var items = so.sales_order_rows || so.rows || [];
  var results = [];
  var orderNumber = so.order_no || ('SO-' + soId);

  // If no embedded rows, fetch separately
  if (items.length === 0) {
    var rowsData = fetchKatanaSalesOrderRows(soId);
    items = rowsData && rowsData.data ? rowsData.data : (rowsData || []);
  }

  for (var i = 0; i < items.length; i++) {
    var item = items[i];
    var variantId = item.variant_id;
    var quantity = item.quantity || 0;

    var variant = fetchKatanaVariant(variantId);
    var sku = variant ? variant.sku : null;

    if (sku && quantity > 0 && !isSkippedSku(sku)) {
      // Mark sync BEFORE WASP call — prevents F2 echo when WASP callout fires
      markSyncedToWasp(sku, FLOWS.PICK_FROM_LOCATION, 'add');

      var result = waspAddInventory(
        sku,
        quantity,
        FLOWS.PICK_FROM_LOCATION,
        'SO Cancelled: ' + orderNumber
      );
      results.push({ sku: sku, quantity: quantity, result: result, uom: resolveVariantUom(variant) });
    }
  }

  // Activity Log — F5 Cancellation
  var f5cSuccess = 0;
  var f5cFail = 0;
  var f5cSubItems = [];
  for (var r = 0; r < results.length; r++) {
    var res = results[r];
    var ok = res.result && res.result.success;
    if (ok) f5cSuccess++; else f5cFail++;
    f5cSubItems.push({
      sku: res.sku,
      qty: res.quantity,
      uom: res.uom || '',
      success: ok,
      status: ok ? 'Returned' : '',
      error: ok ? '' : (res.result ? (res.result.error || parseWaspError(res.result.response, 'Add', res.sku)) : ''),
      action: ok ? buildActivityCompactMeta_(CONFIG.WASP_SITE, FLOWS.PICK_FROM_LOCATION, res.lot || '', res.expiry || '') : '',
      qtyColor: 'green'
    });
  }
  var f5cStatus = f5cFail === 0 ? 'success' : f5cSuccess === 0 ? 'failed' : 'partial';
  var f5cRef = extractCanonicalActivityRef_(orderNumber, 'SO-', soId);
  var f5cDetail = joinActivitySegments_([
    f5cRef,
    buildActivityCountSummary_(results.length, 'line', 'lines', 'returned')
  ]);
  var f5cExecId = logActivity('F5', f5cDetail, f5cStatus, buildActivitySourceActionContext_('Katana', 'cancel return', CONFIG.WASP_SITE), f5cSubItems.length > 1 ? f5cSubItems : null, {
    text: f5cRef,
    url: getKatanaWebUrl('so', soId)
  });

  var f5cFlowItems = [];
  for (var fl = 0; fl < results.length; fl++) {
    var fRes = results[fl];
    var fOk = fRes.result && fRes.result.success;
    f5cFlowItems.push({
      sku: fRes.sku,
      qty: fRes.quantity,
      detail: '→ ' + FLOWS.PICK_FROM_LOCATION,
      status: fOk ? 'Returned' : 'Failed',
      error: fOk ? '' : (fRes.result ? (fRes.result.error || parseWaspError(fRes.result.response, 'Add', fRes.sku)) : ''),
      qtyColor: 'green'
    });
  }
  logFlowDetail('F5', f5cExecId, {
    ref: orderNumber,
    detail: 'CANCELLED  ' + results.length + ' item' + (results.length !== 1 ? 's' : '') + ' returned → ' + FLOWS.PICK_FROM_LOCATION,
    status: f5cStatus === 'success' ? 'Returned' : f5cStatus === 'failed' ? 'Failed' : 'Partial',
    linkText: orderNumber,
    linkUrl: getKatanaWebUrl('so', soId)
  }, f5cFlowItems);

  // Clean up pick order mapping if it exists
  var pickMapping = getPickOrderBySoId(soId);
  if (pickMapping) {
    clearPickOrderMapping(pickMapping.orderNumber);
  }

  sendSlackNotification(
    'SO Cancelled: ' + orderNumber +
    '\nItems returned: ' + results.length +
    '\nLocation: ' + FLOWS.PICK_FROM_LOCATION
  );

  return {
    status: 'processed',
    soId: soId,
    orderNumber: orderNumber,
    itemsReturned: results.length,
    results: results
  };
}

// ============================================
// SALES ORDER UPDATED (Fallback cancel detection)
// ============================================
// Katana may not fire sales_order.cancelled directly.
// This checks sales_order.updated for cancellation status.
// ============================================

function handleSalesOrderUpdated(payload) {
  var soId = payload.object ? payload.object.id : null;
  if (!soId) return { status: 'ignored', reason: 'No SO ID' };

  // Check if payload includes status indicating cancellation
  // toUpperCase() required — 01_Main.gs filters queue using uppercase comparison,
  // so Katana sends uppercase ('CANCELLED'/'VOIDED'). Lowercase check never matched.
  var objStatus = payload.object ? (payload.object.status || '').toUpperCase() : '';

  if (objStatus === 'CANCELLED' || objStatus === 'VOIDED') {
    if (getHotfixFlag_('F5_CONFIRM_CANCEL_VIA_UPDATE')) {
      var confirmedCancelledSO = refetchKatanaEntityAfterConfirmDelay_(fetchKatanaSalesOrder, soId);
      if (!confirmedCancelledSO) {
        return { status: 'skipped', reason: 'F5 cancel-via-update confirmation fetch failed for SO ' + soId };
      }
      var confirmedCancelStatus = (confirmedCancelledSO.status || '').toUpperCase();
      if (confirmedCancelStatus !== 'CANCELLED' && confirmedCancelStatus !== 'VOIDED') {
        return { status: 'ignored', reason: 'F5 cancel-via-update not confirmed; SO status is ' + confirmedCancelStatus };
      }
    }
    logToSheet('SO_CANCEL_VIA_UPDATE', { soId: soId, status: objStatus }, {});
    return handleSalesOrderCancelled(payload);
  }

  // Katana sends sales_order.updated (not sales_order.delivered) for Amazon US SOs.
  // Delegate only when the live SO or payload still resolves to the exact Amazon US customer.
  // Shopify SOs: handled by ShipStation — ignored inside handleSalesOrderDelivered via source check.
  // All other SOs: use sales_order.delivered directly, not .updated.
  if (objStatus === 'DELIVERED' || objStatus === 'PARTIALLY_DELIVERED') {
    var liveSoData = fetchKatanaSalesOrder(soId);
    var liveSo = liveSoData && liveSoData.data ? liveSoData.data : liveSoData;
    liveSo = enrichSalesOrderWithAddresses_(soId, liveSo);
    var isAmazonUSUpdate = isAmazonUSSalesOrder_(liveSo, payload.object || {});
    if (isAmazonUSUpdate) {
      return handleSalesOrderDelivered(payload);
    }
    var soF6Diag = buildAmazonUSSalesOrderMatchDebug_(liveSo, payload.object || {});
    logToSheet('SO_F6_MATCH_DIAG', soF6Diag, {});
    logWebhookQueue(
      { resource_type: 'sales_order', action: 'so_f6_match.diag', object: soF6Diag },
      { status: 'diag', reason: 'Amazon US matcher debug' }
    );
    return { status: 'ignored', reason: objStatus + ' via .updated — not an Amazon US SO' };
  }

  // Check for F6 revert — SO was F6-processed and delivery was reverted
  var f6Stored = PropertiesService.getScriptProperties().getProperty('f6_delivered_' + soId);
  if (f6Stored) {
    // Webhook says NOT_SHIPPED — a delivery was reverted in Katana.
    // Compare current fulfillments against stored items to find what was un-delivered.
    if (objStatus === 'NOT_SHIPPED') {
      try {
        var revertFulfillData = katanaApiCall('sales_order_fulfillments?sales_order_id=' + soId);
        var revertFulfillments = revertFulfillData && revertFulfillData.data ? revertFulfillData.data : (revertFulfillData || []);
        // Build set of currently fulfilled SO row IDs
        var currentlyFulfilledRows = {};
        if (Array.isArray(revertFulfillments)) {
          for (var rfi = 0; rfi < revertFulfillments.length; rfi++) {
            var rfRows = revertFulfillments[rfi].sales_order_fulfillment_rows || revertFulfillments[rfi].rows || [];
            for (var rfr = 0; rfr < rfRows.length; rfr++) {
              var rfRowId = String(rfRows[rfr].sales_order_row_id || rfRows[rfr].id || '');
              if (rfRowId) currentlyFulfilledRows[rfRowId] = true;
            }
          }
        }
        // Parse stored delivery data and find items no longer fulfilled
        var storedData = JSON.parse(f6Stored);
        var storedItems = storedData.items || [];
        var revertedItems = [];
        var remainingStored = [];
        // Get SO rows to map variant_id → row id
        var revertSORows = katanaApiCall('sales_order_rows?sales_order_id=' + soId);
        var revertRows = revertSORows && revertSORows.data ? revertSORows.data : (revertSORows || []);
        var variantToRowId = {};
        for (var rvr = 0; rvr < revertRows.length; rvr++) {
          var rvVariant = fetchKatanaVariant(revertRows[rvr].variant_id);
          if (rvVariant && rvVariant.sku) {
            variantToRowId[rvVariant.sku] = String(revertRows[rvr].id || '');
          }
        }
        for (var rsi = 0; rsi < storedItems.length; rsi++) {
          var rItem = storedItems[rsi];
          var rRowId = variantToRowId[rItem.sku] || '';
          if (rRowId && currentlyFulfilledRows[rRowId]) {
            remainingStored.push(rItem);
          } else {
            revertedItems.push(rItem);
          }
        }
        if (revertedItems.length > 0) {
          // Build partial revert JSON with only the reverted items
          var partialRevertData = {
            orderNo: storedData.orderNo,
            items: revertedItems,
            rowProcessedQty: storedData.rowProcessedQty || {}
          };
          logWebhookQueue(
            { action: 'f6_partial_revert', soId: soId },
            { status: 'diag', message: 'Reverting ' + revertedItems.length + ' items, ' + remainingStored.length + ' remain fulfilled' }
          );
          // Update stored data to only keep remaining items
          if (remainingStored.length > 0) {
            storedData.items = remainingStored;
            PropertiesService.getScriptProperties().setProperty('f6_delivered_' + soId, JSON.stringify(storedData));
          }
          return handleSalesOrderF6Revert(soId, JSON.stringify(partialRevertData));
        }
        // No items reverted — all still fulfilled
        return { status: 'ignored', reason: 'F6 SO: all stored items still fulfilled' };
      } catch (revertErr) {
        logWebhookQueue(
          { action: 'f6_revert_error', soId: soId },
          { status: 'error', message: 'F6 revert fulfillment check error: ' + revertErr.message }
        );
      }
    }
    // Full revert: SO went to a non-delivered state entirely
    var soData = fetchKatanaSalesOrder(soId);
    if (soData) {
      var so = soData.data ? soData.data : soData;
      var currentStatus = (so.status || '').toUpperCase();
      if (currentStatus !== 'DELIVERED' && currentStatus !== 'PARTIALLY_DELIVERED') {
        return handleSalesOrderF6Revert(soId, f6Stored);
      }
    }
    return { status: 'ignored', reason: 'F6 SO still delivered' };
  }

  // Not a cancellation or F6 revert — ignore silently
  return { status: 'ignored', event: 'sales_order.updated' };
}

// ============================================
// F6 REVERT
// ============================================

/**
 * Reverse an F6 delivery — SO was un-delivered (Revert clicked in Katana).
 * Removes from AMAZON-FBA-USA, adds back to PRODUCTION.
 * Uses items stored in ScriptProperties at delivery time.
 */
function handleSalesOrderF6Revert(soId, f6StoredJson) {
  var f6Data;
  try {
    f6Data = JSON.parse(f6StoredJson);
  } catch (parseErr) {
    Logger.log('F6 revert: invalid stored JSON for SO ' + soId);
    return { status: 'error', message: 'F6 revert: invalid stored data' };
  }

  var items = f6Data.items || [];
  var orderNo = f6Data.orderNo || ('SO-' + soId);

  if (items.length === 0) {
    return { status: 'skipped', reason: 'F6 revert: no items stored' };
  }

  // Dedup — catches rapid duplicate sales_order.updated events
  var cache = CacheService.getScriptCache();
  var dedupKey = 'f6_revert_' + soId;
  var useExactF6RevertState = !!(typeof HOTFIX_FLAGS !== 'undefined' && HOTFIX_FLAGS.F6_REVERT_PRESERVE_EXACT_LOT);
  if (cache.get(dedupKey)) {
    return { status: 'skipped', reason: 'F6 revert already in progress' };
  }
  cache.put(dedupKey, 'true', 300);

  if (!useExactF6RevertState) {
    // Legacy behavior: delete immediately to avoid duplicate reverts.
    try {
      PropertiesService.getScriptProperties().deleteProperty('f6_delivered_' + soId);
    } catch (delErr) {
      Logger.log('F6 revert: delete property error: ' + delErr.message);
    }
  }

  var results = [];
  var remainingItems = [];
  var f6RevertUomCache = {};
  for (var i = 0; i < items.length; i++) {
    var item = items[i];
    var sku = item.sku;
    var qty = item.qty;
    var storedLot = String(item.lot || '').trim();
    var storedExpiry = normalizeBusinessDate_(item.expiry || '');

    if (sku && qty > 0) {
      // Step 1: Remove from AMAZON-FBA-USA (lot-aware)
      markSyncedToWasp(sku, FLOWS.AMAZON_FBA_WASP_LOCATION, 'remove', FLOWS.AMAZON_FBA_WASP_SITE);
      var rLotInfo = null;
      if (useExactF6RevertState && storedLot) {
        rLotInfo = { lot: storedLot, dateCode: storedExpiry };
      } else {
        rLotInfo = waspLookupItemLotAndDate(sku, FLOWS.AMAZON_FBA_WASP_LOCATION, FLOWS.AMAZON_FBA_WASP_SITE);
      }
      var removeResult;
      if (rLotInfo && rLotInfo.lot) {
        removeResult = waspRemoveInventoryWithLot(
          sku, qty, FLOWS.AMAZON_FBA_WASP_LOCATION, rLotInfo.lot,
          'F6 Revert: ' + orderNo, FLOWS.AMAZON_FBA_WASP_SITE, rLotInfo.dateCode
        );
      } else {
        removeResult = waspRemoveInventory(
          sku, qty, FLOWS.AMAZON_FBA_WASP_LOCATION,
          'F6 Revert: ' + orderNo,
          FLOWS.AMAZON_FBA_WASP_SITE
        );
      }

      // Step 2: Add back to PRODUCTION (same lot if applicable)
      var addResult = null;
      if (removeResult && removeResult.success) {
        markSyncedToWasp(sku, FLOWS.AMAZON_TRANSFER_LOCATION, 'add', CONFIG.WASP_SITE);
        if (rLotInfo && rLotInfo.lot) {
          addResult = waspAddInventoryWithLot(
            sku, qty, FLOWS.AMAZON_TRANSFER_LOCATION, rLotInfo.lot, rLotInfo.dateCode,
            'F6 Revert: ' + orderNo
          );
        } else {
          addResult = waspAddInventory(
            sku, qty, FLOWS.AMAZON_TRANSFER_LOCATION,
            'F6 Revert: ' + orderNo
          );
        }
      }

      var ok = !!(removeResult && removeResult.success && addResult && addResult.success);
      results.push({
        sku: sku,
        qty: qty,
        uom: item.uom || resolveSkuUom(sku, f6RevertUomCache),
        lot: rLotInfo && rLotInfo.lot ? rLotInfo.lot : storedLot,
        expiry: normalizeBusinessDate_((rLotInfo && rLotInfo.dateCode) ? rLotInfo.dateCode : storedExpiry),
        ok: ok,
        removeResult: removeResult,
        addResult: addResult
      });
      if (useExactF6RevertState && !ok) {
        remainingItems.push({
          sku: sku,
          qty: qty,
          uom: item.uom || resolveSkuUom(sku, f6RevertUomCache),
          lot: storedLot,
          expiry: storedExpiry
        });
      }
    }
  }

  if (useExactF6RevertState) {
    try {
      if (remainingItems.length > 0) {
        PropertiesService.getScriptProperties().setProperty(
          'f6_delivered_' + soId,
          JSON.stringify({ orderNo: orderNo, items: remainingItems })
        );
      } else {
        PropertiesService.getScriptProperties().deleteProperty('f6_delivered_' + soId);
      }
    } catch (f6PersistErr) {
      Logger.log('F6 revert: persist retry state error: ' + f6PersistErr.message);
    }
  }

  // Activity Log
  var rSuccess = 0;
  var rFail = 0;
  var rSubItems = [];
  var rLocationLabel = FLOWS.AMAZON_FBA_WASP_LOCATION + ' \u2192 ' + FLOWS.AMAZON_TRANSFER_LOCATION;
  for (var r = 0; r < results.length; r++) {
    var res = results[r];
    if (res.ok) rSuccess++; else rFail++;
    var rError = '';
    if (!res.ok) {
      var rrOk = res.removeResult && res.removeResult.success;
      rError = rrOk
        ? (res.addResult ? (res.addResult.error || parseWaspError(res.addResult.response, 'Add', res.sku)) : 'Add skipped')
        : (res.removeResult ? (res.removeResult.error || parseWaspError(res.removeResult.response, 'Remove', res.sku)) : '');
    }
    rSubItems.push({
      sku: res.sku,
      qty: res.qty,
      uom: res.uom || '',
      success: res.ok,
      status: res.ok ? 'Reverted' : 'Failed',
      error: rError,
      action: buildActivityCompactMeta_(CONFIG.WASP_SITE, FLOWS.AMAZON_TRANSFER_LOCATION, res.lot, res.expiry),
      qtyColor: 'green'
    });
  }

  var rStatus = rFail === 0 ? 'reverted' : rSuccess === 0 ? 'failed' : 'partial';
  var rRef = extractCanonicalActivityRef_(orderNo, 'SO-', soId);
  var rDetail = joinActivitySegments_([
    rRef,
    buildActivityCountSummary_(results.length, 'line', 'lines', 'reversed')
  ]);

  logActivity('F6', rDetail, rStatus, buildActivityTransferContext_('Katana', 'fba revert', FLOWS.AMAZON_FBA_WASP_SITE, CONFIG.WASP_SITE), rSubItems, {
    text: rRef,
    url: getKatanaWebUrl('so', soId)
  });

  return {
    status: rStatus,
    soId: soId,
    flow: 'F6',
    results: results
  };
}
