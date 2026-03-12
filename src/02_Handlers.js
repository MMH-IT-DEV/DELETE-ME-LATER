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
  sendSlackNotification('PO Created in Katana\nID: ' + poId);
  return { status: 'logged', poId: poId };
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

  // Fetch PO header
  var poData = fetchKatanaPO(poId);
  if (!poData) {
    return { status: 'error', message: 'Failed to fetch PO header' };
  }

  var poNumber = poData.order_no || ('PO-' + poId);
  var isPartialReceive = (poData.status === 'PARTIALLY_RECEIVED');

  // Load any previously stored data for this PO (from earlier partial receives)
  var f1SaveKey = 'f1_recv_' + poNumber.replace(/[^a-zA-Z0-9_]/g, '');
  var storedJson = PropertiesService.getScriptProperties().getProperty(f1SaveKey);
  var previouslySaved = (storedJson && storedJson.length > 2) ? JSON.parse(storedJson) : [];
  var previousBatchIds = {};
  for (var psi = 0; psi < previouslySaved.length; psi++) {
    if (previouslySaved[psi].batchId) previousBatchIds[String(previouslySaved[psi].batchId)] = true;
  }

  // Dedup: partial receives use a short 30s window; full receives use 600s + Activity log check
  var cache = CacheService.getScriptCache();
  var poDedupKey = 'po_received_' + poNumber;
  if (isPartialReceive) {
    var partialDedupKey = poDedupKey + '_partial';
    if (cache.get(partialDedupKey)) {
      return { status: 'skipped', reason: 'PO already processed (cache)', poNumber: poNumber };
    }
    cache.put(partialDedupKey, 'true', 30);
  } else {
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
  var f1CountKey = 'f1_count_' + poNumber.replace(/[^a-zA-Z0-9_]/g, '');
  var receiveCount = parseInt(PropertiesService.getScriptProperties().getProperty(f1CountKey) || '0', 10) + 1;
  PropertiesService.getScriptProperties().setProperty(f1CountKey, String(receiveCount));
  var poRefLabel = receiveCount > 1 ? (poNumber + '/' + receiveCount) : poNumber;

  // Resolve WASP destination from PO "Ship to" location
  var poLocationId = poData.location_id || null;
  var waspDest = { site: CONFIG.WASP_SITE, location: FLOWS.PO_RECEIVING_LOCATION };

  if (poLocationId) {
    var poLocation = fetchKatanaLocation(poLocationId);
    var locName = poLocation ? (poLocation.name || '') : '';
    if (locName && KATANA_LOCATION_TO_WASP[locName]) {
      waspDest = KATANA_LOCATION_TO_WASP[locName];
    }
  }

  // Fetch PO rows
  var poRowsData = fetchKatanaPORows(poId);
  if (!poRowsData) {
    return { status: 'error', message: 'Failed to fetch PO rows' };
  }

  var rows = poRowsData.data || poRowsData || [];
  var results = [];
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
    var variant = fetchKatanaVariant(variantId);
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
      if (isPartialReceive && bt.length === 0) {
        if (typeof HOTFIX_FLAGS !== 'undefined' && HOTFIX_FLAGS.F1_PARTIAL_NON_BATCH_DELTA && poRowId) {
          var cumulativeReceivedQty = parseFloat(row.received_quantity || row.receivedQuantity || row.received_qty || 0) || 0;
          var cumulativeReceivedStockQty = cumulativeReceivedQty * puomRate;
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

  // Save received batch data for revert detection (purchase_order.updated handler)
  // Append to any previously saved data from earlier partial receives
  var f1SaveData = previouslySaved.slice(); // start with what was already stored
  for (var rs = 0; rs < results.length; rs++) {
    var rsItem = results[rs];
    if (rsItem.result && rsItem.result.success) {
      f1SaveData.push({
        sku: rsItem.sku,
        variantId: rsItem.variantId || null,
        poRowId: rsItem.poRowId || null,
        batchId: rsItem.batchId || null,
        lot: rsItem.lot || '',
        qty: rsItem.quantity,
        uom: rsItem.uom || '',
        expiry: normalizeBusinessDate_(rsItem.expiry || ''),
        location: rsItem.location,
        site: rsItem.site
      });
    }
  }
  if (f1SaveData.length > 0) {
    PropertiesService.getScriptProperties().setProperty(f1SaveKey, JSON.stringify(f1SaveData));
  }

  // Track partial receive state so isPOAlreadyReceived allows a follow-up receive
  var f1PartialKey = 'f1_recv_partial_' + poNumber.replace(/[^a-zA-Z0-9_]/g, '');
  if (isPartialReceive) {
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
    }
  }
  var f1Status = f1Fail === 0 ? 'success' : f1Success === 0 ? 'failed' : 'partial';
  var poRef = poData.order_no || ('PO-' + poId);

  // Build location summary for logging (handles mixed destinations)
  var f1LocMap = {};
  for (var lr = 0; lr < results.length; lr++) {
    var rl = results[lr].location || waspDest.location;
    f1LocMap[rl] = true;
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
  var f1ExecId = logActivity('F1', f1Detail, f1Status, '→ ' + f1LocSummary, f1SubItems.length > 0 ? f1SubItems : null, {
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

  return {
    status: 'processed',
    poId: poId,
    itemsProcessed: results.length,
    locations: f1LocKeys,
    site: waspDest.site,
    results: results
  };
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

  var poData = fetchKatanaPO(poId);
  if (!poData) return { status: 'error', message: 'Failed to fetch PO' };

  var poRef = poData.order_no || ('PO-' + poId);
  var f1SaveKey = 'f1_recv_' + poRef.replace(/[^a-zA-Z0-9_]/g, '');
  var storedJson = PropertiesService.getScriptProperties().getProperty(f1SaveKey);

  if (!storedJson) {
    return { status: 'ignored', reason: 'No F1 receive data stored for ' + poRef };
  }

  var storedBatches;
  try { storedBatches = JSON.parse(storedJson); } catch (e) {
    return { status: 'error', message: 'Corrupt stored data for ' + poRef };
  }

  // Primary revert signal: PO status flipped to NOT_RECEIVED → all stored items reverted.
  // This handles non-lot-tracked items (batchId=null) which the batch comparison can't detect.
  var reverted = [];
  var remaining = [];

  // Always fetch rows so we can log diagnostics for partial-revert analysis
  var poRowsDataAll = fetchKatanaPORows(poId);
  var rowsAll = poRowsDataAll ? (poRowsDataAll.data || poRowsDataAll || []) : [];
  var refreshedAfterStatusConfirm = false;

  if (poData.status === 'NOT_RECEIVED' && getHotfixFlag_('F1_CONFIRM_FULL_REVERT')) {
    var confirmedPO = refetchKatanaEntityAfterConfirmDelay_(fetchKatanaPO, poId);
    if (confirmedPO) {
      poData = confirmedPO;
      var confirmRowsDataAll = fetchKatanaPORows(poId);
      if (confirmRowsDataAll) {
        poRowsDataAll = confirmRowsDataAll;
        rowsAll = confirmRowsDataAll.data || confirmRowsDataAll || [];
      }
      refreshedAfterStatusConfirm = true;
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

  if (poData.status === 'NOT_RECEIVED') {
    // Full un-receive: every stored item was removed from Katana
    reverted = storedBatches.slice();
  } else {
    // PO still received — fall back to batch-ID comparison for lot-tracked partial reverts
    var poRowsData = poRowsDataAll;
    var rows = rowsAll;
    var currentBatchIds = {};
    var currentNonBatchReceivedQtyByRow = buildCurrentPOReceivedQtyByRow_(rows);
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
        currentNonBatchReceivedQtyByRow = buildCurrentPOReceivedQtyByRow_(confirmRows);
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
  }
  var rvStatus = rvFail === 0 ? 'reverted' : rvSuccess === 0 ? 'failed' : 'partial';
  var rvDetail = poRef + '  ' + (reverted.length === 1
    ? reverted[0].sku + ' x' + reverted[0].qty + ' reverted'
    : reverted.length + ' items reverted');
  if (rvFail > 0) rvDetail += '  ' + rvFail + ' error' + (rvFail > 1 ? 's' : '');

  logActivity('F1', rvDetail, rvStatus, 'Revert → ' + waspDest.site, rvSubItems, {
    text: poRef,
    url: getKatanaWebUrl('po', poId)
  });

  // Clear dedup cache and partial/count flags so a fresh receive can go through immediately
  var rPoKey = poRef.replace(/[^a-zA-Z0-9_]/g, '');
  CacheService.getScriptCache().remove('po_received_' + poRef);
  PropertiesService.getScriptProperties().deleteProperty('f1_recv_partial_' + rPoKey);
  PropertiesService.getScriptProperties().deleteProperty('f1_count_' + rPoKey);

  return { status: 'reverted', count: reverted.length, poRef: poRef };
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

  var candidateIds = [
    so ? so.customer_id : null,
    payloadObject ? payloadObject.customer_id : null
  ];
  for (var ci = 0; ci < candidateIds.length; ci++) {
    var candidateId = candidateIds[ci];
    if (candidateId !== null && candidateId !== undefined && candidateId !== '') {
      if (idMap[String(candidateId).trim()]) return true;
    }
  }

  var candidateNames = [
    so ? so.customer_name : '',
    so && so.customer && typeof so.customer === 'object' ? (so.customer.name || '') : (so ? so.customer : ''),
    payloadObject ? payloadObject.customer_name : '',
    payloadObject && payloadObject.customer && typeof payloadObject.customer === 'object' ? (payloadObject.customer.name || '') : (payloadObject ? payloadObject.customer : '')
  ];
  for (var cn = 0; cn < candidateNames.length; cn++) {
    var candidateName = String(candidateNames[cn] || '').trim().toLowerCase();
    if (candidateName === 'amazon us') return true;
  }

  return false;
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
  var cache = CacheService.getScriptCache();
  var dedupKey = 'so_delivered_' + soId;
  if (cache.get(dedupKey)) {
    return { status: 'skipped', reason: 'SO already processed', soId: soId };
  }
  if (isF6 || isAmazonFBA) {
    cache.put(dedupKey, 'true', 300); // 5-minute dedup window (only for actionable flows)
  }
  var flowLabel = isF6 ? 'F6' : 'F3';
  var removeLocation = isAmazonFBA ? FLOWS.AMAZON_FBA_WASP_LOCATION : FLOWS.AMAZON_TRANSFER_LOCATION;
  var removeSite = isAmazonFBA ? FLOWS.AMAZON_FBA_WASP_SITE : null; // null = default site (MMH Kelowna)

  var items = so.sales_order_rows || so.rows || [];

  // If no embedded rows, fetch separately
  if (items.length === 0) {
    var rowsData = fetchKatanaSalesOrderRows(soId);
    items = rowsData && rowsData.data ? rowsData.data : (rowsData || []);
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

  for (var i = 0; i < items.length; i++) {
    var item = items[i];
    var variantId = item.variant_id;
    var quantity = (item.delivered_quantity != null) ? item.delivered_quantity : (item.quantity || 0);

    // Get SKU from variant
    var variant = fetchKatanaVariant(variantId);
    var sku = variant ? variant.sku : null;

    if (sku && quantity > 0 && !isSkippedSku(sku)) {
      if (isF6) {
        // F6: SO from MMH Kelowna → FBA shipment
        // Step 1: Remove from SHIPPING-DOCK (lot-aware)
        // Step 2: Add to AMAZON-FBA-USA with same lot if applicable
        markSyncedToWasp(sku, removeLocation, 'remove');
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
          markSyncedToWasp(sku, FLOWS.AMAZON_FBA_WASP_LOCATION, 'add');
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
          isF6: true
        });
      } else {
        // Mark sync BEFORE WASP call — prevents F2 echo when WASP callout fires
        markSyncedToWasp(sku, removeLocation, 'remove');

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
    var f6StoredItems = [];
    var useExactF6RevertState = !!(typeof HOTFIX_FLAGS !== 'undefined' && HOTFIX_FLAGS.F6_REVERT_PRESERVE_EXACT_LOT);
    for (var f6si = 0; f6si < results.length; f6si++) {
      if (results[f6si].isF6) {
        if (useExactF6RevertState) {
          if (!(results[f6si].result && results[f6si].result.success)) continue;
          f6StoredItems.push({
            sku: results[f6si].sku,
            qty: results[f6si].quantity,
            uom: results[f6si].uom || '',
            lot: results[f6si].lot || '',
            expiry: normalizeBusinessDate_(results[f6si].expiry || '')
          });
        } else {
          f6StoredItems.push({ sku: results[f6si].sku, qty: results[f6si].quantity, uom: results[f6si].uom || '' });
        }
      }
    }
    if (f6StoredItems.length > 0) {
      try {
        PropertiesService.getScriptProperties().setProperty(
          'f6_delivered_' + soId,
          JSON.stringify({ orderNo: soRef, items: f6StoredItems })
        );
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
    var actionText = dOk ? (dRes.isF6 ? (removeLocation + ' \u2192 ' + FLOWS.AMAZON_FBA_WASP_LOCATION) : removeLocation) : '';
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
    sdSubItems.push({
      sku: dRes.sku,
      qty: dRes.quantity,
      uom: dRes.uom || '',
      success: dOk,
      status: dRes.skipped ? 'Picked' : (dOk ? 'Complete' : 'Failed'),
      error: itemError,
      action: actionText,
      qtyColor: dRes.isF6 ? 'orange' : 'red'
    });
  }
  var sdStatus = sdFail === 0 ? 'success' : sdSuccess === 0 ? 'failed' : 'partial';
  var sdDetail;
  if (results.length === 1) {
    sdDetail = soRef + '  ' + results[0].sku + ' x' + results[0].quantity;
    if (sdFail > 0 && sdSubItems[0] && sdSubItems[0].error) sdDetail += '  ' + sdSubItems[0].error;
  } else {
    sdDetail = soRef + '  ' + results.length + ' items';
    if (sdFail > 0) sdDetail += '  ' + sdFail + ' error' + (sdFail > 1 ? 's' : '');
  }
  var sdLocationLabel = isF6 ? (removeLocation + ' \u2192 ' + FLOWS.AMAZON_FBA_WASP_LOCATION) : removeLocation;
  var sdExecId = logActivity(flowLabel, sdDetail, sdStatus, '\u2192 ' + sdLocationLabel, (isF6 || sdSubItems.length > 1) ? sdSubItems : null, {
    text: soRef,
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

/**
 * Check if an MO already has a completed F4 entry in Activity.
 * Partial/failed rows are intentionally reprocessable.
 */
function isMOAlreadyCompleted(moRef) {
  try {
    var activitySheet = getActivitySheet();
    var lastRow = activitySheet.getLastRow();
    if (lastRow <= 3) return false;

    // Read A:E so status can distinguish Complete vs Partial/Failed/Reverted.
    var data = activitySheet.getRange(4, 1, lastRow - 3, 5).getValues();

    for (var i = data.length - 1; i >= 0; i--) {
      if (!data[i][0]) continue;
      if (String(data[i][2]) !== 'F4 Manufacturing') continue;
      var details = String(data[i][3]);
      if (details.indexOf(moRef) < 0) continue;
      var statusText = String(data[i][4] || '').trim();
      if (details.indexOf('reverted') >= 0 || statusText === 'Reverted') return false;
      return statusText === 'Complete';
    }
    return false;
  } catch (e) {
    Logger.log('isMOAlreadyCompleted error: ' + e.message);
    return false;
  }
}

/**
 * Get map of ingredient SKUs already consumed for an MO on a previous run.
 * Used for smart retry — prevents double-consumption when MO is reprocessed.
 * Returns object { sku: count, ... } or empty object.
 */
function getMOConsumedMap(moRef) {
  try {
    var props = PropertiesService.getScriptProperties();
    var raw = props.getProperty('mo_consumed_' + moRef);
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
function saveMOConsumedMap(moRef, consumedMap) {
  try {
    var props = PropertiesService.getScriptProperties();
    var hasAny = false;
    for (var k in consumedMap) {
      if (consumedMap.hasOwnProperty(k) && consumedMap[k] > 0) {
        hasAny = true;
        break;
      }
    }
    if (!hasAny) {
      props.deleteProperty('mo_consumed_' + moRef);
    } else {
      props.setProperty('mo_consumed_' + moRef, JSON.stringify(consumedMap));
    }
  } catch (e) {
    Logger.log('saveMOConsumedMap error: ' + e.message);
  }
}

function getMOSnapshot(moRef) {
  try {
    var props = PropertiesService.getScriptProperties();
    var raw = props.getProperty('mo_snapshot_' + moRef);
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
      props.deleteProperty('mo_snapshot_' + moRef);
      props.deleteProperty('mo_id_ref_' + (moId || ''));
      return;
    }
    props.setProperty('mo_snapshot_' + moRef, JSON.stringify(snapshotObj));
    if (moId) props.setProperty('mo_id_ref_' + moId, moRef);
  } catch (e) {
    Logger.log('saveMOSnapshot error: ' + e.message);
  }
}

function buildCurrentPOReceivedQtyByRow_(rows) {
  rows = rows || [];
  var receivedQtyByRow = {};

  for (var i = 0; i < rows.length; i++) {
    var row = rows[i] || {};
    var bt = row.batch_transactions || row.batchTransactions || [];
    var rowId = row.id || row.purchase_order_row_id || row.po_row_id || null;
    if (!rowId || bt.length > 0) continue;

    var rowPuomRate = parseFloat(row.purchase_uom_conversion_rate) || 1;
    var rowReceivedQty = parseFloat(row.received_quantity || row.receivedQuantity || row.received_qty || 0) || 0;
    receivedQtyByRow[String(rowId)] = rowReceivedQty * rowPuomRate;
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
    // If a partial receive was recorded, allow re-receiving (receive all or next partial)
    var partialKey = 'f1_recv_partial_' + poRef.replace(/[^a-zA-Z0-9_]/g, '');
    if (PropertiesService.getScriptProperties().getProperty(partialKey) === 'true') return false;

    var activitySheet = getActivitySheet();
    var lastRow = activitySheet.getLastRow();
    if (lastRow <= 3) return false;

    // Read 3 columns: A(execId), C(flow), D(details)
    var data = activitySheet.getRange(4, 1, lastRow - 3, 4).getValues();

    for (var i = data.length - 1; i >= 0; i--) {
      if (!data[i][0]) continue;
      if (String(data[i][2]) !== 'F1 Receiving') continue;
      var details = String(data[i][3]);
      if (details.indexOf(poRef) < 0) continue;
      // Most-recent F1 entry for this PO found.
      // If it was a revert, allow re-processing (PO can be re-received).
      if (details.indexOf('reverted') >= 0) return false;
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
  var existingSnapshot = getMOSnapshot(moOrderNo);
  if (isMOAlreadyCompleted(moOrderNo)) {
    logToSheet('MO_DONE_DUPLICATE', { moId: moId, moRef: moOrderNo }, { reason: 'Already completed in Activity log' });
    return { status: 'skipped', reason: 'MO already completed', moId: moId, moRef: moOrderNo };
  }

  // Skip auto-assembly (ASM) MOs — bundle products (VB-) assemble finished goods,
  // not raw ingredients at PRODUCTION. F5 (ShipStation) handles all VB- deductions
  // at shipment time, so F4 processing these creates noise with no inventory effect.
  if (moOrderNo.indexOf(' ASM ') >= 0) {
    return { status: 'ignored', reason: 'Auto-assembly MO skipped: ' + moOrderNo };
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

  // Wait for Katana to propagate batch assignments to recipe rows.
  // User assigns batch in the "Done" popup — Katana's backend needs a few seconds
  // before batch_transactions appear on recipe_rows via API.
  Utilities.sleep(15000);

  // Refetch the MO header after the same delay so output_batch / batch_number
  // chosen in the Done popup has time to propagate too.
  var confirmedMOAfterDone = fetchKatanaMO(moId);
  var confirmedMO = confirmedMOAfterDone && confirmedMOAfterDone.data ? confirmedMOAfterDone.data : confirmedMOAfterDone;
  if (confirmedMO) {
    mo = confirmedMO;
    moOrderNo = mo.order_no || moOrderNo;
    outputQuantity = mo.actual_quantity || mo.quantity || outputQuantity;

    outputBatchData = resolveMOOutputBatchData_(mo, variantId, moOrderNo);
    lotNumber = outputBatchData.lot || lotNumber || '';
    expiryDate = outputBatchData.expiry || expiryDate || '';
  }

  // Fetch ingredients (with batch_transactions included)
  var ingredientsData = fetchKatanaMOIngredients(moId);

  var ingredients = ingredientsData && ingredientsData.data ? ingredientsData.data : (ingredientsData || []);

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
  var consumedMap = getMOConsumedMap(moOrderNo);
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

          // For MO consumption, never assign a random WASP lot.
          // If the exact Katana lot cannot be resolved or does not exist in WASP,
          // skip the batch and let the operator reconcile it manually.
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

          var exactBtWasp = waspLookupExactLotAndDate(
            ingSku,
            FLOWS.MO_INGREDIENT_LOCATION,
            btLot,
            CONFIG.WASP_SITE,
            btExpiry || ''
          );

          if (!exactBtWasp || exactBtWasp.ambiguous) {
            var btSkipReason = exactBtWasp && exactBtWasp.ambiguous
              ? 'Exact Katana lot ambiguous in WASP'
              : 'Exact Katana lot not in WASP';
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

          if ((!btExpiry || exactBtWasp.dateToleranceApplied) && exactBtWasp.dateCode) {
            btExpiry = exactBtWasp.dateCode;
          }

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

        var exactIngWasp = null;
        if (ingLot) {
          exactIngWasp = waspLookupExactLotAndDate(
            ingSku,
            FLOWS.MO_INGREDIENT_LOCATION,
            ingLot,
            CONFIG.WASP_SITE,
            ingExpiry || ''
          );

          if (!exactIngWasp || exactIngWasp.ambiguous) {
            var ingSkipReason = exactIngWasp && exactIngWasp.ambiguous
              ? 'Exact Katana lot ambiguous in WASP'
              : 'Exact Katana lot not in WASP';
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

          if ((!ingExpiry || exactIngWasp.dateToleranceApplied) && exactIngWasp.dateCode) {
            ingExpiry = exactIngWasp.dateCode;
          }
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
  saveMOConsumedMap(moOrderNo, consumedMap);

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
      ingAction = FLOWS.MO_INGREDIENT_LOCATION;
      if (ing.lot) ingAction += '  lot:' + ing.lot;
      if (ing.expiry) ingAction += '  exp:' + normalizeBusinessDate_(ing.expiry);
      ingStatus = 'Skip-OK';
      ingErrorMsg = 'Already consumed';
    } else if (isStrictBatchSkip) {
      ingAction = FLOWS.MO_INGREDIENT_LOCATION;
      if (!ing.lot && ing.batchId) ingAction += '  batch_id:' + ing.batchId;
      if (ing.lot) ingAction += '  lot:' + ing.lot;
      if (ing.expiry) ingAction += '  exp:' + normalizeBusinessDate_(ing.expiry);
      ingStatus = 'Skipped';
      ingErrorMsg = ing.strictBatchSkipReason || 'Exact Katana lot skipped';
    } else if (isBatchMismatchSkip) {
      ingAction = FLOWS.MO_INGREDIENT_LOCATION;
      ingStatus = 'Skipped';
      ingErrorMsg = 'Batch mismatch (stale lot)';
    } else if (isInsufficientQty) {
      ingAction = FLOWS.MO_INGREDIENT_LOCATION;
      if (ing.lot) ingAction += '  lot:' + ing.lot;
      if (ing.expiry) ingAction += '  exp:' + normalizeBusinessDate_(ing.expiry);
      ingStatus = 'Skipped';
      ingErrorMsg = 'Not enough in ' + FLOWS.MO_INGREDIENT_LOCATION;
    } else if (ingOk) {
      ingAction = FLOWS.MO_INGREDIENT_LOCATION;
      if (ing.lot) ingAction += '  lot:' + ing.lot;
      if (ing.expiry) ingAction += '  exp:' + normalizeBusinessDate_(ing.expiry);
      ingStatus = 'Consumed';
    } else {
      ingAction = FLOWS.MO_INGREDIENT_LOCATION;
      ingStatus = 'Failed';
      if (ing.result) {
        ingErrorMsg = ing.result.error || (ing.result.response ? parseWaspError(ing.result.response, 'Remove', ing.sku, FLOWS.MO_INGREDIENT_LOCATION) : '');
      }
    }

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
          action: FLOWS.MO_INGREDIENT_LOCATION,
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
          action: ingAction,
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
    var outAction = outLocation;
    if (outOk) {
      if (outLot) outAction += '  lot:' + outLot;
      if (outExpiry) outAction += '  exp:' + normalizeBusinessDate_(outExpiry);
    }
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
      status: outOk ? (outSkipOk ? 'Skip-OK' : 'Produced') : 'Failed',
      error: outError,
      action: outOk ? outAction : '',
      qtyColor: 'green'
    });
  }

  var f4Status = (f4Fail === 0 && f4Skip === 0) ? 'success' : (f4Success === 0 ? 'failed' : 'partial');
  var moRef = mo.order_no || ('MO-' + moId);
  var f4Detail = moRef + '  ' + outputSku + ' x' + outputQuantity;
  if (f4Fail > 0) f4Detail += '  ' + f4Fail + ' error' + (f4Fail > 1 ? 's' : '');
  if (lotNumber) f4Detail += '  ' + lotNumber;
  var f4ExecId = logActivity('F4', f4Detail, f4Status, '→ ' + outputLocation, f4SubItems, {
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
  logFlowDetail('F4', f4ExecId, {
    ref: moRef,
    detail: outputSku + ' x' + outputQuantity + ' → ' + outputLocation,
    status: f4Status === 'success' ? 'Complete' : f4Status === 'failed' ? 'Failed' : 'Partial',
    linkText: moRef,
    linkUrl: getKatanaWebUrl('mo', moId)
  }, f4FlowItems);

  sendSlackNotification(
    'MO Done: ' + moRef + '\n' +
    'Stage: ' + stage + '\n' +
    'Output: ' + outputSku + ' x' + outputQuantity + ' → ' + outputLocation + '\n' +
    'Ingredients removed: ' + results.ingredientsRemoved.length
  );

  // Clean up consumed tracking if MO fully completed (no need for retry)
  if (f4Status === 'success') {
    saveMOConsumedMap(moRef, {});
  }

  // Save reversal snapshot (F4 revert support)
  // Stores what was successfully committed to WASP so it can be undone on MO revert/delete.
  // Includes skippedRetry items — those were consumed on a prior run and still need reversal.
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
  if (snapshotIng.length > 0 || snapshotOut) {
    var snapshotObj = { moRef: moRef, moId: moId, stage: stage, ingredients: snapshotIng, output: snapshotOut };
    saveMOSnapshot(moRef, moId, snapshotObj);
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
 * Only acts on IN_PROGRESS status — stages ingredients to PRODUCTION
 */
function handleManufacturingOrderUpdated(payload) {
  var moId = payload.object ? payload.object.id : null;

  if (!moId) {
    return { status: 'error', message: 'No MO ID in webhook' };
  }

  // Fetch MO to check status
  var moData = fetchKatanaMO(moId);
  if (!moData) {
    return { status: 'error', message: 'Failed to fetch MO' };
  }

  var mo = moData.data ? moData.data : moData;
  var moStatus = (mo.status || '').toUpperCase();

  // DONE is handled by the separate manufacturing_order.done webhook
  if (moStatus === 'DONE') {
    return { status: 'ignored', reason: 'MO status is DONE — handled by manufacturing_order.done' };
  }

  // Check for F4 snapshot BEFORE the IN_PROGRESS staging branch.
  // Snapshot means MO was previously completed (Done) and is now being reverted —
  // this applies to ANY non-DONE status including IN_PROGRESS (Work in progress),
  // NOT_STARTED, BLOCKED, PARTIALLY_COMPLETE, etc.
  var moRefRev = mo.order_no || ('MO-' + moId);
  var snapshotStrRev = PropertiesService.getScriptProperties().getProperty('mo_snapshot_' + moRefRev);
  if (snapshotStrRev) {
    if (getHotfixFlag_('F4_CONFIRM_STATUS_REVERT')) {
      var confirmedMO = refetchKatanaEntityAfterConfirmDelay_(fetchKatanaMO, moId);
      if (!confirmedMO) {
        return { status: 'skipped', reason: 'F4 revert confirmation fetch failed for ' + moRefRev };
      }
      mo = confirmedMO;
      moStatus = (mo.status || '').toUpperCase();
      if (moStatus === 'DONE') {
        return { status: 'ignored', reason: 'F4 revert not confirmed; MO returned to DONE' };
      }
      moRefRev = mo.order_no || ('MO-' + moId);
      snapshotStrRev = PropertiesService.getScriptProperties().getProperty('mo_snapshot_' + moRefRev) || snapshotStrRev;
    }
    return reverseMOSnapshot(moId, moRefRev, snapshotStrRev, 'status_change:' + moStatus);
  }

  // F6 staging disabled — no action when MO goes to IN_PROGRESS. Only revert (via snapshot above) is supported.
  return { status: 'ignored', reason: 'F6 staging disabled: no action on status ' + moStatus };

  // Dedup: skip if already staged
  var cache = CacheService.getScriptCache();
  var dedupKey = 'mo_staging_' + moId;
  if (cache.get(dedupKey)) {
    return { status: 'skipped', reason: 'MO already staged', moId: moId };
  }
  cache.put(dedupKey, 'true', 600); // 10-minute dedup window

  // Get product info
  var variantId = mo.variant_id;
  var productName = '';
  var outputSku = '';

  if (variantId) {
    var variantData = fetchKatanaVariant(variantId);
    var variant = variantData && variantData.data ? variantData.data : variantData;
    outputSku = variant ? (variant.sku || '') : '';
    productName = variant ? (variant.product ? variant.product.name : variant.name || '') : '';
    if (!productName) productName = mo.product_name || '';
  }

  // Fetch ingredients
  var ingredientsData = fetchKatanaMOIngredients(moId);
  var ingredients = ingredientsData && ingredientsData.data ? ingredientsData.data : (ingredientsData || []);

  var results = [];

  for (var i = 0; i < ingredients.length; i++) {
    var ing = ingredients[i];
    var ingSku = '';

    if (ing.variant_id) {
      var ingVariantData = fetchKatanaVariant(ing.variant_id);
      var ingVariant = ingVariantData && ingVariantData.data ? ingVariantData.data : ingVariantData;
      ingSku = ingVariant ? (ingVariant.sku || '') : '';
    }

    var ingQty = ing.quantity || 0;
    if (!ingSku || ingQty <= 0) continue;

    // Extract lot from Katana batch data if available
    var ingLot = extractIngredientBatchNumber(ing);
    var ingExpiry = normalizeBusinessDate_(extractIngredientExpiryDate(ing) || '');

    // Step 1: Try to remove from RECEIVING-DOCK (where PO items land)
    markSyncedToWasp(ingSku, LOCATIONS.RECEIVING, 'remove');

    var removeResult;
    if (ingLot) {
      removeResult = waspRemoveInventoryWithLot(
        ingSku, ingQty, LOCATIONS.RECEIVING, ingLot,
        '[F6-STAGE] MO ' + moId + ' staging: ' + productName,
        null, ingExpiry
      );
    } else {
      removeResult = waspRemoveInventory(
        ingSku, ingQty, LOCATIONS.RECEIVING,
        '[F6-STAGE] MO ' + moId + ' staging: ' + productName
      );

      // -57041 fallback: item is lot-tracked in WASP, look up lot and retry
      if (!removeResult.success && removeResult.response && removeResult.response.indexOf('-57041') >= 0) {
        var waspLotInfo = waspLookupItemLotAndDate(ingSku, LOCATIONS.RECEIVING, CONFIG.WASP_SITE);
        if (waspLotInfo && waspLotInfo.lot) {
          removeResult = waspRemoveInventoryWithLot(
            ingSku, ingQty, LOCATIONS.RECEIVING, waspLotInfo.lot,
            '[F6-STAGE] MO ' + moId + ' staging: ' + productName + ' (lot from WASP)',
            null, waspLotInfo.dateCode
          );
          ingLot = waspLotInfo.lot;
          ingExpiry = waspLotInfo.dateCode || ingExpiry;
        }
      }
    }

    var removeOk = removeResult && removeResult.success;

    // Step 2: Add to PRODUCTION (always attempt, even if removal failed)
    markSyncedToWasp(ingSku, FLOWS.MO_INGREDIENT_LOCATION, 'add');

    var addResult;
    if (ingLot) {
      addResult = waspAddInventoryWithLot(
        ingSku, ingQty, FLOWS.MO_INGREDIENT_LOCATION, ingLot, ingExpiry,
        '[F6-STAGE] MO ' + moId + ' staging: ' + productName
      );
    } else {
      addResult = waspAddInventory(
        ingSku, ingQty, FLOWS.MO_INGREDIENT_LOCATION,
        '[F6-STAGE] MO ' + moId + ' staging: ' + productName
      );
    }

    results.push({
      sku: ingSku,
      quantity: ingQty,
      lot: ingLot || '',
      expiry: ingExpiry || '',
      removeOk: removeOk,
      addResult: addResult,
      uom: resolveVariantUom(ingVariant)
    });
  }

  // Activity Log — F6 Staging
  var f6Success = 0;
  var f6Fail = 0;
  var f6SubItems = [];

  for (var r = 0; r < results.length; r++) {
    var res = results[r];
    var addOk = res.addResult && res.addResult.success;
    if (addOk) f6Success++; else f6Fail++;

    var action = '';
    if (addOk && res.removeOk) {
      action = 'moved ' + LOCATIONS.RECEIVING + ' → ' + FLOWS.MO_INGREDIENT_LOCATION;
    } else if (addOk && !res.removeOk) {
      action = 'added to ' + FLOWS.MO_INGREDIENT_LOCATION + ' (not at ' + LOCATIONS.RECEIVING + ')';
    }
    if (res.lot) action += ' (lot: ' + res.lot + ')';
    if (res.expiry) action += ' (exp: ' + res.expiry + ')';

    f6SubItems.push({
      sku: res.sku,
      qty: res.quantity,
      uom: res.uom || '',
      success: addOk,
      status: addOk ? 'Staged' : 'Failed',
      error: addOk ? '' : (res.addResult ? res.addResult.error : ''),
      action: action,
      qtyColor: 'green'
    });
  }

  var f6Status = f6Fail === 0 ? 'success' : f6Success === 0 ? 'failed' : 'partial';
  var moRef = mo.order_no || ('MO-' + moId);
  var f6Detail = moRef + '  ' + outputSku + '  ' + results.length + ' ingredient' + (results.length !== 1 ? 's' : '') + ' staged';
  if (f6Fail > 0) f6Detail += '  ' + f6Fail + ' error' + (f6Fail > 1 ? 's' : '');

  var f6ExecId = logActivity('F6', f6Detail, f6Status, '→ ' + FLOWS.MO_INGREDIENT_LOCATION, f6SubItems, {
    text: moRef,
    url: getKatanaWebUrl('mo', moId)
  });

  // Log details to F4 tab (same MO lifecycle)
  var f6FlowItems = [];
  for (var fl = 0; fl < results.length; fl++) {
    var fRes = results[fl];
    var fAddOk = fRes.addResult && fRes.addResult.success;
    f6FlowItems.push({
      sku: fRes.sku,
      qty: fRes.quantity,
      uom: fRes.uom || '',
      detail: (fRes.removeOk ? LOCATIONS.RECEIVING + ' → ' : '→ ') + FLOWS.MO_INGREDIENT_LOCATION
        + (fRes.lot ? ' (lot: ' + fRes.lot + ')' : '')
        + (fRes.expiry ? ' (exp: ' + fRes.expiry + ')' : ''),
      status: fAddOk ? 'Complete' : 'Failed',
      error: fAddOk ? '' : (fRes.addResult ? fRes.addResult.error : ''),
      qtyColor: 'green'
    });
  }
  logFlowDetail('F4', f6ExecId, {
    ref: moRef,
    detail: results.length + ' ingredient' + (results.length !== 1 ? 's' : '') + ' staged → ' + FLOWS.MO_INGREDIENT_LOCATION,
    status: f6Status === 'success' ? 'Complete' : f6Status === 'failed' ? 'Failed' : 'Partial',
    linkText: moRef,
    linkUrl: getKatanaWebUrl('mo', moId)
  }, f6FlowItems);

  sendSlackNotification(
    'MO Staging: ' + moRef + '\n' +
    'Product: ' + outputSku + ' (' + productName + ')\n' +
    'Ingredients staged: ' + results.length + ' → ' + FLOWS.MO_INGREDIENT_LOCATION
  );

  return {
    status: 'processed',
    moId: moId,
    moRef: moRef,
    ingredientsStaged: results.length,
    results: results
  };
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

  var snapshotStr = scriptProps.getProperty('mo_snapshot_' + moRef);
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
function reverseMOSnapshot(moId, moRef, snapshotStr, trigger) {
  var snapshot;
  try {
    snapshot = JSON.parse(snapshotStr);
  } catch (e) {
    return { status: 'error', message: 'Invalid snapshot JSON for ' + moRef };
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
  saveMOConsumedMap(moRef, {});

  // Build Activity log sub-items
  var revSubItems = [];
  for (var ri = 0; ri < reversalResults.length; ri++) {
    var rr = reversalResults[ri];
    var rrOk = rr.result && rr.result.success;
    var rrIsAdd = (rr.action === 'add_ingredient');
    var rrLoc = rrIsAdd ? FLOWS.MO_INGREDIENT_LOCATION : (output ? output.location : FLOWS.MO_OUTPUT_LOCATION);
    var rrAction = rrIsAdd ? ('→ ' + rrLoc) : ('← ' + rrLoc);
    if (rr.lot) rrAction += '  lot:' + rr.lot;
    revSubItems.push({
      sku: rr.sku,
      qty: rr.qty,
      uom: normalizeUom(rr.uom || ''),
      success: rrOk,
      status: rrOk ? (rrIsAdd ? 'Restored' : 'Removed') : 'Failed',
      error: rrOk ? '' : (rr.result ? (rr.result.error || parseWaspError(rr.result.response, rrIsAdd ? 'Add' : 'Remove', rr.sku)) : ''),
      action: rrOk ? rrAction : '',
      qtyColor: rrIsAdd ? 'green' : 'red'
    });
  }

  var revStatus = revFail === 0 ? 'reverted' : revSuccess === 0 ? 'failed' : 'partial';
  // Clean trigger: "status_change:NOT_STARTED" → "Not Started", "deleted" → "deleted"
  var cleanTrigger = (trigger || '').replace(/^status_change:/, '').toLowerCase().replace(/_/g, ' ');
  var revDetail = moRef + '  reverted (' + cleanTrigger + ')';
  if (ingredients.length > 0) revDetail += '  ' + ingredients.length + ' ingredient' + (ingredients.length !== 1 ? 's' : '') + ' restored';
  if (revFail > 0) revDetail += '  ' + revFail + ' error' + (revFail > 1 ? 's' : '');

  logActivity('F4', revDetail, revStatus, '', revSubItems, {
    text: moRef,
    url: getKatanaWebUrl('mo', moId || snapshot.moId || '')
  });

  sendSlackNotification(
    'MO Reverted: ' + moRef + '\n' +
    'Trigger: ' + trigger + '\n' +
    'Ingredients restored: ' + ingredients.length + ', Output removed: ' + (output ? 1 : 0)
  );

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

  markSyncedToWasp(sku, location, reverseAction);

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
  var revDetail = 'SA ' + saId + '  reverted (' + (originalAction === 'add' ? 'add undone' : 'remove undone') + ')';

  var subItemAction = location;
  if (lot) subItemAction += '  lot:' + lot;
  if (expiry) subItemAction += '  exp:' + expiry;

  logActivity('F2', revDetail, revStatus, '', [{
    sku: sku,
    qty: qty,
    uom: uom || '',
    success: revOk,
    status: revOk ? (reverseAction === 'add' ? 'Restored' : 'Removed') : 'Failed',
    error: revOk ? '' : (reverseResult && reverseResult.response ? parseWaspError(reverseResult.response, reverseAction === 'add' ? 'Add' : 'Remove', sku) : ''),
    action: subItemAction,
    qtyColor: reverseAction === 'add' ? 'green' : 'red'
  }]);

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
      action: ok ? FLOWS.PICK_FROM_LOCATION : '',
      qtyColor: 'green'
    });
  }
  var f5cStatus = f5cFail === 0 ? 'success' : f5cSuccess === 0 ? 'failed' : 'partial';
  var f5cDetail;
  if (results.length === 1) {
    f5cDetail = orderNumber + '  CANCELLED  ' + results[0].sku + ' x' + results[0].quantity + ' returned';
    if (f5cFail > 0 && f5cSubItems[0] && f5cSubItems[0].error) f5cDetail += '  ' + f5cSubItems[0].error;
  } else {
    f5cDetail = orderNumber + '  CANCELLED  ' + results.length + ' items returned';
    if (f5cFail > 0) f5cDetail += '  ' + f5cFail + ' error' + (f5cFail > 1 ? 's' : '');
  }
  var f5cExecId = logActivity('F5', f5cDetail, f5cStatus, '→ ' + FLOWS.PICK_FROM_LOCATION, f5cSubItems.length > 1 ? f5cSubItems : null, {
    text: orderNumber,
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
  if (objStatus === 'DELIVERED') {
    var liveSoData = fetchKatanaSalesOrder(soId);
    var liveSo = liveSoData && liveSoData.data ? liveSoData.data : liveSoData;
    var isAmazonUSUpdate = isAmazonUSSalesOrder_(liveSo, payload.object || {});
    if (isAmazonUSUpdate) {
      // Guard: if F6 was already successfully processed, ignore subsequent DELIVERED updates.
      // f6_delivered_{soId} is stored by handleSalesOrderDelivered on success.
      // If F6 failed (property absent), allow retry — SO will re-trigger on next Katana update.
      var f6AlreadyDone = PropertiesService.getScriptProperties().getProperty('f6_delivered_' + soId);
      if (f6AlreadyDone) {
        return { status: 'ignored', reason: 'F6 already completed for SO ' + soId };
      }
      return handleSalesOrderDelivered(payload);
    }
    return { status: 'ignored', reason: 'DELIVERED via .updated — not an Amazon US SO' };
  }

  // Check for F6 revert — SO was F6-processed and is no longer DELIVERED
  var f6Stored = PropertiesService.getScriptProperties().getProperty('f6_delivered_' + soId);
  if (f6Stored) {
    var soData = fetchKatanaSalesOrder(soId);
    if (soData) {
      var so = soData.data ? soData.data : soData;
      var currentStatus = (so.status || '').toUpperCase();
      if (currentStatus !== 'DELIVERED') {
        if (getHotfixFlag_('F6_CONFIRM_STATUS_REVERT')) {
          var confirmedF6SO = refetchKatanaEntityAfterConfirmDelay_(fetchKatanaSalesOrder, soId);
          if (!confirmedF6SO) {
            return { status: 'skipped', reason: 'F6 revert confirmation fetch failed for SO ' + soId };
          }
          currentStatus = (confirmedF6SO.status || '').toUpperCase();
          if (currentStatus === 'DELIVERED') {
            return { status: 'ignored', reason: 'F6 revert not confirmed; SO returned to DELIVERED' };
          }
        }
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
 * Removes from AMAZON-FBA-USA, adds back to SHIPPING-DOCK.
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
      markSyncedToWasp(sku, FLOWS.AMAZON_FBA_WASP_LOCATION, 'remove');
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

      // Step 2: Add back to SHIPPING-DOCK (same lot if applicable)
      var addResult = null;
      if (removeResult && removeResult.success) {
        markSyncedToWasp(sku, FLOWS.AMAZON_TRANSFER_LOCATION, 'add');
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
      action: res.ok ? rLocationLabel : '',
      qtyColor: 'green'
    });
  }

  var rStatus = rFail === 0 ? 'reverted' : rSuccess === 0 ? 'failed' : 'partial';
  var rDetail;
  if (results.length === 1) {
    rDetail = orderNo + '  REVERTED  ' + results[0].sku + ' x' + results[0].qty;
    if (rFail > 0 && rSubItems[0] && rSubItems[0].error) rDetail += '  ' + rSubItems[0].error;
  } else {
    rDetail = orderNo + '  REVERTED  ' + results.length + ' items';
    if (rFail > 0) rDetail += '  ' + rFail + ' error' + (rFail > 1 ? 's' : '');
  }

  logActivity('F6', rDetail, rStatus, '\u2192 ' + rLocationLabel, rSubItems, {
    text: orderNo,
    url: getKatanaWebUrl('so', soId)
  });

  return {
    status: rStatus,
    soId: soId,
    flow: 'F6',
    results: results
  };
}
