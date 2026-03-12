// ============================================
// 99_TestSetup.gs - ONE-TIME TEST INVENTORY SETUP
// ============================================
// Reads Katana recipes to find ingredients & finished goods,
// pulls real batch/lot numbers from Katana,
// then seeds WASP with matching lot data at the correct locations.
// DELETE THIS FILE after setup is done.
// ============================================

var TEST_QTY = 100;  // How much of each item to add

// ============================================
// F5 GUARD TEST — run testF5ShopifyGuard() from GAS editor
// ============================================

/**
 * Safe test — no WASP inventory changes.
 * Simulates sales_order.delivered with source=SHOPIFY and verifies it is ignored.
 * Also simulates a non-Shopify SO to confirm F3 still proceeds.
 */
function testF5ShopifyGuard() {
  Logger.log('=== F5 Shopify Guard Test ===');

  // Test 1: Shopify SO — should be ignored
  var shopifyPayload = {
    action: 'sales_order.delivered',
    object: { id: 41627222, status: 'DELIVERED', source: 'SHOPIFY', order_no: 'SO-52' }
  };
  var r1 = routeWebhook(shopifyPayload);
  Logger.log('Test 1 (Shopify SO source=SHOPIFY): ' + JSON.stringify(r1));
  Logger.log(r1.status === 'ignored' ? 'PASS' : 'FAIL — expected ignored');

  // Test 2: sales_order.updated DELIVERED for non-Amazon SO — should be ignored
  var regularPayload = {
    action: 'sales_order.updated',
    object: { id: 99999999, status: 'DELIVERED', customer_id: 99999, order_no: 'SO-99' }
  };
  var r2 = routeWebhook(regularPayload);
  Logger.log('Test 2 (non-Amazon .updated DELIVERED): ' + JSON.stringify(r2));
  Logger.log(r2.status === 'ignored' ? 'PASS' : 'FAIL — expected ignored');

  // Test 3: sales_order.created — should be ignored
  var createdPayload = {
    action: 'sales_order.created',
    object: { id: 12345678 }
  };
  var r3 = routeWebhook(createdPayload);
  Logger.log('Test 3 (sales_order.created): ' + JSON.stringify(r3));
  Logger.log(r3.status === 'ignored' ? 'PASS' : 'FAIL — expected ignored');

  Logger.log('=== Done. Check above for PASS/FAIL ===');
}

// ============================================
// SOURCE FIELD PROBE — run testCheckSOSource()
// ============================================

/**
 * Fetches two SOs and logs their source field:
 *   41631227 — came from Shopify
 *   41585921 — manually created in Katana
 */
function testCheckSOSource() {
  var ids = [41631227, 41585921];
  for (var i = 0; i < ids.length; i++) {
    var soData = fetchKatanaSalesOrder(ids[i]);
    var so = soData && soData.data ? soData.data : soData;
    Logger.log('SO ' + ids[i] + ' order_no=' + JSON.stringify(so ? so.order_no : null) +
      ' | source=' + JSON.stringify(so ? so.source : null) +
      ' | customer_id=' + JSON.stringify(so ? so.customer_id : null));
  }
}

// ============================================
// F6 DIAGNOSTIC — run testF6WithSO52() from GAS editor
// ============================================

/**
 * Simulates a sales_order.updated DELIVERED event for SO-52 (Katana ID 41627222).
 * Checks whether F6 Amazon FBA detection fires correctly after the fix.
 * Run from GAS editor → select testF6WithSO52 → Run, then check Executions logs.
 */
function testF6WithSO52() {
  var soId = 41627222;

  Logger.log('=== F6 Test: SO-52 (ID ' + soId + ') ===');

  // Step 1: Fetch the live SO to see its current customer/status
  var soData = fetchKatanaSalesOrder(soId);
  var so = soData && soData.data ? soData.data : soData;
  if (!so) {
    Logger.log('ERROR: could not fetch SO ' + soId);
    return;
  }
  Logger.log('SO status       : ' + so.status);
  Logger.log('SO customer     : ' + so.customer_name);
  Logger.log('SO order_no     : ' + so.order_no);

  // Step 2: Build a synthetic sales_order.updated DELIVERED payload
  var syntheticPayload = {
    action: 'sales_order.updated',
    object: {
      id: soId,
      status: 'DELIVERED',
      order_no: so.order_no || '',
      customer_name: so.customer_name || ''
    }
  };

  Logger.log('Synthetic payload: ' + JSON.stringify(syntheticPayload));

  // Step 3: Route it — should now hit handleSalesOrderDelivered via the new DELIVERED branch
  var result = routeWebhook(syntheticPayload);
  Logger.log('routeWebhook result: ' + JSON.stringify(result));
}

// ============================================
// UOM DIAGNOSTIC — run diagUom() from GAS editor
// ============================================

/**
 * Fetches MO 15596054, dumps variant UOM fields for every ingredient + output.
 * Run from GAS editor → select diagUom → Run.
 * Read results in View → Executions → click run → Logs.
 */
function diagUom() {
  var moId = 15596054;

  // Fetch MO header
  var moData = fetchKatanaMO(moId);
  var mo = moData && moData.data ? moData.data : moData;
  if (!mo) { Logger.log('ERROR: could not fetch MO ' + moId); return; }

  Logger.log('=== MO ' + moId + ' / ' + (mo.order_no || '?') + ' ===');

  // Output variant + its product
  if (mo.variant_id) {
    var outVarData = fetchKatanaVariant(mo.variant_id);
    var outVar = outVarData && outVarData.data ? outVarData.data : outVarData;
    Logger.log('--- OUTPUT variant_id=' + mo.variant_id + ' sku=' + (outVar ? outVar.sku : '?') + ' ---');
    Logger.log('  variant keys : ' + (outVar ? Object.keys(outVar).join(' | ') : 'null'));
    if (outVar && outVar.product_id) {
      var outProdData = fetchKatanaProduct(outVar.product_id);
      var outProd = outProdData && outProdData.data ? outProdData.data : outProdData;
      Logger.log('  product keys : ' + (outProd ? Object.keys(outProd).join(' | ') : 'null'));
      Logger.log('  product.uom  : ' + JSON.stringify(outProd ? outProd.uom : 'null'));
      Logger.log('  product.unit : ' + JSON.stringify(outProd ? outProd.unit : 'null'));
      Logger.log('  product.unit_of_measure : ' + JSON.stringify(outProd ? outProd.unit_of_measure : 'null'));
    }
  }

  // One ingredient to check product fields
  var ingsData = fetchKatanaMOIngredients(moId);
  var ings = ingsData && ingsData.data ? ingsData.data : (ingsData || []);
  Logger.log('--- INGREDIENTS (product_id + material_id lookup) ---');
  for (var i = 0; i < Math.min(ings.length, 3); i++) {
    var ing = ings[i];
    if (!ing.variant_id) continue;
    var vData = fetchKatanaVariant(ing.variant_id);
    var v = vData && vData.data ? vData.data : vData;
    Logger.log('  [' + i + '] sku=' + (v ? v.sku : '?') +
      ' product_id=' + (v ? v.product_id : '?') +
      ' material_id=' + (v ? v.material_id : '?'));
    // Try product lookup
    var lookupId = (v && v.product_id) ? v.product_id : (v && v.material_id ? v.material_id : null);
    var lookupEndpoint = (v && v.product_id) ? 'products/' : 'materials/';
    if (lookupId) {
      var pData = katanaApiCall(lookupEndpoint + lookupId);
      var p = pData && pData.data ? pData.data : pData;
      Logger.log('    endpoint : ' + lookupEndpoint + lookupId);
      Logger.log('    keys     : ' + (p ? Object.keys(p).join(' | ') : 'null'));
      Logger.log('    uom      : ' + JSON.stringify(p ? p.uom : 'null'));
    }
  }
  Logger.log('=== DONE ===');
}

/**
 * Checks the actual UoM returned by the Katana API for B-WAX (and a few other
 * key items) so we know exactly what label our resolveVariantUom() will display.
 * Run from GAS editor → select diagBwaxUom → Run → View → Executions → Logs.
 */
function diagBwaxUom() {
  // SKUs to check — covers grams (material), PC (material), and pcs (product) cases
  var skus = ['B-WAX', 'EO-LAV', 'EGG-X', 'FBA-FRAGILE', 'B-PINK-2'];

  Logger.log('=== diagBwaxUom — resolveVariantUom spot-check ===');

  for (var i = 0; i < skus.length; i++) {
    var sku = skus[i];
    Logger.log('--- ' + sku + ' ---');

    // Look up variant by SKU
    var varData = getKatanaVariantBySku(sku);
    if (!varData) {
      Logger.log('  NOT FOUND in Katana');
      continue;
    }

    var variant = varData.data ? varData.data : varData;
    Logger.log('  variant_id   : ' + variant.id);
    Logger.log('  product_id   : ' + variant.product_id);
    Logger.log('  material_id  : ' + variant.material_id);
    Logger.log('  variant.uom  : ' + JSON.stringify(variant.uom));

    // Resolve via product or material endpoint
    if (variant.product_id) {
      var pData = fetchKatanaProduct(variant.product_id);
      var p = pData && pData.data ? pData.data : pData;
      Logger.log('  [product] uom: ' + JSON.stringify(p ? p.uom : 'null'));
    } else if (variant.material_id) {
      var mData = fetchKatanaMaterial(variant.material_id);
      var m = mData && mData.data ? mData.data : mData;
      Logger.log('  [material] uom: ' + JSON.stringify(m ? m.uom : 'null'));
    } else {
      Logger.log('  no product_id or material_id');
    }

    // What resolveVariantUom() would return
    var resolved = resolveVariantUom(variant);
    Logger.log('  resolveVariantUom() => "' + resolved + '"');
  }

  Logger.log('=== DONE ===');
}

/**
 * Scans recent Katana PO rows to find every SKU that has a purchase UoM
 * conversion in practice (purchase_uom_conversion_rate != 1).
 *
 * This catches supplier-level conversions that are NOT stored on the material
 * record itself (e.g. B-WAX: material.purchase_uom is blank, but every PO row
 * for B-WAX has purchase_uom=kg and conversion_rate=1000).
 *
 * Also catches material-level conversions (like SWABS: BOX → EA, rate 100).
 *
 * Output: one line per unique SKU with the conversion details.
 * Run from GAS editor → select diagAllConversions → Run → View → Executions → Logs.
 */
function diagAllConversions() {
  Logger.log('=== PO ROW FIELD KEYS (sample) ===');
  var sampleResult = katanaApiCall('purchase_order_rows?per_page=1');
  var sampleRows = sampleResult && sampleResult.data ? sampleResult.data : [];
  if (sampleRows.length > 0) {
    Logger.log(Object.keys(sampleRows[0]).join(' | '));
  }
  Logger.log('');

  Logger.log('=== SKUS WITH PURCHASE UOM CONVERSION (from PO rows) ===');
  Logger.log('SKU | Stock UoM | Purchase UoM | Conversion Rate | Last seen on PO');
  Logger.log('----------------------------------------------------------------------');

  // Map: variantId → { sku, stockUom, purchUom, rate, poNumber }
  var seen = {};
  var page = 1;

  while (true) {
    var result = katanaApiCall('purchase_order_rows?per_page=100&page=' + page);
    if (!result) { Logger.log('ERROR: API call failed on page ' + page); break; }
    var rows = result.data || result || [];
    if (rows.length === 0) break;

    for (var i = 0; i < rows.length; i++) {
      var row = rows[i];
      var rate = parseFloat(row.purchase_uom_conversion_rate || 1);
      if (rate <= 1) continue; // no conversion — skip

      var vid = row.variant_id;
      if (!vid || seen[vid]) continue; // already recorded

      // Fetch variant to get SKU and stock UoM
      var vData = fetchKatanaVariant(vid);
      var v = vData && vData.data ? vData.data : vData;
      if (!v) continue;

      var sku = v.sku || ('variant:' + vid);
      var stockUom = resolveVariantUom(v) || '?';
      var purchUom = row.purchase_uom || '?';
      var poNum = row.purchase_order_id || '?';

      seen[vid] = { sku: sku, stockUom: stockUom, purchUom: purchUom, rate: rate, po: poNum };
      Logger.log(sku + ' | stock:' + stockUom + ' | purchase:' + purchUom + ' | rate:' + rate + ' | PO#' + poNum);
    }

    if (rows.length < 100) break;
    page++;
  }

  var count = Object.keys(seen).length;
  Logger.log('');
  Logger.log('Found ' + count + ' unique SKUs with a purchase UoM conversion.');
  Logger.log('=== DONE ===');
}

// ============================================
// WASP UOM SYNC FROM KATANA
// ============================================

/**
 * Maps a Katana stock UoM string to the exact WASP StockingUnit value.
 * WASP UoM names must match what's configured in WASP Settings → Units of Measure.
 */
function katanaUomToWaspUnit(katanaUom) {
  if (!katanaUom) return 'Each';
  var lc = (katanaUom + '').toLowerCase().trim();
  if (lc === 'each' || lc === 'ea' || lc === 'pc' || lc === 'pcs') return 'Each';
  if (lc === 'grams' || lc === 'g' || lc === 'gram') return 'grams';
  if (lc === 'kg' || lc === 'kilogram') return 'kg';
  if (lc === 'ml' || lc === 'milliliter' || lc === 'millilitre') return 'mL';
  // Unknown — return as-is so it shows up in the report for manual review
  return katanaUom;
}

/**
 * Complete list of Katana items that have a purchase UoM conversion.
 * Supplier-level conversions don't appear in the Katana API material records —
 * this hardcoded list (from wasp-uom-adjustments.txt) is the source of truth.
 *
 * stockUom = Katana stock UoM = what WASP StockingUnit should be set to.
 * Group 4 (NI-*) items are skipped automatically if not found in WASP.
 */
var KATANA_CONVERSION_ITEMS = [
  // Group 1: Raw materials — stock in grams
  { sku: 'B-WAX',           stockUom: 'g' },
  { sku: 'EO-THY',          stockUom: 'g' },
  { sku: 'EO-LAV',          stockUom: 'g' },
  { sku: 'EO-PEP',          stockUom: 'g' },
  { sku: 'EO-CS',           stockUom: 'g' },
  { sku: 'EO-TT',           stockUom: 'g' },
  { sku: 'EO-CB',           stockUom: 'g' },
  { sku: 'B-PROP',          stockUom: 'g' },
  { sku: 'H-ARNF',          stockUom: 'g' },
  { sku: 'H-COML',          stockUom: 'g' },
  // Group 2: Raw materials — stock in pcs/each
  { sku: 'EGG-X',           stockUom: 'each' },
  { sku: 'INFBAG-160',      stockUom: 'each' },
  // Group 3: Packaging — stock in pcs/each
  { sku: 'FBA-BARCODE',     stockUom: 'each' },
  { sku: 'FBA-FRAGILE',     stockUom: 'each' },
  // Group 4: Non-inventory — stock in each (skip if not in WASP)
  { sku: 'NI-GLOVEMD',      stockUom: 'each' },
  { sku: 'NI-GLOVESM',      stockUom: 'each' },
  { sku: 'NI-GLOVELR',      stockUom: 'each' },
  { sku: 'NI-GLOVECR',      stockUom: 'each' },
  { sku: 'NI-FIRE-BNKT',    stockUom: 'each' },
  { sku: 'NI-NAPKINS',      stockUom: 'each' },
  { sku: 'NI-PAPERTOWEL8',  stockUom: 'each' },
  { sku: 'NI-SCRUB-10',     stockUom: 'each' },
  { sku: 'NI-STEELWOOL',    stockUom: 'each' },
  { sku: 'NI-ATPSWABS',     stockUom: 'each' },
  { sku: 'NI-TOILETPAPER24',stockUom: 'each' }
];

/**
 * Checks every item in KATANA_CONVERSION_ITEMS against WASP and either
 * reports mismatches (DRY_RUN=true) or fixes them (DRY_RUN=false).
 *
 * Run from GAS editor → select syncWaspUomFromKatana → Run → Executions → Logs.
 * Set DRY_RUN = false (line below) when ready to apply changes.
 */
function syncWaspUomFromKatana() {
  var DRY_RUN = true; // ← change to false to actually update WASP

  Logger.log('=== syncWaspUomFromKatana  DRY_RUN=' + DRY_RUN + ' ===');
  Logger.log('Total items to check: ' + KATANA_CONVERSION_ITEMS.length);
  Logger.log('');

  var items = KATANA_CONVERSION_ITEMS;

  if (items.length === 0) {
    Logger.log('No items found — nothing to do.');
    return;
  }

  // ---- Step 2: Check WASP and apply updates -----------------------------
  var waspSearchUrl = CONFIG.WASP_BASE_URL + '/public-api/ic/item/advancedinfosearch';
  var waspUpdateUrl = CONFIG.WASP_BASE_URL + '/public-api/ic/item/updateInventoryItems';

  // Helper: extract item list from any known WASP response shape.
  function parseWaspList(res) {
    if (!res || !res.response) return [];
    var parsed;
    try { parsed = JSON.parse(res.response); } catch (e) { return []; }
    var list = parsed.Data || parsed.data || parsed.Result || parsed.Items ||
               (Array.isArray(parsed) ? parsed : []);
    return Array.isArray(list) ? list : [];
  }

  // Helper: search WASP for an item by SKU across two endpoints/payload styles.
  // waspApiCall returns { success, code, response } — response is a raw JSON string.
  function waspFindItem(sku) {
    // Attempt 1: infosearch with ItemNumber (item-definition endpoint)
    var r1 = waspApiCall(CONFIG.WASP_BASE_URL + '/public-api/ic/item/infosearch',
                         { ItemNumber: sku });
    var list1 = parseWaspList(r1);
    for (var a = 0; a < list1.length; a++) {
      if (String(list1[a].ItemNumber || '').toUpperCase() === sku.toUpperCase()) return list1[a];
    }

    // Attempt 2: inventorysearch with ItemNumber (inventory endpoint — may include item fields)
    var r2 = waspApiCall(CONFIG.WASP_BASE_URL + '/public-api/ic/item/inventorysearch',
                         { ItemNumber: sku, SiteName: CONFIG.WASP_SITE });
    var list2 = parseWaspList(r2);
    for (var b = 0; b < list2.length; b++) {
      if (String(list2[b].ItemNumber || '').toUpperCase() === sku.toUpperCase()) return list2[b];
    }

    return null;
  }


  Logger.log('SKU | Katana stock | WASP StockingUnit | Target | Action');
  Logger.log('---------------------------------------------------------------');

  var alreadyOk = 0, needsUpdate = 0, notInWasp = 0, updateOk = 0, updateFail = 0;

  for (var ii = 0; ii < items.length; ii++) {
    var item = items[ii];
    var target = katanaUomToWaspUnit(item.stockUom);

    var waspItem = waspFindItem(item.sku);

    if (!waspItem) {
      notInWasp++;
      Logger.log(item.sku + ' | ' + item.stockUom + ' | NOT IN WASP | ' + target + ' | SKIP');
      continue;
    }

    var current = waspItem.StockingUnit || waspItem.UnitOfMeasure || waspItem.Uom || '?';

    if (current === target) {
      alreadyOk++;
      Logger.log(item.sku + ' | ' + item.stockUom + ' | ' + current + ' | ' + target + ' | OK');
      continue;
    }

    needsUpdate++;

    if (DRY_RUN) {
      Logger.log(item.sku + ' | ' + item.stockUom + ' | ' + current + ' → ' + target + ' | DRY_RUN (would update)');
    } else {
      Logger.log(item.sku + ' | ' + item.stockUom + ' | ' + current + ' → ' + target + ' | UPDATING...');
      var updatePayload = [{
        ItemNumber:    item.sku,
        StockingUnit:  target,
        PurchaseUnit:  target,
        SalesUnit:     target,
        DimensionInfo: { DimensionUnit: 'Inch', WeightUnit: 'Pound', VolumeUnit: 'Cubic Inch' }
      }];
      var updateRes = waspApiCall(waspUpdateUrl, updatePayload);
      if (updateRes && updateRes.success) {
        updateOk++;
        Logger.log('  → OK');
      } else {
        updateFail++;
        Logger.log('  → FAILED: ' + String(updateRes ? updateRes.response : '').substring(0, 150));
      }
    }
  }

  Logger.log('');
  Logger.log('=== SUMMARY ===');
  Logger.log('Items with conversions in Katana : ' + items.length);
  Logger.log('Already correct in WASP          : ' + alreadyOk);
  Logger.log('Not found in WASP (skip)         : ' + notInWasp);
  Logger.log('Need update                      : ' + needsUpdate);
  if (!DRY_RUN) {
    Logger.log('Updated OK                       : ' + updateOk);
    Logger.log('Update failed                    : ' + updateFail);
  } else {
    Logger.log('(set DRY_RUN=false to apply changes)');
  }
  Logger.log('=== DONE ===');
}

/**
 * STEP 1: Preview — shows what will be added, including batch info
 */
function testSetup_PREVIEW() {
  var plan = buildInventoryPlan();

  Logger.log('===== INGREDIENTS (need at PRODUCTION) =====');
  for (var i = 0; i < plan.ingredients.length; i++) {
    var ing = plan.ingredients[i];
    var batchInfo = ing.batches.length > 0
      ? ' [batches: ' + ing.batches.map(function(b) { return b.lot; }).join(', ') + ']'
      : ' [no batches]';
    Logger.log('  ' + ing.sku + batchInfo + ' — ' + ing.recipeName);
  }

  Logger.log('');
  Logger.log('===== FINISHED GOODS (need at SHIPPING-DOCK) =====');
  for (var j = 0; j < plan.finishedGoods.length; j++) {
    var fg = plan.finishedGoods[j];
    var fgBatchInfo = fg.batches.length > 0
      ? ' [batches: ' + fg.batches.map(function(b) { return b.lot; }).join(', ') + ']'
      : ' [no batches]';
    Logger.log('  ' + fg.sku + fgBatchInfo);
  }

  Logger.log('');
  Logger.log('Total ingredients: ' + plan.ingredients.length);
  Logger.log('Total finished goods: ' + plan.finishedGoods.length);
  Logger.log('');
  Logger.log('Items WITH batches will be added with each real lot number.');
  Logger.log('Items WITHOUT batches will be added with lot NO-LOT.');
  Logger.log('');
  Logger.log('Review the list above. If it looks right, run testSetup_EXECUTE()');
}

/**
 * STEP 2: Execute — adds inventory to WASP with matching batch/lot numbers
 */
function testSetup_EXECUTE() {
  var plan = buildInventoryPlan();
  var results = { success: 0, failed: 0, errors: [] };

  Logger.log('===== ADDING INGREDIENTS TO PRODUCTION =====');
  for (var i = 0; i < plan.ingredients.length; i++) {
    var ing = plan.ingredients[i];
    var ingResults = addWithBatches(ing.sku, ing.batches, LOCATIONS.PRODUCTION, 'ingredient');
    results.success += ingResults.success;
    results.failed += ingResults.failed;
    for (var e = 0; e < ingResults.errors.length; e++) {
      results.errors.push(ingResults.errors[e]);
    }
  }

  Logger.log('');
  Logger.log('===== ADDING FINISHED GOODS TO SHIPPING-DOCK =====');
  for (var j = 0; j < plan.finishedGoods.length; j++) {
    var fg = plan.finishedGoods[j];
    var fgResults = addWithBatches(fg.sku, fg.batches, LOCATIONS.SHIPPING, 'finished good');
    results.success += fgResults.success;
    results.failed += fgResults.failed;
    for (var e2 = 0; e2 < fgResults.errors.length; e2++) {
      results.errors.push(fgResults.errors[e2]);
    }
  }

  Logger.log('');
  Logger.log('===== DONE =====');
  Logger.log('Success: ' + results.success + '  Failed: ' + results.failed);

  if (results.errors.length > 0) {
    Logger.log('');
    Logger.log('ERRORS:');
    for (var e3 = 0; e3 < results.errors.length; e3++) {
      Logger.log('  ' + results.errors[e3]);
    }
  }
}

/**
 * Add inventory with real batch numbers (or NO-LOT if no batches exist)
 */
function addWithBatches(sku, batches, location, label) {
  var out = { success: 0, failed: 0, errors: [] };

  if (batches.length === 0) {
    // No batches — add with default lot
    var result = waspAddInventoryWithLot(sku, TEST_QTY, location, null, null, 'Test setup — ' + label);
    if (result.success) {
      out.success++;
      Logger.log('  OK: ' + sku + ' x' + TEST_QTY + ' [NO-LOT] → ' + location);
    } else {
      out.failed++;
      out.errors.push(sku + ' @ ' + location + ': ' + result.response.substring(0, 100));
      Logger.log('  FAIL: ' + sku + ' [NO-LOT] — ' + result.response.substring(0, 100));
    }
  } else {
    // Add once per batch with the real lot number and expiry
    for (var i = 0; i < batches.length; i++) {
      var batch = batches[i];
      var result2 = waspAddInventoryWithLot(sku, TEST_QTY, location, batch.lot, batch.expiry, 'Test setup — ' + label);
      if (result2.success) {
        out.success++;
        Logger.log('  OK: ' + sku + ' x' + TEST_QTY + ' [' + batch.lot + '] → ' + location);
      } else {
        out.failed++;
        out.errors.push(sku + ' [' + batch.lot + '] @ ' + location + ': ' + result2.response.substring(0, 100));
        Logger.log('  FAIL: ' + sku + ' [' + batch.lot + '] — ' + result2.response.substring(0, 100));
      }
    }
  }

  return out;
}

/**
 * Build the inventory plan from Katana recipes + batch data
 * Returns { ingredients: [{sku, variantId, batches, recipeName}], finishedGoods: [...] }
 */
function buildInventoryPlan() {
  var ingredientMap = {};   // sku → { variantId, outputSku }
  var finishedGoodMap = {}; // sku → { variantId }

  // Cache variant lookups
  var variantCache = {};

  function lookupVariant(variantId) {
    if (!variantId) return null;
    if (variantCache[variantId]) return variantCache[variantId];

    var variant = fetchKatanaVariant(variantId);
    if (variant) {
      var data = variant.data ? variant.data : variant;
      variantCache[variantId] = data;
      return data;
    }
    variantCache[variantId] = null;
    return null;
  }

  // Fetch all recipe rows (paginated)
  var page = 1;
  var hasMore = true;

  while (hasMore) {
    var result = katanaApiCall('recipes?per_page=100&page=' + page);

    if (!result || !result.data || result.data.length === 0) {
      hasMore = false;
      break;
    }

    var rows = result.data;

    for (var i = 0; i < rows.length; i++) {
      var row = rows[i];

      // Output (finished good)
      var outputData = lookupVariant(row.product_variant_id);
      var outputSku = outputData ? (outputData.sku || '') : '';
      if (outputSku && !isSkippedSku(outputSku)) {
        finishedGoodMap[outputSku] = { variantId: row.product_variant_id };
      }

      // Ingredient
      var ingData = lookupVariant(row.ingredient_variant_id);
      var ingSku = ingData ? (ingData.sku || '') : '';
      if (ingSku && !isSkippedSku(ingSku)) {
        ingredientMap[ingSku] = { variantId: row.ingredient_variant_id, outputSku: outputSku };
      }
    }

    if (rows.length < 100) {
      hasMore = false;
    } else {
      page++;
    }
  }

  // Now fetch batch data for each unique variant
  Logger.log('Fetching batch data from Katana...');

  var ingredients = [];
  for (var ingSku2 in ingredientMap) {
    var ingInfo = ingredientMap[ingSku2];
    var batches = fetchBatchesForVariant(ingInfo.variantId);
    ingredients.push({
      sku: ingSku2,
      variantId: ingInfo.variantId,
      batches: batches,
      recipeName: 'used in: ' + ingInfo.outputSku
    });
  }

  var finishedGoods = [];
  for (var fgSku in finishedGoodMap) {
    if (!ingredientMap[fgSku]) {
      var fgInfo = finishedGoodMap[fgSku];
      var fgBatches = fetchBatchesForVariant(fgInfo.variantId);
      finishedGoods.push({
        sku: fgSku,
        variantId: fgInfo.variantId,
        batches: fgBatches,
        recipeName: ''
      });
    }
  }

  return { ingredients: ingredients, finishedGoods: finishedGoods };
}

/**
 * Fetch all batch/lot numbers for a variant from Katana
 * Returns [{lot: 'BATCH-123', expiry: '2027-01-01'}, ...]
 */
// ============================================
// F5 DIAGNOSTIC — run diagF5() from GAS editor
// ============================================

/**
 * Checks everything that could prevent F5 from working:
 *   1. ShipStation credentials configured
 *   2. SS_SHIP_POLL_LAST timestamp (is polling stuck in the future?)
 *   3. processWebhookQueue trigger installed
 *   4. Live ShipStation API call (are credentials valid?)
 *   5. Recent shipments from ShipStation (last 24h)
 */
function diagF5() {
  var props = PropertiesService.getScriptProperties();
  Logger.log('=== F5 DIAGNOSTIC ===');

  // 1. Credentials
  var apiKey    = props.getProperty('SHIPSTATION_API_KEY');
  var apiSecret = props.getProperty('SHIPSTATION_API_SECRET');
  Logger.log('\n--- 1. Credentials ---');
  Logger.log('SHIPSTATION_API_KEY:    ' + (apiKey    ? 'SET (' + apiKey.substring(0, 6) + '...)' : 'MISSING'));
  Logger.log('SHIPSTATION_API_SECRET: ' + (apiSecret ? 'SET (' + apiSecret.substring(0, 4) + '...)' : 'MISSING'));

  // 2. Poll timestamp
  var lastPoll = props.getProperty('SS_SHIP_POLL_LAST');
  Logger.log('\n--- 2. SS_SHIP_POLL_LAST ---');
  Logger.log('Value: ' + (lastPoll || 'NOT SET'));
  if (lastPoll) {
    var diff = (new Date() - new Date(lastPoll)) / 60000;
    Logger.log('Age:   ' + diff.toFixed(1) + ' minutes ago');
    if (diff < 0) Logger.log('WARNING: timestamp is in the FUTURE — poller will skip all shipments');
  }

  // 3. Triggers
  Logger.log('\n--- 3. Triggers ---');
  var triggers = ScriptApp.getProjectTriggers();
  var found = false;
  for (var t = 0; t < triggers.length; t++) {
    var fn = triggers[t].getHandlerFunction();
    var type = triggers[t].getEventType();
    Logger.log('  ' + fn + ' (' + type + ')');
    if (fn === 'processWebhookQueue') found = true;
  }
  if (!found) Logger.log('WARNING: processWebhookQueue trigger NOT FOUND — polling is disabled');

  // 4. Live API call
  Logger.log('\n--- 4. ShipStation API test ---');
  if (!apiKey || !apiSecret) {
    Logger.log('SKIP — credentials missing');
  } else {
    var creds = Utilities.base64Encode(apiKey + ':' + apiSecret);
    var resp = UrlFetchApp.fetch('https://ssapi.shipstation.com/shipments?pageSize=1', {
      method: 'GET',
      headers: { 'Authorization': 'Basic ' + creds, 'Content-Type': 'application/json' },
      muteHttpExceptions: true
    });
    var code = resp.getResponseCode();
    Logger.log('Response code: ' + code);
    if (code === 200) {
      Logger.log('Credentials OK');
    } else if (code === 401) {
      Logger.log('ERROR: 401 Unauthorized — credentials are wrong');
    } else {
      Logger.log('ERROR: unexpected code — ' + resp.getContentText().substring(0, 200));
    }
  }

  // 5. Recent shipments (last 24h)
  Logger.log('\n--- 5. Recent shipments (last 24h) ---');
  if (!apiKey || !apiSecret) {
    Logger.log('SKIP — credentials missing');
  } else {
    var since = new Date(Date.now() - 86400000).toISOString();
    var creds2 = Utilities.base64Encode(apiKey + ':' + apiSecret);
    var resp2 = UrlFetchApp.fetch(
      'https://ssapi.shipstation.com/shipments?shipDateStart=' + encodeURIComponent(since) + '&includeShipmentItems=true&pageSize=10',
      {
        method: 'GET',
        headers: { 'Authorization': 'Basic ' + creds2, 'Content-Type': 'application/json' },
        muteHttpExceptions: true
      }
    );
    var code2 = resp2.getResponseCode();
    if (code2 !== 200) {
      Logger.log('API call failed: ' + code2);
    } else {
      var data = JSON.parse(resp2.getContentText());
      var shipments = data.shipments || [];
      Logger.log('Shipments found: ' + shipments.length);
      for (var s = 0; s < shipments.length; s++) {
        var sh = shipments[s];
        var items = sh.shipmentItems || sh.items || [];
        var cached = CacheService.getScriptCache().get('ss_shipped_' + sh.shipmentId);
        Logger.log('  #' + sh.orderNumber + '  shipmentId=' + sh.shipmentId +
          '  items=' + items.length +
          '  voided=' + (sh.voidDate ? 'YES' : 'no') +
          '  cached=' + (cached ? 'YES (already processed)' : 'no'));
      }
      if (shipments.length === 0) {
        Logger.log('No shipments in last 24h — nothing to process');
      }
    }
  }

  Logger.log('\n=== DIAGNOSTIC COMPLETE ===');
}

// ============================================
// F5 BACKFILL — backfillF5Shipments(daysBack)
// Run manually from GAS editor to catch missed ShipStation shipments.
// ============================================

/**
 * Backfill missed ShipStation (F5) shipments for the last N days.
 * Fetches ALL shipments in the date range using pagination (handles >100 shipments).
 * Each shipment goes through the same processShipment() as live polling.
 * Voided shipments are skipped. Already-processed shipments (in CacheService <24h) are skipped.
 *
 * Run manually: backfillF5Shipments()        — last 30 days (default)
 *               backfillF5Shipments(60)      — last 60 days
 *               backfillF5Shipments(7)       — last 7 days
 *
 * After running, check View → Logs and the Activity tab for new F5 entries.
 */
function backfillF5Shipments(daysBack) {
  daysBack = daysBack || 30;

  var execId = 'BKF-' + Utilities.getUuid().substring(0, 8);
  var since = new Date(Date.now() - (daysBack * 86400000)).toISOString();

  Logger.log('=== backfillF5Shipments START ===');
  Logger.log('execId: ' + execId);
  Logger.log('Backfilling shipments since: ' + since + ' (' + daysBack + ' days ago)');

  var page = 1;
  var pageSize = 100;
  var totalFetched = 0;
  var totalProcessed = 0;
  var totalSkipped = 0;
  var totalFailed = 0;
  var hasMore = true;

  while (hasMore) {
    var endpoint = '/shipments'
      + '?shipDateStart=' + encodeURIComponent(since)
      + '&includeShipmentItems=true'
      + '&pageSize=' + pageSize
      + '&page=' + page
      + '&sortBy=ShipDate&sortDir=ASC';

    var result = callShipStationAPI(endpoint, 'GET');

    if (result.code !== 200) {
      Logger.log('ERROR: ShipStation API returned ' + result.code + ' on page ' + page);
      break;
    }

    var data = result.data || {};
    var shipments = data.shipments || [];
    var pages = data.pages || 1;

    Logger.log('Page ' + page + '/' + pages + ': ' + shipments.length + ' shipments');
    totalFetched += shipments.length;

    for (var i = 0; i < shipments.length; i++) {
      var sh = shipments[i];
      if (sh.voidDate) continue; // Voided shipments handled by pollVoidedLabels

      var r = processShipment(sh, execId);
      if (r.skipped) {
        totalSkipped++;
      } else if (r.success) {
        totalProcessed++;
      } else {
        totalFailed++;
      }
    }

    if (page >= pages || shipments.length < pageSize) {
      hasMore = false;
    } else {
      page++;
      Utilities.sleep(500); // Brief pause between pages to respect rate limits
    }
  }

  Logger.log('=== backfillF5Shipments DONE ===');
  Logger.log('Days backfilled: ' + daysBack);
  Logger.log('Total shipments fetched: ' + totalFetched);
  Logger.log('New (processed): ' + totalProcessed);
  Logger.log('Skipped (already done in <24h): ' + totalSkipped);
  Logger.log('Failed: ' + totalFailed);
  if (totalProcessed > 0) Logger.log('Check Activity tab for new F5 entries.');
}

// ============================================
// CLEANUP: Remove ASM (auto-assembly) entries from Activity + F4 tab
// Run once from GAS editor: removeAsmActivityEntries()
// ============================================

/**
 * Deletes all auto-assembly (ASM) MO entries from:
 *   - Activity tab (header row + all sub-item rows)
 *   - F4 Manufacturing tab (header row + all sub-item rows)
 * Safe to run multiple times — only removes rows containing " ASM ".
 */
function removeAsmActivityEntries() {
  var ss = SpreadsheetApp.openById(CONFIG.DEBUG_SHEET_ID);

  var removed = 0;

  // Helper: delete ASM blocks from a sheet.
  // Header rows have WK- in col A; sub-items have empty col A.
  // Col D contains the details text with "ASM" for assembly MOs.
  function cleanSheet(sheetName, detailCol) {
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) { Logger.log('Sheet not found: ' + sheetName); return; }

    var lastRow = sheet.getLastRow();
    if (lastRow < 4) return;

    var data = sheet.getRange(1, 1, lastRow, detailCol).getValues();

    // Collect row ranges to delete (1-based). Go bottom-to-top to avoid index shift.
    var toDelete = [];

    for (var i = lastRow - 1; i >= 3; i--) {
      var colA = String(data[i][0]).trim();
      var colD = String(data[i][detailCol - 1]).trim();

      // Only act on header rows for ASM MOs
      if (colA.indexOf('WK-') !== 0) continue;
      if (colD.indexOf(' ASM ') < 0) continue;

      // Collect this header row + all following sub-item rows
      var startRow = i + 1; // 1-based
      var endRow = startRow;
      for (var j = i + 1; j < lastRow; j++) {
        var subA = String(data[j][0]).trim();
        if (subA.indexOf('WK-') === 0) break;
        endRow = j + 1; // 1-based
      }

      toDelete.push({ start: startRow, count: endRow - startRow + 1 });
    }

    // Delete from bottom to top
    for (var d = 0; d < toDelete.length; d++) {
      sheet.deleteRows(toDelete[d].start, toDelete[d].count);
      removed += toDelete[d].count;
      Logger.log('  Deleted ' + toDelete[d].count + ' rows at row ' + toDelete[d].start + ' in ' + sheetName);
    }
  }

  Logger.log('=== removeAsmActivityEntries ===');
  cleanSheet('Activity', 4);
  cleanSheet('F4 Manufacturing', 4);
  Logger.log('Done. Total rows removed: ' + removed);
}

function fetchBatchesForVariant(variantId) {
  if (!variantId) return [];

  var result = katanaApiCall('batch_stocks?variant_id=' + variantId + '&per_page=100');

  if (!result || !result.data || result.data.length === 0) {
    return [];
  }

  var batches = [];
  for (var i = 0; i < result.data.length; i++) {
    var batch = result.data[i];
    var lot = batch.batch_number || batch.batch_nr || batch.nr || batch.number || '';
    var expiry = batch.expiration_date || batch.expiry_date || '';
    if (expiry) expiry = expiry.split('T')[0];

    if (lot) {
      batches.push({ lot: lot, expiry: expiry });
    }
  }

  return batches;
}

// ============================================
// F2 SINGLE-ITEM TEST — run testF2OneItem() from GAS editor
// ============================================

/**
 * Quick single-item F2 test: one add + one remove for 4OZSEAL.
 * Net Katana effect = zero. Runtime ~20s.
 * Change TEST_SKU below if needed.
 */
function testF2OneItem() {
  var TEST_SKU = '4OZSEAL';
  var TEST_QTY = 1;
  var TEST_LOC = 'PRODUCTION';
  var TEST_SITE = CONFIG.WASP_SITE || 'MMH Kelowna';

  function payload(sku, qty) {
    return { ItemNumber: sku, Quantity: String(qty), SiteName: TEST_SITE, LocationCode: TEST_LOC, Notes: 'F2 DIAG TEST — safe to ignore' };
  }

  Logger.log('=== F2 Single-Item Test: ' + TEST_SKU + ' ===');

  Logger.log('--- Add x' + TEST_QTY + ' ---');
  var r1 = handleWaspQuantityAdded(payload(TEST_SKU, TEST_QTY));
  Logger.log('Result: ' + JSON.stringify(r1));

  Logger.log('--- Remove x' + TEST_QTY + ' ---');
  var r2 = handleWaspQuantityRemoved(payload(TEST_SKU, TEST_QTY));
  Logger.log('Result: ' + JSON.stringify(r2));

  Logger.log('=== DONE — check Activity tab + F2 Adjustments tab ===');
}

// ============================================
// F2 CALLOUT TEST — run testF2Callouts() from GAS editor
// ============================================

/**
 * Tests all 4 F2 callout scenarios without touching WASP manually.
 * Simulates WASP quantity_added / quantity_removed callouts directly.
 * Creates real Katana SAs — net qty change is zero (+1 then -1 per SKU).
 *
 * Tests:
 *   1. Add (no lot)
 *   2. Remove (no lot)
 *   3. Add (with lot)    — SKU auto-discovered from Katana batch_stocks
 *   4. Remove (with lot)
 *
 * Run: GAS editor → select testF2Callouts → Run
 * Then check: Activity tab, F2 Adjustments tab, Katana SA list
 * Each test takes ~10s (batch window) — total runtime ~45s.
 */
function testF2Callouts() {
  Logger.log('=== F2 Callout Test — 4 scenarios ===');
  Logger.log('Creates real Katana SAs. Net qty effect = zero (+1 then -1 per SKU).');
  Logger.log('');

  // ── Config ─────────────────────────────────────────────────────────────────
  // Change SKU_NOLOT if 4OZSEAL is not in your Katana catalog.
  var SKU_NOLOT   = '4OZSEAL';
  var TEST_SITE   = CONFIG.WASP_SITE || 'MMH Kelowna';
  var TEST_LOC    = 'PRODUCTION'; // must NOT be SHOPIFY (F5-owned)
  var TEST_QTY    = 1;

  // Lot-tracked SKU — auto-discovered below
  var SKU_LOT  = null;
  var LOT_NUM  = null;
  var LOT_EXP  = null;

  // ── Auto-discover a lot-tracked SKU from Katana ────────────────────────────
  Logger.log('Searching Katana batch_stocks for a lot-tracked SKU...');
  var bsResult = katanaApiCall('batch_stocks?per_page=20');
  var bsRows   = (bsResult && bsResult.data) ? bsResult.data : [];

  for (var bi = 0; bi < bsRows.length; bi++) {
    var bs  = bsRows[bi];
    var lot = bs.batch_number || bs.batch_nr || bs.nr || '';
    if (!lot || !bs.variant_id) continue;

    var vData = fetchKatanaVariant(bs.variant_id);
    var v     = (vData && vData.data) ? vData.data : vData;
    if (v && v.sku) {
      SKU_LOT = v.sku;
      LOT_NUM = lot;
      LOT_EXP = bs.expiration_date ? bs.expiration_date.split('T')[0] : '';
      Logger.log('Found lot SKU: ' + SKU_LOT + '  lot=' + LOT_NUM + '  exp=' + (LOT_EXP || 'none'));
      break;
    }
  }
  if (!SKU_LOT) Logger.log('WARNING: no lot-tracked SKU found — tests 3 & 4 will be skipped');

  // ── Payload builder ─────────────────────────────────────────────────────────
  function makePayload(sku, qty, lot, exp) {
    var p = {
      ItemNumber:   sku,
      Quantity:     String(qty),
      SiteName:     TEST_SITE,
      LocationCode: TEST_LOC,
      Notes:        'F2 DIAG TEST — safe to ignore'
    };
    if (lot) { p.Lot = lot; if (exp) p.DateCode = exp; }
    return p;
  }

  // ── TEST 1: Add without lot ────────────────────────────────────────────────
  Logger.log('');
  Logger.log('=== TEST 1: Add (no lot)  ' + SKU_NOLOT + ' x' + TEST_QTY + ' ===');
  var r1 = handleWaspQuantityAdded(makePayload(SKU_NOLOT, TEST_QTY, null, null));
  Logger.log('Result: ' + JSON.stringify(r1));
  Logger.log(r1 && r1.status === 'processed' ? 'PASS — SA created in Katana' : 'CHECK — see Activity tab');

  // ── TEST 2: Remove without lot ─────────────────────────────────────────────
  Logger.log('');
  Logger.log('=== TEST 2: Remove (no lot)  ' + SKU_NOLOT + ' x' + TEST_QTY + ' ===');
  var r2 = handleWaspQuantityRemoved(makePayload(SKU_NOLOT, TEST_QTY, null, null));
  Logger.log('Result: ' + JSON.stringify(r2));
  Logger.log(r2 && r2.status === 'processed' ? 'PASS — SA created in Katana' : 'CHECK — see Activity tab');

  if (SKU_LOT && LOT_NUM) {
    // ── TEST 3: Add with lot ───────────────────────────────────────────────────
    Logger.log('');
    Logger.log('=== TEST 3: Add (with lot)  ' + SKU_LOT + '  lot=' + LOT_NUM + '  x' + TEST_QTY + ' ===');
    var r3 = handleWaspQuantityAdded(makePayload(SKU_LOT, TEST_QTY, LOT_NUM, LOT_EXP));
    Logger.log('Result: ' + JSON.stringify(r3));
    Logger.log(r3 && r3.status === 'processed' ? 'PASS — SA with batch created in Katana' : 'CHECK — see Activity tab');

    // ── TEST 4: Remove with lot ────────────────────────────────────────────────
    Logger.log('');
    Logger.log('=== TEST 4: Remove (with lot)  ' + SKU_LOT + '  lot=' + LOT_NUM + '  x' + TEST_QTY + ' ===');
    var r4 = handleWaspQuantityRemoved(makePayload(SKU_LOT, TEST_QTY, LOT_NUM, LOT_EXP));
    Logger.log('Result: ' + JSON.stringify(r4));
    Logger.log(r4 && r4.status === 'processed' ? 'PASS — SA with batch created in Katana' : 'CHECK — see Activity tab');
  } else {
    Logger.log('');
    Logger.log('=== TEST 3 & 4: SKIPPED — no lot-tracked SKU available ===');
  }

  Logger.log('');
  Logger.log('=== DONE ===');
  Logger.log('→ Activity tab: look for "WASP Adjustment" rows labelled "F2 DIAG TEST"');
  Logger.log('→ F2 Adjustments tab: 4 rows should be logged');
  Logger.log('→ Katana SAs: 2-4 "WASP Adjustment" SAs (net qty = 0)');
}
