// DIAGNOSTIC — test creating NI-LADDER10 in WASP
function testCreateNILADDER10() {
  var token = getStoredWaspToken_();
  var base = getWaspBase();

  // Step 1: Check if item already exists
  Logger.log('=== STEP 1: Check if item exists ===');
  var searchResp = UrlFetchApp.fetch(base + 'ic/item/infosearch', {
    method: 'POST',
    headers: { 'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json' },
    payload: JSON.stringify({ ItemNumber: 'NI-LADDER10', AltItemNumber: '', ItemDescription: '' }),
    muteHttpExceptions: true
  });
  Logger.log('Search HTTP ' + searchResp.getResponseCode());
  var searchText = searchResp.getContentText();
  var found = false;
  try {
    var searchData = JSON.parse(searchText);
    var items = searchData.Data || [];
    for (var i = 0; i < items.length; i++) {
      if (items[i].ItemNumber === 'NI-LADDER10') {
        Logger.log('Item EXISTS: ' + JSON.stringify(items[i]).substring(0, 500));
        found = true;
        break;
      }
    }
    if (!found) Logger.log('Item NOT in WASP catalog');
  } catch (e) { Logger.log('Search parse error: ' + e.message); }

  // Step 2: Try creating item
  Logger.log('=== STEP 2: Create item ===');
  var createPayload = [{
    ItemNumber: 'NI-LADDER10',
    ItemDescription: 'LADDER',
    StockingUnit: 'EA',
    PurchaseUnit: 'EA',
    SalesUnit: 'EA',
    Cost: 319,
    CategoryDescription: 'SUPPLIES',
    DimensionInfo: { DimensionUnit: 'Inch', Height: 0, Width: 0, Depth: 0, WeightUnit: 'Pound', Weight: 0, VolumeUnit: 'Cubic Inch', MaxVolume: 0 },
    ItemLocationSettings: [{ SiteName: 'MMH Mayfair', LocationCode: 'PRODUCTION', PrimaryLocation: true }],
    ItemSiteSettings: [{ SiteName: 'MMH Mayfair' }]
  }];
  Logger.log('Create payload: ' + JSON.stringify(createPayload).substring(0, 500));
  var createResp = UrlFetchApp.fetch(base + 'ic/item/createInventoryItems', {
    method: 'POST',
    headers: { 'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json' },
    payload: JSON.stringify(createPayload),
    muteHttpExceptions: true
  });
  Logger.log('Create HTTP ' + createResp.getResponseCode());
  Logger.log('Create response: ' + createResp.getContentText().substring(0, 1000));

  // Step 3: Try adding stock
  Logger.log('=== STEP 3: Add stock ===');
  var addPayload = [{ ItemNumber: 'NI-LADDER10', Quantity: 1, SiteName: 'MMH Mayfair', LocationCode: 'PRODUCTION', Notes: 'Test create from diagnostic' }];
  var addResp = UrlFetchApp.fetch(base + 'transactions/item/add', {
    method: 'POST',
    headers: { 'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json' },
    payload: JSON.stringify(addPayload),
    muteHttpExceptions: true
  });
  Logger.log('Add HTTP ' + addResp.getResponseCode());
  Logger.log('Add response: ' + addResp.getContentText().substring(0, 1000));
}

function testDebugSheetId() {
  var debugId = PropertiesService.getScriptProperties().getProperty('DEBUG_SHEET_ID');
  Logger.log('DEBUG_SHEET_ID = ' + debugId);
  if (debugId) {
    var ss = SpreadsheetApp.openById(debugId);
    Logger.log('Opened: ' + ss.getName());
  }
}

function inspectWaspScriptProperties() {
  var props = PropertiesService.getScriptProperties().getProperties();
  var tokenInfo = getWaspTokenInfo_();
  var summary = {
    WASP_BASE_URL_raw: props.WASP_BASE_URL || '',
    WASP_INSTANCE_raw: props.WASP_INSTANCE || '',
    WASP_API_TOKEN_present: !!props.WASP_API_TOKEN,
    WASP_API_TOKEN_length: props.WASP_API_TOKEN ? props.WASP_API_TOKEN.length : 0,
    WASP_TOKEN_present: !!props.WASP_TOKEN,
    WASP_TOKEN_length: props.WASP_TOKEN ? props.WASP_TOKEN.length : 0,
    token_property_used: tokenInfo.propertyName || '',
    token_client_id: tokenInfo.clientId || '',
    token_user: tokenInfo.email || '',
    token_issued: tokenInfo.issued || '',
    token_expires: tokenInfo.expires || '',
    token_roles: tokenInfo.roles || ''
  };
  Logger.log(JSON.stringify(summary, null, 2));
}

function diagRemoveBug() {
  var props = PropertiesService.getScriptProperties().getProperties();

  // 1. Check key script properties
  Logger.log('=== SCRIPT PROPERTIES ===');
  Logger.log('GAS_WEBHOOK_URL : ' + (props.GAS_WEBHOOK_URL  ? props.GAS_WEBHOOK_URL.substring(0, 60) + '...' : 'NOT SET'));
  Logger.log('WASP_API_TOKEN  : ' + (props.WASP_API_TOKEN   ? '(set, ' + props.WASP_API_TOKEN.length + ' chars)' : 'NOT SET'));
  Logger.log('WASP_TOKEN      : ' + (props.WASP_TOKEN       ? '(set, ' + props.WASP_TOKEN.length + ' chars)' : 'NOT SET'));
  Logger.log('WASP_BASE_URL   : ' + (props.WASP_BASE_URL    || 'NOT SET'));
  Logger.log('DEBUG_SHEET_ID  : ' + (props.DEBUG_SHEET_ID   || 'NOT SET'));

  // 2. Test enginPreMark_ directly
  Logger.log('=== TESTING enginPreMark_ ===');
  if (!props.GAS_WEBHOOK_URL) {
    Logger.log('SKIP — GAS_WEBHOOK_URL not set, pre-mark will never fire');
  } else {
    try {
      var markUrl = props.GAS_WEBHOOK_URL;
      var markResp = UrlFetchApp.fetch(markUrl, {
        method: 'POST',
        contentType: 'application/json',
        payload: JSON.stringify({ action: 'engin_mark', sku: 'DIAG-TEST', location: 'PRODUCTION', op: 'remove' }),
        muteHttpExceptions: true
      });
      Logger.log('enginPreMark_ response: ' + markResp.getResponseCode() + ' ' + markResp.getContentText().substring(0, 200));
    } catch (e) {
      Logger.log('enginPreMark_ ERROR: ' + e.message);
    }
  }

  // 3. End-to-end cache test: set pre-mark then simulate quantity_removed callout
  // If cache is shared, the simulated callout should return 'skipped'.
  // If it returns 'processed', the cache is not shared between the two script executions.
  Logger.log('=== END-TO-END CACHE TEST ===');
  if (!props.GAS_WEBHOOK_URL) {
    Logger.log('SKIP — GAS_WEBHOOK_URL not set');
  } else {
    try {
      var webhookUrl = props.GAS_WEBHOOK_URL;
      // Step A: set pre-mark for CACHE-TEST
      var markResp2 = UrlFetchApp.fetch(webhookUrl, {
        method: 'POST', contentType: 'application/json',
        payload: JSON.stringify({ action: 'engin_mark', sku: 'CACHE-TEST', location: 'PRODUCTION', op: 'remove' }),
        muteHttpExceptions: true
      });
      Logger.log('Pre-mark set: ' + markResp2.getContentText().substring(0, 100));

      // Step B: immediately fire simulated quantity_removed for CACHE-TEST
      var calloutResp = UrlFetchApp.fetch(webhookUrl, {
        method: 'POST', contentType: 'application/json',
        payload: JSON.stringify({ action: 'quantity_removed', ItemNumber: 'CACHE-TEST', Quantity: '1', SiteName: 'MMH Kelowna', LocationCode: 'PRODUCTION', Notes: 'DIAG TEST - safe to ignore' }),
        muteHttpExceptions: true
      });
      Logger.log('Simulated callout result: ' + calloutResp.getContentText().substring(0, 200));
      // Expected: {"status":"skipped",...}  — means cache works
      // If:       {"status":"processed",...} — means cache NOT shared, pre-mark is ineffective
    } catch (e) {
      Logger.log('Cache test ERROR: ' + e.message);
    }
  }

  // 4. Call WASP remove for 4OZSEAL directly and log the FULL raw response
  // Adjust qty below if 4OZSEAL currently has different stock.
  // This shows exactly what WASP returns and why HasError:true fires.
  Logger.log('=== WASP REMOVE RAW RESPONSE (4OZSEAL qty=2) ===');
  try {
    var token = getStoredWaspToken_();
    var removeUrl = getWaspBase() + 'transactions/item/remove';
    var removePayload = JSON.stringify([{ ItemNumber: '4OZSEAL', Quantity: 2, SiteName: 'MMH Kelowna', LocationCode: 'PRODUCTION', Notes: 'DIAG TEST - will be re-added manually' }]);
    var removeResp = UrlFetchApp.fetch(removeUrl, {
      method: 'post',
      headers: { 'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json' },
      payload: removePayload,
      muteHttpExceptions: true
    });
    Logger.log('HTTP status: ' + removeResp.getResponseCode());
    Logger.log('Raw body: ' + removeResp.getContentText());
  } catch (e) {
    Logger.log('WASP remove call ERROR: ' + e.message);
  }

  // 5. Exact 4OZSEAL scenario: pre-mark at PRODUCTION then simulate callout at PRODUCTION.
  // This checks whether the pre-mark is actually blocking the real SKU.
  // Also logs the FULL pre-mark response body — if it says {"status":"error"} the
  // pre-mark was rejected (likely because location is empty or sku missing) which
  // means no cache key was ever set.
  Logger.log('=== 5. 4OZSEAL PRE-MARK + CALLOUT TEST ===');
  if (!props.GAS_WEBHOOK_URL) {
    Logger.log('SKIP — GAS_WEBHOOK_URL not set');
  } else {
    try {
      var wh5 = props.GAS_WEBHOOK_URL;

      // Step A: pre-mark exactly as the push engine would for 4OZSEAL
      var mark5 = UrlFetchApp.fetch(wh5, {
        method: 'POST', contentType: 'application/json',
        payload: JSON.stringify({ action: 'engin_mark', sku: '4OZSEAL', location: 'PRODUCTION', op: 'remove' }),
        muteHttpExceptions: true
      });
      Logger.log('Pre-mark body: ' + mark5.getContentText().substring(0, 300));
      // Expect: {"status":"ok","sku":"4OZSEAL","location":"PRODUCTION","op":"remove"}
      // If:     {"status":"error","reason":"missing sku or location"} → location guard rejected it

      // Step B: simulate WASP callout exactly as the real one
      var callout5 = UrlFetchApp.fetch(wh5, {
        method: 'POST', contentType: 'application/json',
        payload: JSON.stringify({ action: 'quantity_removed', ItemNumber: '4OZSEAL', Quantity: '1', SiteName: 'MMH Kelowna', LocationCode: 'PRODUCTION', Notes: 'DIAG TEST - safe to ignore' }),
        muteHttpExceptions: true
      });
      Logger.log('Callout result: ' + callout5.getContentText().substring(0, 300));
      // Expect: {"status":"skipped",...} — pre-mark is blocking ✓
      // If:     {"status":"processed",...} — pre-mark is NOT blocking — SA will be created
    } catch (e) {
      Logger.log('Section 5 ERROR: ' + e.message);
    }
  }

  // 6. Location-mismatch fallback test — checks whether the anyKey fallback is DEPLOYED.
  // Sends pre-mark with location='FAKE-LOC', then callout with location='PRODUCTION'.
  // If they share a cache via the SKU-only key, the callout should be skipped.
  // If it returns "processed", the deployed src code does NOT have the anyKey fallback
  // → you need to re-deploy the src web app in the GAS UI (not just clasp push).
  Logger.log('=== 6. ANYLOCATION FALLBACK DEPLOYMENT CHECK ===');
  if (!props.GAS_WEBHOOK_URL) {
    Logger.log('SKIP — GAS_WEBHOOK_URL not set');
  } else {
    try {
      var wh6 = props.GAS_WEBHOOK_URL;

      // Pre-mark with a location that will NOT match the callout
      var mark6 = UrlFetchApp.fetch(wh6, {
        method: 'POST', contentType: 'application/json',
        payload: JSON.stringify({ action: 'engin_mark', sku: 'FALLBACK-VERIFY', location: 'FAKE-LOC', op: 'remove' }),
        muteHttpExceptions: true
      });
      Logger.log('Fallback pre-mark: ' + mark6.getContentText().substring(0, 200));

      // Callout arrives with a DIFFERENT location
      var callout6 = UrlFetchApp.fetch(wh6, {
        method: 'POST', contentType: 'application/json',
        payload: JSON.stringify({ action: 'quantity_removed', ItemNumber: 'FALLBACK-VERIFY', Quantity: '1', SiteName: 'MMH Kelowna', LocationCode: 'PRODUCTION', Notes: 'DIAG TEST - safe to ignore' }),
        muteHttpExceptions: true
      });
      Logger.log('Fallback callout: ' + callout6.getContentText().substring(0, 300));
      // {"status":"skipped"} → anyKey fallback IS deployed ✓ new code is live
      // {"status":"processed"} → anyKey fallback NOT deployed → re-deploy src web app in GAS UI
    } catch (e) {
      Logger.log('Section 6 ERROR: ' + e.message);
    }
  }
}

/**
 * Run this IMMEDIATELY after a real "Push to WASP" test that creates an SA.
 * Reads the debug properties written by the src webhook during the callout.
 * Shows exactly what SKU/location/cache-keys were checked and what was in cache.
 */
/**
 * Probes candidate WASP API endpoints for catalog item deletion.
 * Run this, then check Execution Log to see which endpoint returns success.
 * Uses NI-MEDGLOVES (W.Qty=0, already targeted for deletion) as a safe test SKU.
 */
function testWaspDeleteEndpoints() {
  var token = getStoredWaspToken_();
  var base = getWaspBase();
  var testSku = 'NI-MEDGLOVES';

  Logger.log('=== WASP DELETE ENDPOINT PROBE (sku=' + testSku + ') ===');

  // Round 3: probe infosearch with different field names to find item in catalog,
  // then check updateInventoryItems with full required fields.
  // Background: all delete-style paths 404. updateInventoryItems returns 200 but
  // fails with -57066 (missing UOM) — meaning the endpoint is valid but needs full payload.
  var candidates = [
    // Try different infosearch parameter names
    { endpoint: 'ic/item/infosearch',            payload: { Search: testSku } },
    { endpoint: 'ic/item/infosearch',            payload: { ItemNo: testSku } },
    { endpoint: 'ic/item/infosearch',            payload: { Number: testSku } },
    // Try advancedinfosearch
    { endpoint: 'ic/item/advancedinfosearch',    payload: { ItemNumber: testSku } },
    { endpoint: 'ic/item/advancedinfosearch',    payload: { Search: testSku } },
    // updateInventoryItems with StockingUnit to fix the UOM error — see if IsActive works
    { endpoint: 'ic/item/updateInventoryItems',  payload: [{ ItemNumber: testSku, StockingUnit: 'PCS', IsActive: false }] },
    { endpoint: 'ic/item/updateInventoryItems',  payload: [{ ItemNumber: testSku, StockingUnit: 'PCS', Active: 0 }] }
  ];

  for (var i = 0; i < candidates.length; i++) {
    var c = candidates[i];
    Logger.log('--- POST ' + c.endpoint + ' | payload: ' + JSON.stringify(c.payload));
    try {
      var resp = UrlFetchApp.fetch(base + c.endpoint, {
        method: 'POST',
        headers: { 'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json' },
        payload: JSON.stringify(c.payload),
        muteHttpExceptions: true
      });
      Logger.log('HTTP ' + resp.getResponseCode() + ': ' + resp.getContentText().substring(0, 500));
    } catch (e) {
      Logger.log('ERROR: ' + e.message);
    }
    Utilities.sleep(800);
  }
  Logger.log('=== Done. Check which endpoint returned HTTP 200 without HasError:true ===');
}

/**
 * TEST: Update a single WASP item description to verify updateInventoryItems works.
 * Uses NI-CUTTER as the test item — expects Katana name "EDGE PROTECTOR CUTTER".
 * Run this first. If SuccessfullResults: 1, proceed to updateAllWaspDescriptions().
 */
function testUpdateOneDescription() {
  var token = getStoredWaspToken_();
  var base = getWaspBase();
  var testSku = 'NI-CUTTER';
  var testDescription = 'EDGE PROTECTOR CUTTER';

  // Step 1: Fetch current item data from WASP to get existing fields
  Logger.log('=== STEP 1: Fetch current item data for ' + testSku + ' ===');
  var searchResp = UrlFetchApp.fetch(base + 'ic/item/infosearch', {
    method: 'POST',
    headers: { 'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json' },
    payload: JSON.stringify({ ItemNumber: testSku, AltItemNumber: '', ItemDescription: '' }),
    muteHttpExceptions: true
  });
  Logger.log('Search HTTP ' + searchResp.getResponseCode());
  var searchBody = searchResp.getContentText();
  var existingItem = null;
  try {
    var searchData = JSON.parse(searchBody);
    var items = searchData.Data || [];
    for (var i = 0; i < items.length; i++) {
      if (items[i].ItemNumber === testSku) {
        existingItem = items[i];
        Logger.log('Found item. Current description: "' + (items[i].ItemDescription || '') + '"');
        Logger.log('StockingUnit: ' + (items[i].StockingUnit || 'N/A'));
        Logger.log('Full item: ' + JSON.stringify(items[i]).substring(0, 800));
        break;
      }
    }
    if (!existingItem) {
      Logger.log('Item ' + testSku + ' NOT found in WASP — cannot update');
      return;
    }
  } catch (e) {
    Logger.log('Search parse error: ' + e.message);
    return;
  }

  // Step 2: Try updateInventoryItems with full required fields
  Logger.log('=== STEP 2: Update description to "' + testDescription + '" ===');
  var uom = existingItem.StockingUnit || 'EA';
  var updatePayload = [{
    ItemNumber: testSku,
    ItemDescription: testDescription,
    StockingUnit: uom,
    PurchaseUnit: uom,
    SalesUnit: uom,
    DimensionInfo: {
      DimensionUnit: 'Inch',
      Height: 0,
      Width: 0,
      Depth: 0,
      WeightUnit: 'Pound',
      Weight: 0,
      VolumeUnit: 'Cubic Inch',
      MaxVolume: 0
    }
  }];
  Logger.log('Update payload: ' + JSON.stringify(updatePayload));
  var updateResp = UrlFetchApp.fetch(base + 'ic/item/updateInventoryItems', {
    method: 'POST',
    headers: { 'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json' },
    payload: JSON.stringify(updatePayload),
    muteHttpExceptions: true
  });
  Logger.log('Update HTTP ' + updateResp.getResponseCode());
  Logger.log('Update response: ' + updateResp.getContentText().substring(0, 1000));

  // Step 3: Verify — re-fetch and check description
  Logger.log('=== STEP 3: Verify description was updated ===');
  Utilities.sleep(1000);
  var verifyResp = UrlFetchApp.fetch(base + 'ic/item/infosearch', {
    method: 'POST',
    headers: { 'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json' },
    payload: JSON.stringify({ ItemNumber: testSku, AltItemNumber: '', ItemDescription: '' }),
    muteHttpExceptions: true
  });
  try {
    var verifyData = JSON.parse(verifyResp.getContentText());
    var vItems = verifyData.Data || [];
    for (var v = 0; v < vItems.length; v++) {
      if (vItems[v].ItemNumber === testSku) {
        Logger.log('After update — description: "' + (vItems[v].ItemDescription || '') + '"');
        if (vItems[v].ItemDescription === testDescription) {
          Logger.log('SUCCESS — description updated correctly!');
        } else {
          Logger.log('MISMATCH — expected "' + testDescription + '", got "' + (vItems[v].ItemDescription || '') + '"');
        }
        break;
      }
    }
  } catch (e) {
    Logger.log('Verify parse error: ' + e.message);
  }
}

/**
 * Batch update ALL WASP item descriptions from Katana names.
 * Run testUpdateOneDescription() first to confirm the API works.
 *
 * Flow:
 * 1. Fetch all Katana items (SKU → Name mapping)
 * 2. Fetch all WASP items via advancedinfosearch (paginated)
 * 3. For each WASP item with empty/mismatched description, call updateInventoryItems
 * 4. Log results
 */
function updateAllWaspDescriptions() {
  var startTime = new Date();
  Logger.log('=== updateAllWaspDescriptions START ===');

  // Step 1: Build SKU → Name map from Katana
  Logger.log('--- Step 1: Fetching Katana data ---');
  var katanaData = fetchAllKatanaData();
  var katanaNameMap = {};
  for (var ki = 0; ki < katanaData.items.length; ki++) {
    var kItem = katanaData.items[ki];
    if (kItem.sku && kItem.name) {
      katanaNameMap[kItem.sku] = kItem.name;
    }
  }
  var katanaSkuCount = Object.keys(katanaNameMap).length;
  Logger.log('Katana: ' + katanaSkuCount + ' unique SKU→Name mappings');

  // Step 2: Fetch all WASP items (paginated) to get current descriptions and UOM
  Logger.log('--- Step 2: Fetching WASP items ---');
  var token = getStoredWaspToken_();
  var base = getWaspBase();
  var allWaspItems = [];
  var page = 1;
  var hasMore = true;
  var totalCount = 0;

  while (hasMore && page <= 20) {
    var searchResp = UrlFetchApp.fetch(base + 'ic/item/advancedinfosearch', {
      method: 'POST',
      headers: { 'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json' },
      payload: JSON.stringify({ SearchPattern: '', PageSize: 100, PageNumber: page }),
      muteHttpExceptions: true
    });

    if (searchResp.getResponseCode() !== 200) {
      Logger.log('Search failed on page ' + page + ': HTTP ' + searchResp.getResponseCode());
      break;
    }

    var body = JSON.parse(searchResp.getContentText());
    var data = body.Data || [];
    if (page === 1) {
      totalCount = body.TotalRecordsLongCount || body.TotalRecordsCount || 0;
      Logger.log('WASP total items: ' + totalCount);
    }

    for (var d = 0; d < data.length; d++) {
      allWaspItems.push(data[d]);
    }

    hasMore = body.HasSuccessWithMoreDataRemaining === true;
    page++;
    Utilities.sleep(300);
  }
  Logger.log('Fetched ' + allWaspItems.length + ' WASP items across ' + (page - 1) + ' pages');

  // Step 3: Find items needing description updates
  Logger.log('--- Step 3: Identifying items to update ---');
  var toUpdate = [];
  var alreadyCorrect = 0;
  var noKatanaMatch = 0;

  for (var wi = 0; wi < allWaspItems.length; wi++) {
    var wItem = allWaspItems[wi];
    var sku = wItem.ItemNumber || '';
    if (!sku) continue;

    var katanaName = katanaNameMap[sku];
    if (!katanaName) {
      noKatanaMatch++;
      continue;
    }

    var currentDesc = wItem.ItemDescription || '';
    if (currentDesc === katanaName) {
      alreadyCorrect++;
      continue;
    }

    toUpdate.push({
      sku: sku,
      currentDesc: currentDesc,
      newDesc: katanaName,
      uom: wItem.StockingUnit || 'EA',
      cost: parseFloat(wItem.Cost) || 0,
      category: wItem.CategoryDescription || '',
      dimensionInfo: wItem.DimensionInfo || null
    });
  }

  Logger.log('Already correct: ' + alreadyCorrect);
  Logger.log('No Katana match: ' + noKatanaMatch);
  Logger.log('Need update: ' + toUpdate.length);

  // Step 4: Update descriptions one at a time with delay
  Logger.log('--- Step 4: Updating ' + toUpdate.length + ' descriptions ---');
  var successCount = 0;
  var failCount = 0;
  var failedItems = [];

  for (var ui = 0; ui < toUpdate.length; ui++) {
    var item = toUpdate[ui];
    // CRITICAL: include ALL fields — WASP resets omitted fields to defaults
    var updatePayload = [{
      ItemNumber: item.sku,
      ItemDescription: item.newDesc,
      StockingUnit: item.uom,
      PurchaseUnit: item.uom,
      SalesUnit: item.uom,
      Cost: item.cost,
      CategoryDescription: item.category,
      DimensionInfo: item.dimensionInfo || {
        DimensionUnit: 'Inch',
        Height: 0,
        Width: 0,
        Depth: 0,
        WeightUnit: 'Pound',
        Weight: 0,
        VolumeUnit: 'Cubic Inch',
        MaxVolume: 0
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
      var respCode = resp.getResponseCode();

      if (respCode === 200) {
        try {
          var parsed = JSON.parse(respBody);
          if (parsed.HasError === false) {
            successCount++;
            if (ui < 5 || ui % 50 === 0) {
              Logger.log('OK [' + (ui + 1) + '/' + toUpdate.length + '] ' + item.sku + ': "' + item.currentDesc + '" -> "' + item.newDesc + '"');
            }
          } else {
            failCount++;
            failedItems.push(item.sku);
            Logger.log('FAIL [' + (ui + 1) + '] ' + item.sku + ': ' + respBody.substring(0, 300));
          }
        } catch (e) {
          failCount++;
          failedItems.push(item.sku);
          Logger.log('PARSE ERROR [' + (ui + 1) + '] ' + item.sku + ': ' + e.message);
        }
      } else {
        failCount++;
        failedItems.push(item.sku);
        Logger.log('HTTP ' + respCode + ' [' + (ui + 1) + '] ' + item.sku + ': ' + respBody.substring(0, 300));
      }
    } catch (e) {
      failCount++;
      failedItems.push(item.sku);
      Logger.log('ERROR [' + (ui + 1) + '] ' + item.sku + ': ' + e.message);
    }

    Utilities.sleep(300);
  }

  var duration = ((new Date()) - startTime) / 1000;
  Logger.log('=== DONE ===');
  Logger.log('Success: ' + successCount + ' | Failed: ' + failCount + ' | Duration: ' + duration.toFixed(1) + 's');
  if (failedItems.length > 0) {
    Logger.log('Failed SKUs: ' + failedItems.join(', '));
  }
}

/**
 * Find WASP catalog items that have NO site/location (orphaned).
 * These exist in the item catalog but not in inventory — invisible to sync.
 * Cross-references with Katana to show which orphans matter.
 */
function findOrphanedWaspItems() {
  var token = getStoredWaspToken_();
  var base = getWaspBase();

  // Step 1: Fetch all WASP catalog items (advancedinfosearch)
  Logger.log('=== Step 1: Fetching WASP catalog items ===');
  var catalogItems = {};
  var page = 1;
  var hasMore = true;
  while (hasMore && page <= 20) {
    var resp = UrlFetchApp.fetch(base + 'ic/item/advancedinfosearch', {
      method: 'POST',
      headers: { 'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json' },
      payload: JSON.stringify({ SearchPattern: '', PageSize: 100, PageNumber: page }),
      muteHttpExceptions: true
    });
    if (resp.getResponseCode() !== 200) break;
    var body = JSON.parse(resp.getContentText());
    var data = body.Data || [];
    for (var d = 0; d < data.length; d++) {
      catalogItems[data[d].ItemNumber] = {
        description: data[d].ItemDescription || '',
        cost: parseFloat(data[d].Cost) || 0,
        uom: data[d].StockingUnit || ''
      };
    }
    hasMore = body.HasSuccessWithMoreDataRemaining === true;
    page++;
    Utilities.sleep(300);
  }
  var catalogCount = Object.keys(catalogItems).length;
  Logger.log('Catalog: ' + catalogCount + ' items');

  // Step 2: Fetch all WASP inventory items (advancedinventorysearch)
  Logger.log('=== Step 2: Fetching WASP inventory items ===');
  var inventorySkus = {};
  var waspData = fetchAllWaspData();
  for (var wi = 0; wi < waspData.items.length; wi++) {
    var sku = waspData.items[wi].itemNumber;
    if (sku) inventorySkus[sku] = true;
  }
  var inventoryCount = Object.keys(inventorySkus).length;
  Logger.log('Inventory: ' + inventoryCount + ' unique SKUs');

  // Step 3: Find orphans (in catalog but NOT in inventory)
  Logger.log('=== Step 3: Finding orphaned items ===');
  var orphans = [];
  var catalogSkus = Object.keys(catalogItems);
  for (var ci = 0; ci < catalogSkus.length; ci++) {
    var cSku = catalogSkus[ci];
    if (!inventorySkus[cSku]) {
      orphans.push(cSku);
    }
  }
  Logger.log('Orphaned (catalog only, no inventory): ' + orphans.length);

  // Step 4: Cross-reference with Katana
  Logger.log('=== Step 4: Cross-referencing with Katana ===');
  var katanaData = fetchAllKatanaData();
  var katanaSkus = {};
  for (var ki = 0; ki < katanaData.items.length; ki++) {
    if (katanaData.items[ki].sku) katanaSkus[katanaData.items[ki].sku] = true;
  }

  var orphansInKatana = [];
  var orphansNotInKatana = [];
  for (var oi = 0; oi < orphans.length; oi++) {
    var oSku = orphans[oi];
    var info = catalogItems[oSku];
    var line = oSku + ' | desc: "' + info.description + '" | cost: ' + info.cost + ' | uom: ' + info.uom;
    if (katanaSkus[oSku]) {
      orphansInKatana.push(line);
    } else {
      orphansNotInKatana.push(line);
    }
  }

  Logger.log('');
  Logger.log('=== ORPHANS ALSO IN KATANA (' + orphansInKatana.length + ') — need site/location ===');
  for (var a = 0; a < orphansInKatana.length; a++) {
    Logger.log('  ' + orphansInKatana[a]);
  }

  Logger.log('');
  Logger.log('=== ORPHANS NOT IN KATANA (' + orphansNotInKatana.length + ') — WASP-only, no action needed ===');
  for (var b = 0; b < orphansNotInKatana.length; b++) {
    Logger.log('  ' + orphansNotInKatana[b]);
  }

  Logger.log('');
  Logger.log('=== SUMMARY ===');
  Logger.log('Catalog items: ' + catalogCount);
  Logger.log('Inventory items: ' + inventoryCount);
  Logger.log('Orphaned total: ' + orphans.length);
  Logger.log('  In Katana (need fix): ' + orphansInKatana.length);
  Logger.log('  WASP-only (ignore): ' + orphansNotInKatana.length);
}

/**
 * TEST: Update cost, UOM, and category on a single WASP item (NI-CUTTER).
 * Run this first to verify updateInventoryItems works for these fields.
 */
function testUpdateOneCostUomCategory() {
  var token = getStoredWaspToken_();
  var base = getWaspBase();
  var testSku = 'NI-CUTTER';

  // Step 1: Get current WASP item
  Logger.log('=== STEP 1: Fetch current WASP data for ' + testSku + ' ===');
  var searchResp = UrlFetchApp.fetch(base + 'ic/item/infosearch', {
    method: 'POST',
    headers: { 'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json' },
    payload: JSON.stringify({ ItemNumber: testSku, AltItemNumber: '', ItemDescription: '' }),
    muteHttpExceptions: true
  });
  var existingItem = null;
  try {
    var searchData = JSON.parse(searchResp.getContentText());
    var items = searchData.Data || [];
    for (var i = 0; i < items.length; i++) {
      if (items[i].ItemNumber === testSku) {
        existingItem = items[i];
        Logger.log('Current Cost: ' + (items[i].Cost || 0));
        Logger.log('Current StockingUnit: ' + (items[i].StockingUnit || ''));
        Logger.log('Current Category: ' + (items[i].CategoryDescription || ''));
        break;
      }
    }
    if (!existingItem) { Logger.log('Item not found'); return; }
  } catch (e) { Logger.log('Parse error: ' + e.message); return; }

  // Step 2: Get Katana data for this SKU
  Logger.log('=== STEP 2: Fetch Katana data ===');
  var katanaData = fetchAllKatanaData();
  var katanaCost = 0;
  var katanaUom = '';
  var katanaCategory = '';
  for (var k = 0; k < katanaData.items.length; k++) {
    if (katanaData.items[k].sku === testSku) {
      var kCost = parseFloat(katanaData.items[k].avgCost) || 0;
      if (kCost > 0 && kCost > katanaCost) katanaCost = kCost;
      katanaUom = katanaData.items[k].uom || katanaUom;
      katanaCategory = katanaData.items[k].category || katanaCategory;
    }
  }
  Logger.log('Katana avgCost: ' + katanaCost);
  Logger.log('Katana UOM: ' + katanaUom);
  Logger.log('Katana Category: ' + katanaCategory);

  // Step 3: Build update payload (NO ItemDescription — avoids wiping it)
  var waspUom = normalizeUomForWasp(katanaUom);
  var updatePayload = [{
    ItemNumber: testSku,
    StockingUnit: waspUom,
    PurchaseUnit: waspUom,
    SalesUnit: waspUom,
    DimensionInfo: existingItem.DimensionInfo || {
      DimensionUnit: 'Inch', Height: 0, Width: 0, Depth: 0,
      WeightUnit: 'Pound', Weight: 0, VolumeUnit: 'Cubic Inch', MaxVolume: 0
    }
  }];
  if (katanaCost > 0) {
    updatePayload[0].Cost = katanaCost;
  }
  if (katanaCategory) {
    updatePayload[0].CategoryDescription = katanaCategory;
  }

  Logger.log('=== STEP 3: Update payload ===');
  Logger.log(JSON.stringify(updatePayload));

  var resp = UrlFetchApp.fetch(base + 'ic/item/updateInventoryItems', {
    method: 'POST',
    headers: { 'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json' },
    payload: JSON.stringify(updatePayload),
    muteHttpExceptions: true
  });
  Logger.log('Update HTTP ' + resp.getResponseCode());
  Logger.log('Update response: ' + resp.getContentText().substring(0, 1000));

  // Step 4: Verify
  Logger.log('=== STEP 4: Verify ===');
  Utilities.sleep(1000);
  var verifyResp = UrlFetchApp.fetch(base + 'ic/item/infosearch', {
    method: 'POST',
    headers: { 'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json' },
    payload: JSON.stringify({ ItemNumber: testSku, AltItemNumber: '', ItemDescription: '' }),
    muteHttpExceptions: true
  });
  try {
    var vData = JSON.parse(verifyResp.getContentText());
    var vItems = vData.Data || [];
    for (var v = 0; v < vItems.length; v++) {
      if (vItems[v].ItemNumber === testSku) {
        Logger.log('After — Cost: ' + (vItems[v].Cost || 0) + ' (wanted: ' + katanaCost + ')');
        Logger.log('After — StockingUnit: ' + (vItems[v].StockingUnit || '') + ' (wanted: ' + waspUom + ')');
        Logger.log('After — Category: ' + (vItems[v].CategoryDescription || '') + ' (wanted: ' + katanaCategory + ')');
        break;
      }
    }
  } catch (e) { Logger.log('Verify error: ' + e.message); }
}

/**
 * Batch sync Cost, UOM, and Category from Katana to all WASP items.
 *
 * Cost logic:
 *   - Use Katana average_cost (highest non-zero across locations per SKU)
 *   - If Katana cost is 0 (stock depleted), keep current WASP cost unchanged
 *
 * UOM logic:
 *   - Normalize Katana UOM to WASP format, update if different
 *
 * Category logic:
 *   - Set from Katana category, skip if WASP returns -57006 (category not in WASP)
 *
 * Run testUpdateOneCostUomCategory() first to verify the API accepts these fields.
 */
function syncAllWaspCostUomCategory() {
  var startTime = new Date();
  Logger.log('=== syncAllWaspCostUomCategory START ===');

  // Step 1: Build Katana SKU data: cost, uom, category
  Logger.log('--- Step 1: Fetching Katana data ---');
  var katanaData = fetchAllKatanaData();
  var katanaSkuMap = {};  // sku -> { cost, uom, category }

  for (var ki = 0; ki < katanaData.items.length; ki++) {
    var kItem = katanaData.items[ki];
    if (!kItem.sku) continue;
    var kCost = parseFloat(kItem.avgCost) || 0;

    if (!katanaSkuMap[kItem.sku]) {
      katanaSkuMap[kItem.sku] = {
        cost: kCost,
        uom: kItem.uom || 'ea',
        category: kItem.category || ''
      };
    } else {
      // Keep highest non-zero cost across locations
      if (kCost > 0 && kCost > katanaSkuMap[kItem.sku].cost) {
        katanaSkuMap[kItem.sku].cost = kCost;
      }
      // Fill in uom/category if not set yet
      if (!katanaSkuMap[kItem.sku].uom || katanaSkuMap[kItem.sku].uom === 'ea') {
        katanaSkuMap[kItem.sku].uom = kItem.uom || katanaSkuMap[kItem.sku].uom;
      }
      if (!katanaSkuMap[kItem.sku].category) {
        katanaSkuMap[kItem.sku].category = kItem.category || '';
      }
    }
  }
  var katanaSkuCount = Object.keys(katanaSkuMap).length;
  Logger.log('Katana: ' + katanaSkuCount + ' SKUs with cost/uom/category data');

  // Step 2: Fetch all WASP items via infosearch (has Cost + Category fields)
  Logger.log('--- Step 2: Fetching WASP items ---');
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

    if (searchResp.getResponseCode() !== 200) {
      Logger.log('Search failed page ' + page + ': HTTP ' + searchResp.getResponseCode());
      break;
    }

    var body = JSON.parse(searchResp.getContentText());
    var data = body.Data || [];
    if (page === 1) {
      Logger.log('WASP total items: ' + (body.TotalRecordsLongCount || body.TotalRecordsCount || 0));
      // Log first item fields to confirm Cost/Category available
      if (data.length > 0) {
        Logger.log('WASP item fields: ' + Object.keys(data[0]).join(', '));
      }
    }

    for (var d = 0; d < data.length; d++) {
      allWaspItems.push(data[d]);
    }

    hasMore = body.HasSuccessWithMoreDataRemaining === true;
    page++;
    Utilities.sleep(300);
  }
  Logger.log('Fetched ' + allWaspItems.length + ' WASP items');

  // Step 3: Compare and build update list
  Logger.log('--- Step 3: Identifying items to update ---');
  var toUpdate = [];
  var skippedNoCost = 0;
  var noKatanaMatch = 0;
  var alreadyCorrect = 0;

  for (var wi = 0; wi < allWaspItems.length; wi++) {
    var wItem = allWaspItems[wi];
    var sku = wItem.ItemNumber || '';
    if (!sku) continue;

    var kData = katanaSkuMap[sku];
    if (!kData) { noKatanaMatch++; continue; }

    var waspCost = parseFloat(wItem.Cost) || 0;
    var waspUom = wItem.StockingUnit || 'Each';
    var waspCategory = wItem.CategoryDescription || '';

    var newUom = normalizeUomForWasp(kData.uom);
    var newCost = kData.cost;
    var newCategory = kData.category || '';

    // Determine what needs updating
    var costNeedsUpdate = false;
    var uomNeedsUpdate = false;
    var categoryNeedsUpdate = false;

    // Cost: update only if Katana has non-zero cost AND it differs
    if (newCost > 0 && Math.abs(newCost - waspCost) > 0.001) {
      costNeedsUpdate = true;
    } else if (newCost === 0) {
      skippedNoCost++;
    }

    // UOM: update if different
    if (newUom !== waspUom) {
      uomNeedsUpdate = true;
    }

    // Category: update if Katana has one and WASP is empty or different
    if (newCategory && newCategory !== waspCategory) {
      categoryNeedsUpdate = true;
    }

    if (!costNeedsUpdate && !uomNeedsUpdate && !categoryNeedsUpdate) {
      alreadyCorrect++;
      continue;
    }

    toUpdate.push({
      sku: sku,
      waspCost: waspCost,
      newCost: costNeedsUpdate ? newCost : null,
      waspUom: waspUom,
      newUom: uomNeedsUpdate ? newUom : waspUom,
      waspCategory: waspCategory,
      newCategory: categoryNeedsUpdate ? newCategory : null,
      costChange: costNeedsUpdate,
      uomChange: uomNeedsUpdate,
      categoryChange: categoryNeedsUpdate,
      description: wItem.ItemDescription || '',
      currentCost: waspCost,
      currentCategory: waspCategory,
      dimensionInfo: wItem.DimensionInfo || null
    });
  }

  Logger.log('Already correct: ' + alreadyCorrect);
  Logger.log('No Katana match: ' + noKatanaMatch);
  Logger.log('Skipped (Katana cost=0, keeping WASP cost): ' + skippedNoCost);
  Logger.log('Need update: ' + toUpdate.length);

  // Log breakdown
  var costUpdates = 0;
  var uomUpdates = 0;
  var catUpdates = 0;
  for (var ci = 0; ci < toUpdate.length; ci++) {
    if (toUpdate[ci].costChange) costUpdates++;
    if (toUpdate[ci].uomChange) uomUpdates++;
    if (toUpdate[ci].categoryChange) catUpdates++;
  }
  Logger.log('  Cost changes: ' + costUpdates);
  Logger.log('  UOM changes: ' + uomUpdates);
  Logger.log('  Category changes: ' + catUpdates);

  // Step 4: Update each item
  Logger.log('--- Step 4: Updating ' + toUpdate.length + ' items ---');
  var successCount = 0;
  var failCount = 0;
  var catFailCount = 0;
  var failedItems = [];

  for (var ui = 0; ui < toUpdate.length; ui++) {
    var item = toUpdate[ui];
    var uomToSend = item.newUom;
    // CRITICAL: include ALL fields — WASP resets omitted fields to defaults
    var updatePayload = [{
      ItemNumber: item.sku,
      ItemDescription: item.description,
      StockingUnit: uomToSend,
      PurchaseUnit: uomToSend,
      SalesUnit: uomToSend,
      Cost: item.newCost !== null ? item.newCost : item.currentCost,
      CategoryDescription: item.newCategory !== null ? item.newCategory : item.currentCategory,
      DimensionInfo: item.dimensionInfo || {
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
      var respCode = resp.getResponseCode();

      if (respCode === 200) {
        var parsed = JSON.parse(respBody);
        if (parsed.HasError === false) {
          successCount++;
          if (ui < 5 || ui % 50 === 0) {
            var changes = [];
            if (item.costChange) changes.push('cost:' + item.waspCost + '->' + item.newCost);
            if (item.uomChange) changes.push('uom:' + item.waspUom + '->' + item.newUom);
            if (item.categoryChange) changes.push('cat:"' + item.waspCategory + '"->"' + item.newCategory + '"');
            Logger.log('OK [' + (ui + 1) + '/' + toUpdate.length + '] ' + item.sku + ': ' + changes.join(', '));
          }
        } else {
          // Retry logic for known recoverable errors
          var retried = false;

          // UOM not found (-57072) — fall back to current WASP UOM
          if (respBody.indexOf('-57072') > -1) {
            updatePayload[0].StockingUnit = item.waspUom;
            updatePayload[0].PurchaseUnit = item.waspUom;
            updatePayload[0].SalesUnit = item.waspUom;
            retried = true;
          }

          // Category not found (-57006) — fall back to current WASP category
          if (respBody.indexOf('-57006') > -1) {
            catFailCount++;
            updatePayload[0].CategoryDescription = item.currentCategory;
            retried = true;
          }

          if (retried) {
            var retryResp = UrlFetchApp.fetch(base + 'ic/item/updateInventoryItems', {
              method: 'POST',
              headers: { 'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json' },
              payload: JSON.stringify(updatePayload),
              muteHttpExceptions: true
            });
            var retryBody = retryResp.getContentText();
            var retryParsed = JSON.parse(retryBody);
            if (retryParsed.HasError === false) {
              successCount++;
              var skipNote = [];
              if (respBody.indexOf('-57072') > -1) skipNote.push('UOM "' + item.newUom + '" not in WASP');
              if (respBody.indexOf('-57006') > -1) skipNote.push('category "' + item.newCategory + '" not in WASP');
              Logger.log('OK [' + (ui + 1) + '] ' + item.sku + ': updated (' + skipNote.join(', ') + ' — kept current)');
            } else {
              failCount++;
              failedItems.push(item.sku);
              Logger.log('FAIL [' + (ui + 1) + '] ' + item.sku + ' (retry): ' + retryBody.substring(0, 300));
            }
          } else {
            failCount++;
            failedItems.push(item.sku);
            Logger.log('FAIL [' + (ui + 1) + '] ' + item.sku + ': ' + respBody.substring(0, 300));
          }
        }
      } else {
        failCount++;
        failedItems.push(item.sku);
        Logger.log('HTTP ' + respCode + ' [' + (ui + 1) + '] ' + item.sku + ': ' + respBody.substring(0, 300));
      }
    } catch (e) {
      failCount++;
      failedItems.push(item.sku);
      Logger.log('ERROR [' + (ui + 1) + '] ' + item.sku + ': ' + e.message);
    }

    Utilities.sleep(300);
  }

  var duration = ((new Date()) - startTime) / 1000;
  Logger.log('=== DONE ===');
  Logger.log('Success: ' + successCount + ' | Failed: ' + failCount + ' | Category skipped: ' + catFailCount + ' | Duration: ' + duration.toFixed(1) + 's');
  if (failedItems.length > 0) {
    Logger.log('Failed SKUs: ' + failedItems.join(', '));
  }
}

function diagReadCalloutDebug() {
  Logger.log('=== READING SRC WEBHOOK DEBUG PROPERTIES ===');
  var webhookUrl = PropertiesService.getScriptProperties().getProperty('GAS_WEBHOOK_URL');
  if (!webhookUrl) {
    Logger.log('No GAS_WEBHOOK_URL — cannot call src webhook');
    return;
  }
  var resp = UrlFetchApp.fetch(webhookUrl, {
    method: 'POST', contentType: 'application/json',
    payload: JSON.stringify({ action: 'read_debug' }),
    muteHttpExceptions: true
  });
  var raw = resp.getContentText();
  Logger.log('Raw response: ' + raw.substring(0, 1000));
  try {
    var parsed = JSON.parse(raw);
    Logger.log('lastPremark : ' + parsed.lastPremark);
    Logger.log('lastCallout : ' + parsed.lastCallout);
  } catch(e) {
    Logger.log('Parse error: ' + e.message);
  }
}
