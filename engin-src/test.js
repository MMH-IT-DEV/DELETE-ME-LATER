function testDebugSheetId() {
  var debugId = PropertiesService.getScriptProperties().getProperty('DEBUG_SHEET_ID');
  Logger.log('DEBUG_SHEET_ID = ' + debugId);
  if (debugId) {
    var ss = SpreadsheetApp.openById(debugId);
    Logger.log('Opened: ' + ss.getName());
  }
}

function diagRemoveBug() {
  var props = PropertiesService.getScriptProperties().getProperties();

  // 1. Check key script properties
  Logger.log('=== SCRIPT PROPERTIES ===');
  Logger.log('GAS_WEBHOOK_URL : ' + (props.GAS_WEBHOOK_URL  ? props.GAS_WEBHOOK_URL.substring(0, 60) + '...' : 'NOT SET'));
  Logger.log('WASP_API_TOKEN  : ' + (props.WASP_API_TOKEN   ? '(set, ' + props.WASP_API_TOKEN.length + ' chars)' : 'NOT SET'));
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
    var token = PropertiesService.getScriptProperties().getProperty('WASP_API_TOKEN');
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
  var token = PropertiesService.getScriptProperties().getProperty('WASP_API_TOKEN');
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
