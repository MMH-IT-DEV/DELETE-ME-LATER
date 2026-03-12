// ============================================
// 01_Router.gs — doPost entry point + webhook/callout routing
// ============================================
// ALL incoming HTTP POSTs arrive here.
//
// Sources:
//   Katana webhooks   → payload.action  (e.g. 'sales_order.delivered')
//   WASP callouts     → payload.action  (e.g. 'quantity_removed')
//
// Both use the "action" key — no conflict because Katana values always
// contain a dot (purchase_order.received) while WASP values do not.
//
// WASP callout JSON template must include:
//   "action": "quantity_removed"   (or quantity_added / item_moved)
//   "ItemNumber": "{trans.AssetTag}"
// ============================================

/**
 * GAS Web App entry point — receives all inbound POSTs.
 * Parses JSON body, merges URL ?event= param, then routes.
 */
function doPost(e) {
  var body = (e && e.postData && e.postData.contents) ? e.postData.contents : '';
  var payload = {};

  try {
    if (body) {
      payload = JSON.parse(body);
    }
  } catch (parseErr) {
    Logger.log('doPost JSON parse error: ' + parseErr.message + ' | body[0..200]: ' + String(body).substring(0, 200));
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: 'invalid JSON' }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // Allow URL query param to set/override event type
  // e.g. .../exec?event=quantity_removed
  if (e && e.parameter && e.parameter.event) {
    payload.event = e.parameter.event;
  }

  // ── Log receipt BEFORE processing — captures arrival even if processing crashes ──
  var wqRow = -1;
  try { wqRow = logWebhookReceipt(payload); } catch (wqErr) { Logger.log('Webhook receipt error: ' + wqErr.message); }

  var result;
  try {
    result = routeWebhook(payload);
  } catch (routeErr) {
    Logger.log('doPost routing error: ' + routeErr.message + '\n' + routeErr.stack);
    result = { status: 'error', message: routeErr.message };
  }

  // ── Update receipt row with final result (or fallback to full append) ──
  try {
    if (wqRow > 0) {
      updateWebhookQueueRow(wqRow, result);
    } else {
      logWebhookQueue(payload, result);
    }
  } catch (wqErr) { Logger.log('Webhook Queue update error: ' + wqErr.message); }

  // Heartbeat — report to Health Monitor on every processed webhook
  try {
    var hbStatus  = (result && result.status === 'error') ? 'error' : 'success';
    var hbDetails = (result && result.status === 'error')
      ? (result.message || 'Routing error')
      : ('action=' + (payload.action || payload.event || 'unknown'));
    sendHeartbeatToMonitor_('2026-Katana-WASP_DebugLog', hbStatus, hbDetails);
  } catch (hbErr) { Logger.log('Heartbeat error: ' + hbErr.message); }

  return ContentService
    .createTextOutput(JSON.stringify(result || { status: 'ok' }))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Route incoming payload to the correct handler.
 * Called by doPost() and also by backfillMissedSOs() in 97_WebhookAudit.gs.
 *
 * @param {Object} payload - Parsed JSON payload from the POST body.
 * @return {Object} Handler result (always an object with at least {status}).
 */
function routeWebhook(payload) {
  if (!payload) {
    Logger.log('routeWebhook: null payload');
    return { status: 'error', message: 'null payload' };
  }

  // WASP uses "action" field (e.g. "quantity_added"), same key as Katana.
  // Katana action values always contain a dot (e.g. "sales_order.delivered"),
  // so WASP-specific values are unambiguous and are checked first.
  var action = payload.action || payload.event || '';

  // ── Engin sheet pre-mark (suppress F2 echo for engin-sheet WASP calls) ──
  if (action === 'engin_mark') {
    return handleEnginMark(payload);
  }

  // ── Debug: read back properties written during last callout ──
  if (action === 'read_debug') {
    var dp = PropertiesService.getScriptProperties();
    return {
      status: 'ok',
      lastPremark:  dp.getProperty('DEBUG_LAST_PREMARK')  || 'none',
      lastCallout:  dp.getProperty('DEBUG_LAST_CALLOUT')  || 'none'
    };
  }

  // ── WASP callout routing ──────────────────────────────────────────────────
  if (action === 'quantity_added') {
    return handleWaspQuantityAdded(payload);
  }
  if (action === 'quantity_removed') {
    return handleWaspQuantityRemoved(payload);
  }
  if (action === 'item_moved' || action === 'quantity_moved') {
    return handleWaspItemMoved(payload);
  }

  // ── ShipStation SHIP_NOTIFY webhook ──────────────────────────────────────
  if (payload.resource_type === 'SHIP_NOTIFY' && payload.resource_url) {
    return handleShipNotify(payload.resource_url);
  }

  // ── Katana webhook routing ────────────────────────────────────────────────
  if (action === 'purchase_order.created') {
    return handlePurchaseOrderCreated(payload);
  }
  if (action === 'purchase_order.received' || action === 'purchase_order.partially_received') {
    return handlePurchaseOrderReceived(payload);
  }
  if (action === 'purchase_order.updated') {
    return handlePurchaseOrderUpdated(payload);
  }
  if (action === 'sales_order.created') {
    // F5 is now ShipStation-only. Katana SO creation no longer triggers WASP pick orders.
    return { status: 'ignored', action: 'sales_order.created' };
  }
  if (action === 'sales_order.delivered') {
    return handleSalesOrderDelivered(payload);
  }
  if (action === 'sales_order.cancelled') {
    return handleSalesOrderCancelled(payload);
  }
  if (action === 'sales_order.updated') {
    return handleSalesOrderUpdated(payload);
  }
  if (action === 'manufacturing_order.done' || action === 'manufacturing_order.completed') {
    return handleManufacturingOrderDone(payload);
  }
  if (action === 'manufacturing_order.updated') {
    return handleManufacturingOrderUpdated(payload);
  }
  if (action === 'manufacturing_order.deleted') {
    return handleManufacturingOrderDeleted(payload);
  }
  if (action === 'stock_adjustment.deleted') {
    return handleSADeleted(payload);
  }

  Logger.log('routeWebhook: unhandled payload — action="' + action + '" keys=' + Object.keys(payload).join(','));
  return { status: 'ignored', action: action };
}
