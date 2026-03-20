// ============================================
// 20_HealthCheckHandler.gs — F7 Health Check Activity Logger
// ============================================
// Receives summary from sync script's daily health check.
// Logs to Activity tab as F7, sends Slack alert on mismatches.
// ============================================

/**
 * Handle health_check_report POST from sync script.
 * @param {Object} payload - Summary with mismatches counts
 * @returns {Object} Processing result
 */
function handleHealthCheckReport(payload) {
  var mismatches = payload.mismatches || {};
  var totalMismatches = parseInt(payload.totalMismatches || 0) || 0;
  var fixed = parseInt(payload.fixed || 0) || 0;
  var katanaCount = parseInt(payload.katanaCount || 0) || 0;
  var waspCount = parseInt(payload.waspCount || 0) || 0;
  var timestamp = payload.timestamp || '';
  var skusWithMismatches = payload.skusWithMismatches || [];

  // Build sub-items — one per dimension with mismatches
  var subItems = [];
  var dims = ['lot', 'qty', 'uom', 'category', 'cost'];
  var dimLabels = {
    lot: 'lot tracking mismatch',
    qty: 'qty mismatch',
    uom: 'UOM mismatch',
    category: 'category mismatch',
    cost: 'cost mismatch (>5% delta)'
  };

  for (var d = 0; d < dims.length; d++) {
    var dim = dims[d];
    var count = parseInt(mismatches[dim] || 0) || 0;
    if (count > 0) {
      subItems.push({
        sku: dim,
        qty: count,
        action: dimLabels[dim] || dim,
        status: 'Mismatch',
        success: false,
        qtyColor: 'red'
      });
    }
  }

  // If no mismatches, add a single "all clear" sub-item
  if (subItems.length === 0) {
    subItems.push({
      sku: 'all',
      qty: katanaCount,
      action: 'items compared — no mismatches',
      status: 'Checked',
      success: true,
      qtyColor: 'green'
    });
  }

  // Determine status
  var status = totalMismatches === 0 ? 'success' : 'failed';

  // Build header details
  var headerDetails = katanaCount + ' items compared';
  if (totalMismatches > 0) {
    headerDetails += ' | ' + totalMismatches + ' mismatches';
  }

  var context = 'K=' + katanaCount + ' W=' + waspCount;
  if (fixed > 0) context += ' | fixed=' + fixed;

  var headerError = totalMismatches > 0 ? totalMismatches + ' mismatches' : '';

  // Log to Activity tab
  logActivity('F7', headerDetails, status, context, subItems, null, null, headerError);

  // Send Slack notification only on mismatches
  if (totalMismatches > 0) {
    var skuList = skusWithMismatches.slice(0, 5).join(', ');
    if (skusWithMismatches.length > 5) {
      skuList += ' (+' + (skusWithMismatches.length - 5) + ' more)';
    }

    var slackParts = [];
    for (var s = 0; s < dims.length; s++) {
      var sCount = parseInt(mismatches[dims[s]] || 0) || 0;
      if (sCount > 0) {
        slackParts.push(dims[s].charAt(0).toUpperCase() + dims[s].slice(1) + ': ' + sCount);
      }
    }

    var slackMsg = '[Health Check] ' + (timestamp || new Date().toISOString().substring(0, 16)) +
      ' — ' + totalMismatches + ' mismatches\n' +
      slackParts.join(' | ') + '\n' +
      'SKUs: ' + skuList;

    sendSlackNotification(slackMsg);
  }

  return {
    status: 'ok',
    logged: true,
    totalMismatches: totalMismatches
  };
}
