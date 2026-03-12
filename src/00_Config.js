// ============================================
// 00_Config.gs - CONFIGURATION
// ============================================
// All settings in one place for easy management
// ============================================
// FIXED: Added WASP_SITE_NAME alias for consistency
// FIXED: Changed const to var for Google Apps Script compatibility
// UPDATED: Added site-to-Katana location mapping
// ============================================

var CONFIG = {
  // WASP InventoryCloud
  WASP_BASE_URL: 'https://mymagichealer.waspinventorycloud.com',
  WASP_TOKEN: PropertiesService.getScriptProperties().getProperty('WASP_TOKEN'),
  WASP_SITE: 'MMH Kelowna',
  WASP_SITE_NAME: 'MMH Kelowna',  // FIXED: Added alias for functions that use WASP_SITE_NAME

  // Katana MRP
  KATANA_API_KEY: PropertiesService.getScriptProperties().getProperty('KATANA_API_KEY'),
  KATANA_BASE_URL: 'https://api.katanamrp.com/v1',
  KATANA_WEB_URL: 'https://factory.katanamrp.com',

  // ShipStation
  SHIPSTATION_API_KEY: PropertiesService.getScriptProperties().getProperty('SHIPSTATION_API_KEY'),
  SHIPSTATION_API_SECRET: PropertiesService.getScriptProperties().getProperty('SHIPSTATION_API_SECRET'),
  SHIPSTATION_BASE_URL: 'https://ssapi.shipstation.com',

  // Slack Notifications
  SLACK_WEBHOOK_URL: PropertiesService.getScriptProperties().getProperty('SLACK_WEBHOOK_URL'),

  // Google Sheets
  DEBUG_SHEET_ID: '1eX7MCU-Is5CMmROL1PfuhGoB73yRF7dYdyXHqzMYOUQ',
  SYNC_SHEET_ID: '1FiG8G3J-IbKoCzOiQ4aVCg6N1JBS01w76igmpkECJSI',
  PICK_MAPPINGS_SHEET: 'PickOrderMappings'
};

// ============================================
// WASP SITE TO KATANA LOCATION MAPPING
// ============================================
// Maps WASP site names to Katana location names
// Left side = exact WASP site name (from payload.SiteName)
// Right side = exact Katana location name
// Sites not mapped will be skipped during sync
// ============================================
var SITE_TO_KATANA_LOCATION = {
  'MMH Kelowna': 'MMH Kelowna',
  'MMH Mayfair': 'MMH Mayfair',
  'Storage Warehouse': 'Storage Warehouse'
  // Add your WASP sites mapped to Katana locations:
  // 'WASP Amazon US': 'Amazon USA',
  // 'WASP Amazon CA': 'Amazon CA',
  // 'WASP Amazon AU': 'Amazon AU',
  // 'WASP Shopify': 'Shopify'
};

// ============================================
// KATANA LOCATION → WASP DESTINATION MAP
// ============================================
// Maps Katana location names (from PO "Ship to") to WASP destination
// Used by F1 to route received items to the correct WASP site/location
// ============================================
var KATANA_LOCATION_TO_WASP = {
  'MMH Kelowna': { site: 'MMH Kelowna', location: 'RECEIVING-DOCK' },
  'MMH Mayfair': { site: 'MMH Mayfair', location: 'QA-Hold-1' },
  'Storage Warehouse': { site: 'Storage Warehouse', location: 'SW-STORAGE' }
};

// ============================================
// REVERSIBLE HOTFIX FLAGS
// ============================================
// Toggle these to false for a fast rollback of the March 11, 2026 flow fixes.
var HOTFIX_FLAGS = {
  F1_PARTIAL_NON_BATCH_DELTA: true,
  F1_CONFIRM_NON_BATCH_REVERT: true,
  F1_CONFIRM_FULL_REVERT: true,
  F3_USE_CONFIG_DRIVEN_LOCATION_MAP: true,
  F4_CONFIRM_STATUS_REVERT: true,
  F5_CONFIRM_CANCEL_VIA_UPDATE: true,
  F5_LEDGER_RETRY_GUARD: true,
  F6_REVERT_PRESERVE_EXACT_LOT: true,
  F6_CONFIRM_STATUS_REVERT: true,
  FLOW_COVERAGE_COUNT_PARTIAL_PO: true
};

// F3 uses its own map so transfer routing can be changed without editing logic.
// Keys are Katana location names; values are WASP {site, location}.
var F3_LOCATION_OVERRIDES = {
  'MMH Kelowna': { site: 'MMH Kelowna', location: 'RECEIVING-DOCK' },
  'MMH Mayfair': { site: 'MMH Mayfair', location: 'QA-Hold-1' },
  'Storage Warehouse': { site: 'Storage Warehouse', location: 'SW-STORAGE' }
};

// ============================================
// UOM CONVERSION MAP
// ============================================
// Items where Katana and WASP intentionally use DIFFERENT units.
// factor = multiply Katana qty by this to get WASP qty.
// e.g. EGG-X: Katana tracks in "dozen", WASP in "pcs" → factor=12
//
// For items NOT listed here, Katana and WASP use the same unit (factor=1).
// Update this after running auditItemUoms() to confirm what conversions exist.
// ============================================
var UOM_CONVERSIONS = {
  // No quantity conversions confirmed — all mismatches are naming differences only.
  // See WASP_UOM_MAP below for the name normalisation that IS needed.
};

/**
 * Get quantity conversion factor for a SKU (Katana qty → WASP qty).
 * Returns 1 if no conversion defined (same UOM in both systems).
 */
function getUomConversionFactor(sku) {
  var conv = UOM_CONVERSIONS[sku];
  return conv ? conv.factor : 1;
}

// ============================================
// WASP UOM NAME MAP
// ============================================
// Katana uses abbreviations (EA, PC, g, pcs).
// WASP stores and expects full unit names (Each, PCS, grams).
// Apply mapKatanaUomToWasp_() before writing any UOM to WASP.
//
// Confirmed from audit (checkWaspUnitsForConversions):
//   EA / PC / pc  → Each   (same unit, WASP calls it "Each")
//   pcs / PCS     → PCS
//   g / gram      → grams
//   lbs / lb      → Pound
//   kg            → kg
//   ROLLS         → ROLLS
//   pack          → pack
//   BOX / box     → BOX
//   dozen / dz    → dozen
// ============================================
var WASP_UOM_MAP = {
  'EA':    'Each',  'ea':    'Each',  'Each':  'Each',  'each':  'Each',
  'PC':    'Each',  'pc':    'Each',
  'PCS':   'PCS',   'pcs':   'PCS',   'Pcs':   'PCS',
  'g':     'grams', 'gram':  'grams', 'grams': 'grams', 'Grams': 'grams',
  'kg':    'kg',    'KG':    'kg',
  'lbs':   'Pound', 'lb':    'Pound', 'LBS':   'Pound', 'Pound': 'Pound',
  'oz':    'oz',
  'ROLLS': 'ROLLS', 'rolls': 'ROLLS', 'Roll':  'Roll',  'roll':  'Roll',
  'pack':  'pack',  'Pack':  'pack',  'PK':    'pack',  'pk':    'pack',
  'BOX':   'BOX',   'box':   'BOX',   'Box':   'BOX',
  'dozen': 'dozen', 'dz':    'dozen',
  '6 pack':'6 Pack','6 Pack':'6 Pack',
  'bucket':'bucket'
};

/**
 * Translate a Katana UOM string to the matching WASP unit name.
 * Falls back to the raw value if not in the map (logs a warning).
 */
function mapKatanaUomToWasp_(uom) {
  if (!uom) return 'Each';
  var mapped = WASP_UOM_MAP[uom.trim()];
  if (!mapped) {
    Logger.log('WARN: no WASP_UOM_MAP entry for Katana UOM "' + uom + '" — sending as-is');
    return uom.trim();
  }
  return mapped;
}

// ============================================
// SKIP SKU PREFIXES — Items to exclude from SO processing
// ============================================
// Any SKU starting with these prefixes is filtered from pick orders,
// delivery deductions, and cancellation returns.
// (e.g., OP-195, OP-300 = Order Protection service items)
// ============================================
var SKIP_SKU_PREFIXES = ['OP-'];

/**
 * Check if a SKU should be skipped (service/virtual items)
 */
function isSkippedSku(sku) {
  if (!sku) return false;
  for (var i = 0; i < SKIP_SKU_PREFIXES.length; i++) {
    if (sku.indexOf(SKIP_SKU_PREFIXES[i]) === 0) return true;
  }
  return false;
}

// ============================================
// WASP LOCATIONS (within a site)
// ============================================
var LOCATIONS = {
  RECEIVING: 'RECEIVING-DOCK',    // PO receiving
  SHIPPING: 'SHIPPING-DOCK',      // Order fulfillment picking
  PRODUCTION: 'PRODUCTION',       // MO ingredient consumption
  PROD_RECEIVING: 'PROD-RECEIVING', // MO finished goods output
  SHOPIFY: 'SHOPIFY'              // Shopify orders (legacy)
};

// ============================================
// FLOW CONFIGURATION
// ============================================
var FLOWS = {
  // F1: PO Receiving - where items go when PO received
  PO_RECEIVING_LOCATION: LOCATIONS.RECEIVING,

  // F4: MO Complete - where ingredients come from, where output goes
  MO_INGREDIENT_LOCATION: LOCATIONS.PRODUCTION,
  MO_OUTPUT_LOCATION: LOCATIONS.PROD_RECEIVING,

  // F5: Order Fulfillment - where to pick from (Shopify orders)
  PICK_FROM_LOCATION: LOCATIONS.SHOPIFY,

  // F3: Amazon Transfer (manual - no webhook) — removes from Kelowna staging
  AMAZON_TRANSFER_LOCATION: LOCATIONS.SHIPPING,

  // F3-Amazon: SO delivered from Amazon USA location → remove from Amazon WASP site
  AMAZON_FBA_KATANA_LOCATION: 'Amazon USA',    // Katana location name that triggers this route
  AMAZON_FBA_WASP_SITE: 'Amazon USA',          // WASP site name
  AMAZON_FBA_WASP_LOCATION: 'AMAZON-FBA-USA', // WASP location code

  // F6: Amazon US customer IDs — Katana SO API returns customer_id, not customer_name.
  // ONLY 'Amazon US' (exact name) should trigger F6. Add IDs here if multiple Amazon US
  // accounts exist. Do NOT add Amazon CA, Amazon UK, or other regional variants.
  AMAZON_CUSTOMER_IDS: [43058502]              // Katana customer_id for 'Amazon US'
};

// ============================================
// F2 SKIP LOCATIONS
// ============================================
// WASP locations that F2 (Adjustments) must NEVER process.
// Any quantity_added/removed callout from these locations is silently skipped.
// SHOPIFY: exclusively managed by F5 (ShipStation deductions) —
//   F5 calls waspRemoveInventory → WASP fires callout → F2 must ignore it
//   to prevent double-deduction in Katana.
// ============================================
var F2_SKIP_LOCATIONS = [LOCATIONS.SHOPIFY];

// ============================================
// SYNC CACHE (Prevents feedback loops)
// ============================================
// Keep a wider buffer here: engin-src pre-marks can arrive well before the
// matching WASP callout, and 30s proved too short in practice.
var SYNC_CACHE_SECONDS = 120;

// ============================================
// INVENTORY SYNC CONFIG
// ============================================
var SYNC_CONFIG = {
  ABORT_THRESHOLD: 0.80,
  MAX_EXECUTE_ITEMS: 999,
  RATE_LIMIT_MS: 300,
  DRY_RUN: true,             // Set to false when ready to actually run
  COMPARISON_SHEET_ID: '1mqdZ1Yp9fIzpSMxlJYgfbYGSe1ii2N836L1stc4AtNU',
  SYNC_SHEET_ID: '1FiG8G3J-IbKoCzOiQ4aVCg6N1JBS01w76igmpkECJSI'  // Sync Sheet (Zero Plan / Re-Add Plan tabs)
};

// ============================================
// SYNC LOCATION MAP
// ============================================
// Maps "KatanaLocation|itemType" to WASP site + location
// itemType: product, material, or unknown
// ============================================
var SYNC_LOCATION_MAP = {
  'MMH Kelowna|product': { site: 'MMH Kelowna', location: 'SHOPIFY' },
  'MMH Kelowna|material': { site: 'MMH Kelowna', location: 'PRODUCTION' },
  'MMH Kelowna|intermediate': { site: 'MMH Kelowna', location: 'PRODUCTION' },
  'MMH Kelowna|unknown': { site: 'MMH Kelowna', location: 'PRODUCTION' },
  'MMH Mayfair|product': { site: 'MMH Mayfair', location: 'QA-Hold-1' },
  'MMH Mayfair|material': { site: 'MMH Mayfair', location: 'QA-Hold-1' },
  'MMH Mayfair|intermediate': { site: 'MMH Mayfair', location: 'QA-Hold-1' },
  'MMH Mayfair|unknown': { site: 'MMH Mayfair', location: 'QA-Hold-1' },
  'Storage Warehouse|product': { site: 'Storage Warehouse', location: 'SW-STORAGE' },
  'Storage Warehouse|material': { site: 'Storage Warehouse', location: 'SW-STORAGE' },
  'Storage Warehouse|unknown': { site: 'Storage Warehouse', location: 'SW-STORAGE' }
};

// ============================================
// KATANA WEB URL PATHS (for Activity Log links)
// ============================================
// All paths confirmed from browser (Feb 9, 2026).
// ============================================
var KATANA_WEB_PATHS = {
  so: '/salesorder/',
  mo: '/manufacturingorder/',
  po: '/purchaseorder/',
  st: '/stocktransfer/',
  sa: '/stockadjustment/'
};

/**
 * Build a clickable Katana web URL for the Activity log
 * @param {string} type - 'so', 'mo', 'po', or 'st'
 * @param {string|number} id - Katana internal ID
 * @return {string} Full URL or empty string
 */
function getKatanaWebUrl(type, id) {
  if (!id) return '';
  var path = KATANA_WEB_PATHS[type];
  if (!path) return '';
  return CONFIG.KATANA_WEB_URL + path + id;
}

// ============================================
// HELPER FUNCTIONS
// ============================================

/**
 * Get Katana location name for a WASP site
 * Returns null if site is not mapped (will skip sync)
 */
function getKatanaLocationForSite(waspSiteName) {
  if (!waspSiteName) {
    return SITE_TO_KATANA_LOCATION[CONFIG.WASP_SITE] || null;
  }
  return SITE_TO_KATANA_LOCATION[waspSiteName] || null;
}
