# Current Configuration Snapshot — 2026-03-19

All location mappings confirmed working and deployed at **@410**.

---

## Master Location Rule

**ALL sites route to PRODUCTION** (except Storage Warehouse → SW-STORAGE).
This applies across ALL flows (F1, F2, F3, F4, F5) and the sync script.

---

## Debug Script — `debug/script/00_Config.js`

### KATANA_LOCATION_TO_WASP (used by F1 PO Receiving)
```javascript
var KATANA_LOCATION_TO_WASP = {
  'MMH Kelowna':       { site: 'MMH Kelowna',       location: 'PRODUCTION' },
  'MMH Mayfair':       { site: 'MMH Mayfair',        location: 'PRODUCTION' },
  'Storage Warehouse': { site: 'Storage Warehouse',  location: 'SW-STORAGE' }
};
```

### F3_LOCATION_OVERRIDES (used by F3 Stock Transfers — overrides KATANA_LOCATION_TO_WASP)
```javascript
var F3_LOCATION_OVERRIDES = {
  'MMH Kelowna':       { site: 'MMH Kelowna',       location: 'PRODUCTION' },
  'MMH Mayfair':       { site: 'MMH Mayfair',        location: 'PRODUCTION' },
  'Storage Warehouse': { site: 'Storage Warehouse',  location: 'SW-STORAGE' }
};
```

### HOTFIX_FLAGS
```javascript
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
```

### F2 Suppression Config
```javascript
var F2_SKIP_LOCATIONS = [LOCATIONS.SHOPIFY];
var SYNC_CACHE_SECONDS = 120;
var BATCH_WINDOW_MS = 10000;  // 10-second batch window for WASP callouts
```

### F3 Polling Config
```javascript
var F3_CONFIG = {
  POLL_INTERVAL_MINUTES: 1,
  SYNC_ON_STATUS: ['completed', 'done', 'received', 'partial', 'in_transit'],
  MIN_CREATED_DATE: '2026-03-01T00:00:00Z',
  MAX_AGE_DAYS: 30
};
var F3_REVERSE_STATUS = ['cancelled', 'voided', 'draft'];
var F3_SKIP_LOCATIONS = ['amazon usa', 'shopify'];
```

### F3 Config Source Priority
```javascript
// KATANA_LOCATION_TO_WASP loads first, F3_LOCATION_OVERRIDES overwrites (last wins)
var configSources = [KATANA_LOCATION_TO_WASP, F3_LOCATION_OVERRIDES];
```

---

## Debug Script — `debug/script/07_F3_AutoPoll.js`

### F3_SITE_MAP (fallback, matches F3_LOCATION_OVERRIDES)
```javascript
var F3_SITE_MAP = {
  'mmh kelowna':       { site: 'MMH Kelowna',       location: 'PRODUCTION' },
  'mmh mayfair':       { site: 'MMH Mayfair',        location: 'PRODUCTION' },
  'storage warehouse': { site: 'Storage Warehouse',  location: 'SW-STORAGE' }
};
```

---

## Sync Script — `sync/script/04_SyncEngine.js`

### SYNC_LOCATION_MAP (used by sync tab writes + push routing)
```javascript
var SYNC_LOCATION_MAP = {
  'MMH Kelowna|Product':      { site: 'MMH Kelowna',       location: 'PRODUCTION' },
  'MMH Kelowna|Material':     { site: 'MMH Kelowna',       location: 'PRODUCTION' },
  'MMH Kelowna|Intermediate': { site: 'MMH Kelowna',       location: 'PRODUCTION' },
  'MMH Kelowna|':             { site: 'MMH Kelowna',       location: 'PRODUCTION' },
  'MMH Mayfair|Product':      { site: 'MMH Mayfair',        location: 'PRODUCTION' },
  'MMH Mayfair|Material':     { site: 'MMH Mayfair',        location: 'PRODUCTION' },
  'MMH Mayfair|Intermediate': { site: 'MMH Mayfair',        location: 'PRODUCTION' },
  'MMH Mayfair|':             { site: 'MMH Mayfair',        location: 'PRODUCTION' },
  'Storage Warehouse|Product':      { site: 'Storage Warehouse', location: 'SW-STORAGE' },
  'Storage Warehouse|Material':     { site: 'Storage Warehouse', location: 'SW-STORAGE' },
  'Storage Warehouse|Intermediate': { site: 'Storage Warehouse', location: 'SW-STORAGE' },
  'Storage Warehouse|':             { site: 'Storage Warehouse', location: 'SW-STORAGE' }
};
```

---

## Sync Script — `sync/script/04b_RawTabs.js`

### getKatanaSyncTarget_ (overrides getKatanaPushTarget_)
```javascript
function getKatanaSyncTarget_(katanaLoc, itemType) {
  // 1. Checks SYNC_LOCATION_MAP first (type-specific key, then fallback)
  // 2. Hardcoded fallbacks:
  //    Storage Warehouse → SW-STORAGE
  //    MMH Mayfair       → PRODUCTION
  //    MMH Kelowna       → PRODUCTION
}
```

---

## Sync Script — `sync/script/05_PushEngine.js`

### getKatanaPushTarget_ (delegates to getKatanaSyncTarget_ if available)
```javascript
function getKatanaPushTarget_(katanaLoc, itemType) {
  // Delegates to getKatanaSyncTarget_ when available
  // Hardcoded fallbacks (same as above):
  //    Storage Warehouse → SW-STORAGE
  //    MMH Mayfair       → PRODUCTION
  //    MMH Kelowna       → PRODUCTION
}
```

---

## Flow Labels and Colors — `debug/script/08_Logging.js`

```javascript
var FLOW_LABELS = {
  'F1': 'F1 Receiving',
  'F2': 'F2 Adjustments',
  'F3': 'F3 Transfers',
  'F4': 'F4 Manufacturing',
  'F5': 'F5 Shipping',
  'F6': 'F6 Amazon FBA',
  'F7': 'F7 Health Check'
};

var FLOW_COLORS = {
  'F1': '#cce5ff',   // blue
  'F2': '#b2dfdb',   // teal
  'F3': '#fce4d6',   // peach
  'F4': '#f0e6f6',   // purple
  'F5': '#d1ecf1',   // cyan
  'F6': '#ffe0b2',   // orange
  'F7': '#e8eaf6'    // indigo
};
```

---

## Webhook URLs (all pointing to @410)

```
Debug webhook (Katana + WASP + ShipStation):
https://script.google.com/macros/s/AKfycbwgacIohk10D3wdNp5EoHQkB1L890-7SMIF3ct6HXopcEgdg_ySS-f6y6QS1hIxe3eIiw/exec

GAS_WEBHOOK_URL (sync script → debug script for enginPreMark_):
Same URL as above
```

### Where the URL is configured:
- **Katana**: Settings → Webhooks → "WASP Inventory Sync" → Endpoint URL
- **WASP**: Callouts → "Katana Sync - Item Added" → URL
- **WASP**: Callouts → "Katana Sync - Item Removed" → URL
- **ShipStation**: Settings → Webhooks → SHIP_NOTIFY → URL
- **Sync script**: Script Properties → GAS_WEBHOOK_URL

**CRITICAL**: When redeploying, ALL 5 URLs must be updated simultaneously.

---

## Deployment History

| Version | Date | Changes |
|---------|------|---------|
| @405 | Pre-2026-03-18 | Original deployment |
| @406 | 2026-03-19 | F2 lock fix, F3 transfers, F4 lot fallback, F5 concurrent fix, F7 health check |
| @407 | 2026-03-19 | F1 MMH Mayfair QA-Hold-1 → PRODUCTION |
| @408 | 2026-03-19 | F1 allow retry after Failed PO receive |
| @409 | 2026-03-19 | F1 skip open-row diff when receive handler just processed |
| @410 | 2026-03-19 | F1 MMH Kelowna RECEIVING-DOCK → PRODUCTION |
