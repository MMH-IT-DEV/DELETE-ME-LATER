# F3 — Stock Transfer Flow: Setup, Rules & Location Mapping

## What F3 Does

F3 polls Katana for stock transfers every 1 minute. When a transfer reaches "received"
status, F3 removes items from WASP at the source site and adds them at the destination
site. If a transfer is reverted (status changes to "draft"), F3 automatically reverses
the WASP moves.

## Polling Config

- **Interval**: 1 minute (`POLL_INTERVAL_MINUTES`)
- **Min created date**: 2026-03-01 (`MIN_CREATED_DATE`)
- **Max age**: 30 days (`MAX_AGE_DAYS`)
- **Cursor**: `F3_LAST_SYNC_TIMESTAMP` in CacheService (6-hour TTL)
- **Lock**: `LockService.getScriptLock().tryLock(5000ms)`
- **Trigger function**: `pollKatanaStockTransfers`

## Location Mapping

### Actual Transfer Routes

| Route | Happens? |
|-------|----------|
| Storage Warehouse → MMH Kelowna | Yes |
| MMH Kelowna → MMH Mayfair | Yes |
| MMH Kelowna → Storage Warehouse | Yes |
| MMH Kelowna → Shopify | Skipped (F3_SKIP_LOCATIONS) |
| Shopify → MMH Kelowna | Skipped (F3_SKIP_LOCATIONS) |
| MMH Kelowna → Amazon USA | Skipped (F3_SKIP_LOCATIONS) |
| MMH Mayfair → anywhere | Never happens |

### WASP Location Mapping (F3_LOCATION_OVERRIDES)

| Katana Location | Direction | WASP Site | WASP Location |
|-----------------|-----------|-----------|---------------|
| MMH Kelowna | Source (FROM) | MMH Kelowna | Searches all locations, removes where found |
| MMH Kelowna | Destination (TO) | MMH Kelowna | **PRODUCTION** |
| MMH Mayfair | Destination (TO) only | MMH Mayfair | **PRODUCTION** |
| Storage Warehouse | Both directions | Storage Warehouse | **SW-STORAGE** |

**Source (FROM) behavior**: F3 searches ALL WASP locations at the source site
(SHOPIFY, PRODUCTION, RECEIVING-DOCK, PROD-RECEIVING, etc.) for each item.
It removes from wherever the stock is found. No fixed "from" location.

**Destination (TO) behavior**: Items always land at the mapped location
(PRODUCTION or SW-STORAGE).

### Config Files

**F3_LOCATION_OVERRIDES** in `scripts/debug/00_Config.js`:
```javascript
var F3_LOCATION_OVERRIDES = {
  'MMH Kelowna': { site: 'MMH Kelowna', location: 'PRODUCTION' },
  'MMH Mayfair': { site: 'MMH Mayfair', location: 'PRODUCTION' },
  'Storage Warehouse': { site: 'Storage Warehouse', location: 'SW-STORAGE' }
};
```

**F3_SITE_MAP** (fallback) in `scripts/debug/07_F3_AutoPoll.js` — matches above.

**Config source priority**: `KATANA_LOCATION_TO_WASP` loads first, then
`F3_LOCATION_OVERRIDES` overwrites (last wins). Fixed 2026-03-18 — previously
the order was reversed causing F3_LOCATION_OVERRIDES to be ignored.

## Status Handling

### Sync statuses (triggers WASP sync)
`['completed', 'done', 'received', 'partial', 'in_transit']`

### Reverse statuses (triggers WASP reversal)
`['cancelled', 'voided', 'draft']`

Note: Katana uses "draft" for reverted STs (confirmed 2026-03-18).

### Skip locations (no WASP sync)
`['amazon usa', 'shopify']`

## Deduplication

Each ST is tracked in the **StockTransfers** tab in the debug sheet. Once a ST
has status Synced/Failed/Skip/Reversed/Partial, it won't be re-processed on
subsequent polls.

## Activity Log Format

Sub-items show: `SKU xQTY UOM  lot:XXX  exp:YYYY  move FROM -> TO`

Route (`move FROM -> TO`) is always shown on every sub-item.
Lot and expiry appear before the route when present.

## Key Files

| File | Role |
|------|------|
| `scripts/debug/07_F3_AutoPoll.js` | Polling, processing, sync, reverse logic |
| `scripts/debug/00_Config.js` | F3_LOCATION_OVERRIDES, KATANA_LOCATION_TO_WASP |
| `scripts/debug/08_Logging.js` | Activity log writing, format helpers |

## Utility Functions

- `setupF3Complete` — creates StockTransfers tab + trigger (run once)
- `testF3Poll` — manual poll test
- `resetF3CursorAndPoll` — reset cursor to MIN_CREATED_DATE and poll
- `clearF3SheetAndPoll` — clear all StockTransfers entries and poll fresh
- `removeF3Trigger` — stop the 1-minute polling
