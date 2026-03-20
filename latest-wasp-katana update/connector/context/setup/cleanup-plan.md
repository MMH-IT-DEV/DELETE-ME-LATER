# Cleanup Plan — Connector Sheet + Menu

## Problem
The account section has 31 physical columns with duplicates (MFA x3, Privileged x3, Source Request ID x2).
Migrations haven't fully cleaned up. Need a forced rebuild, not another migration attempt.

---

## Step 1: Force rebuild of Access Control account section

The schema defines 14 columns but the sheet still has 31.
Instead of migrating, we will:
1. Read all account rows using the CURRENT sheet headers (whatever they are)
2. Map each row to the 14 new column keys by matching header names
3. Rewrite the account header row (14 columns)
4. Rewrite all account data rows (14 columns)
5. Clear everything in columns 15-31
6. Delete excess physical columns from the sheet

Files: workbook.js (add a forceRebuildAccountSection_ function)

---

## Step 2: Define the final 14 account columns

1. Access ID
2. System
3. Person
4. Company Email
5. Platform Account
6. Access Level
7. Access Status
8. MFA
9. Privileged
10. Last Verified
11. Next Review Due
12. Review Status
13. Source Request
14. Audit Log

---

## Step 3: Clean up the menu

Current menu has items for activity tracking that no longer exists.

REMOVE from menu:
- Refresh Katana Accounts (activity tracking — removed feature)
- Refresh Activity Signals (activity tracking — removed feature)
- Refresh Shopify Activity (activity tracking — removed feature)

RENAME for clarity:
- "Setup 3-Tab Workbook" → "Setup Sheet"
- "Run Daily Maintenance" → move to Advanced (it's automated, rarely manual)

Final menu:
```
System Security
├── Sync Requests
├── Generate Reviews
├── Refresh Sheet
├── Advanced ►
│   ├── Setup Sheet
│   ├── Setup Triggers
│   ├── Run Maintenance
│   └── Test Slack
```

4 main items, 4 advanced items. Clean.

Files: active-directory-sync.js (onOpen menu)

---

## Step 4: Verify all tabs

After rebuild:
- ACCESS REQUESTS: 25 columns (request schema — unchanged)
- ACTIVE SYSTEM ACCOUNTS: 14 columns (rebuilt)
- ARCHIVED REQUESTS: 25 columns (same as requests)
- REVIEW CHECKS: 24 columns (13 visible + hidden — unchanged)
- INCIDENTS: 18 columns (12 visible + hidden — unchanged)

---

## Execution order
1. Update menu (active-directory-sync.js)
2. Add forceRebuildAccountSection_ (workbook.js)
3. Call it from refreshSheet and setupSheet
4. Push + deploy
5. User runs Refresh Sheet once → clean 14-column account section
