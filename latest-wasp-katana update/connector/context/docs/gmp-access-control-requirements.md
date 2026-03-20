# GMP Access Control Requirements — MyMagicHealer
Researched: 2026-03-18 | Industry: NHP (Health Canada) | Status: Preparing for audit

---

## Regulations that apply:
1. GUI-0158 v4.0 — NHP GMP Guide (enforceable since March 4, 2026)
2. GUI-0050 — Annex 11, Computerized Systems (the key one for this tracker)
3. SOR/2003-196 Part 3 — the actual law (Sections 43-62)
4. PIC/S PI 011-3 — operational guidance for computerized systems

---

## What auditors specifically look for:
1. Every person with access is listed (completeness)
2. No departed employees still have access (currency)
3. Documented access matches actual system config (consistency)
4. Every access change traces to an approval + reason (traceability)
5. Access revoked promptly when no longer needed (timeliness)
6. Periodic review evidence with decisions documented (review evidence)
7. No conflicting roles — reviewer can't review their own access (separation of duties)
8. No shared accounts on GMP-critical systems
9. Training completed before access was granted

---

## What you check during each review:
1. Does the account still exist in the system?
2. Is the access level correct? (matches what's documented)
3. Does the person still work here?
4. Is the email/username correct?
5. Is MFA turned on?
6. Is it a personal account, not shared?

---

## System risk classification:
- HIGH: Katana (manufacturing), 1Password (controls all access), Google Drive (SOPs/batch records)
- MEDIUM: Shopify (orders), Amazon Seller (marketplace), Wasp
- LOW: ShipStation (shipping)

---

## Review frequency (implemented):
- HIGH systems — Admin: 30 days, Full Access: 60 days, Read-Only: 90 days
- MEDIUM systems — Admin: 60 days, Full Access: 90 days, Read-Only: 180 days
- LOW systems — Admin: 90 days, Full Access: 180 days, Read-Only: 365 days

---

## What must be documented for every access change:
1. Who was granted/changed/revoked access (full name)
2. What system and what level of access
3. When it happened (date and time)
4. Who authorized the change
5. Why — business justification
6. Previous state vs new state

---

## When to do an unscheduled review:
1. Employee leaves or changes role (immediate)
2. Security incident occurs
3. System upgrade or migration
4. Audit finding (internal or external)
5. Regulatory change

---

## Record retention:
- Minimum 3 years, or 1 year past the latest product expiry date — whichever is longer

---

## Common audit gaps (things that get flagged):
1. No formal access control SOP
2. Shared/generic accounts on GMP systems
3. Access granted but never reviewed
4. Departed employees still have active access
5. Systems not classified by risk level
6. No audit trail for access changes
7. Access given before training completed
8. Admin accounts used for routine operations
9. No business justification documented for access
10. No MFA on critical systems
11. Google Docs used for SOPs without validation

---

## Are we currently following it?
- [x] Access requests documented with reason (Slack command + sheet)
- [x] Approval tracked (Approval + Approved By columns)
- [x] Periodic reviews scheduled (Review Checks tab with GMP frequencies)
- [x] Risk-based review frequency (High/Medium/Low system classification)
- [x] Review decisions documented (Reviewer Action + Review Status)
- [x] Shared account flag available (column exists, needs filling)
- [x] MFA column available (needs filling for each account)
- [ ] Training completion tracked (column proposed, not yet added)
- [ ] Employment status tracked (column proposed, not yet added)
- [ ] Formal access control SOP written
- [ ] Google Drive validated as document management system
- [ ] Separation of duties for reviews (only Erik reviews currently)

---

## Sources:
- SOR/2003-196: https://laws-lois.justice.gc.ca/eng/regulations/sor-2003-196/
- GUI-0158 v4.0: https://www.canada.ca/en/health-canada/services/drugs-health-products/compliance-enforcement/good-manufacturing-practices/guidance-documents/guide-natural-health-products-0158.html
- GUI-0050: https://www.canada.ca/en/health-canada/services/drugs-health-products/compliance-enforcement/good-manufacturing-practices/guidance-documents/annex-11-guide-computerized-systems-gui-0050.html
- PIC/S PI 011-3: https://picscheme.org/docview/3444
