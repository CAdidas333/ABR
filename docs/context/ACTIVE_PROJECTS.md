# ABR — Active Projects
*Last updated: 2026-03-20*
*Living document -- updated each session*

---

## CURRENT PRIORITY: V1 Stabilization and Testing
**STATUS: IN PROGRESS**
**PRIORITY: HIGH**

### What's Done
- [x] BofA bank statement parser (sectioned CSV)
- [x] R&R DMS GL parser (9-column XLSX)
- [x] Pass-rule matching engine (Phases -1 through 4)
- [x] CVR many-to-one and reverse split matching
- [x] Prior-period GL fallback (April import + re-run)
- [x] Staging/accept/reject workflow
- [x] Manual match creation
- [x] Audit trail logging
- [x] Reset-and-rerun utility
- [x] Tested on real Toyota May 2025 data: 95.7% match rate, $0.00 net diff
- [x] GitHub repo created and pushed (https://github.com/CAdidas333/ABR)
- [x] Documentation (README, CLAUDE.md, context docs)
- [x] BofA BAI flat columnar CSV parser (Honda format — 14 columns with BAI codes)
- [x] Tested on real Honda Feb 2026 data: 95.7% match rate, 90.3% auto-accept
- [x] Sweep/securities exclusion (BAI 501/233 → reconciling items)
- [x] Outstanding deposit protection (BATCH date guard)
- [x] Phase 1 check-number veto (prevents false matches on duplicate amounts)
- [x] Prior outstanding checks parser for bank rec format
- [x] Multi-agent QA: verified all 6 outstanding deposits, 0 false positives

### What's Next
- [ ] Get January 2026 Honda bank rec from Jay (to match 18 prior-period checks)
- [ ] Wire up UserForms (frmDashboard, frmMatchReview, frmManualMatch, frmCVRGrouping)
- [ ] Improve Phase -1 reversal detection (AHM ACH pairs with same ref not caught)
- [ ] Parse ARP returned check numbers from description text
- [ ] Add mixed-sign CVR fragment support (Worldpay net-of-fee deposits)
- [ ] Handle Truist floorplan sweep transfers as reconciling items
- [ ] Test with a third location
- [ ] Build the automated workbook generator (build script)

### How to Run/Test
```bash
# In Excel VBA Immediate Window:
ModResetAndRerun.ResetAndRerun    ' Full reset + matching pipeline + auto-accept

# Or step by step:
ModImportBank.ImportBankStatement "/path/to/Bank Statement.csv"
ModImportDMS.ImportDMSExport "/path/to/DMS GL.xlsx"
ModMatchEngine.RunMatching ModImportBank.LoadBankTransactions(), ModImportDMS.LoadDMSTransactions()
ModStagingManager.AcceptAllHighConfidence
```

---

## BACKLOG

### V2: Advanced Matching
**STATUS: NOT STARTED**
- Worldpay mixed-sign CVR (gross minus processing fees)
- Many-to-many matching (payroll)
- Business-day date conversion (WORKDAY.INTL)
- Per-vendor fee tolerance configuration
- Hungarian algorithm for duplicate-amount optimal assignment
- Reversal pair detection improvements

### V3: Multi-Location Deployment
**STATUS: NOT STARTED**
- Test with all 7 locations
- Workbook generator for per-location instances
- Truist bank parser validation
- Training materials for controllers

### V4: Reporting
**STATUS: NOT STARTED**
- Power BI reporting layer for cross-location visibility
- Month-over-month trend analysis
- Outstanding items aging report
- Stale check detection (180+ days for abandoned property compliance)
