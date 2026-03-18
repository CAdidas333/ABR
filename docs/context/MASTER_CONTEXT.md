# ABR (Auto Bank Reconciliation) — Master Context
*Last updated: 2026-03-17*

---

## Who This Is For
This file is the authoritative context document for any Claude instance (chat or Claude Code)
working on this project. Read this first. Always update at the end of every session.

---

## The Person
- Name: Chris Whitney
- Role: Controller / operator at Jim Coleman Automotive
- Goal: Replace a deceased developer's custom bank reconciliation tool with something better
- Ownership: Jim Coleman Automotive (internal tool)

---

## The Problem This Solves
Jim Coleman Automotive is a 7-location dealership group in the Maryland/DC area. The controllers at each store must reconcile bank statements against the Reynolds & Reynolds DMS General Ledger every month. The original custom Excel/VBA tool was built by a developer who passed away in 2024, and all licenses expired. Stores have been doing manual reconciliation since then, which is slower and error-prone. This project rebuilds the tool — better, not cloned.

---

## Architecture

### Matching Pipeline
Pass-rule architecture modeled after BlackLine, Oracle Cash Management, and Treasury Software:

```
Phase -1: Detect DMS self-canceling pairs (voids/reversals) and exclude
Phase  0: Pass Rules (deterministic, no scoring)
  Rule 0: Check# confirmed + exact amount -> 100% confidence
  Rule 1: Unique exact amount + date corroboration -> 95%
  Rule 1b: Unique exact amount, no date corroboration -> 85%
Phase  1: Scored matching for duplicate-amount groups
  Score on date proximity (70%) + description similarity (30%)
  Confidence = f(margin between best and runner-up candidate)
Phase  2: CVR many-to-one (bank fragments -> DMS lump sum)
Phase  3: Reverse split (DMS fragments -> bank lump sum)
Phase  4: Near-amount matching ($0.01 tolerance, always review)
Phase  5: Prior-period fallback (re-run against prior month GL)
```

### Data Flow
```
Bank Statement CSV (BofA sectioned format)
  -> ModImportBank.ParseBankOfAmerica()
  -> BankData sheet (788 rows for Toyota May 2025)

DMS GL Export XLSX (R&R 9-column format)
  -> ModImportDMS.ParseDMSFile()
  -> DMSData sheet (827 rows May + 957 rows April)

BankData + DMSData
  -> ModMatchEngine.RunMatching() (Phases -1, 0, 1, 4)
  -> ModMatchCVR.RunCVRMatching() (Phase 2)
  -> ModMatchCVR.RunReverseSplitMatching() (Phase 3)
  -> StagedMatches sheet (all proposals)
  -> Controller reviews, accepts/rejects
  -> Reconciled sheet (accepted matches)
```

---

## Tech Stack
- Language/Framework: Excel/VBA (non-negotiable -- controllers use Excel daily)
- Data Sources: Bank of America CSV, Truist CSV, Reynolds & Reynolds XLSX
- No external dependencies -- runs on any machine with Excel
- macOS compatible (no VBScript.RegExp, no COM objects)
- Python reference implementation for algorithm testing

---

## File Structure
```
ABR/
  src/vba/
    modules/        -- 12 VBA standard modules
    classes/        -- 3 class modules
    forms/          -- 4 UserForms (not yet wired up)
  tests/            -- Python reference + 21 test scenarios
  build/            -- Build config and test data generator
  docs/context/     -- THIS folder -- context docs
  CLAUDE.md         -- Claude Code session context
  README.md         -- Project documentation
  .gitignore        -- Excludes real financial data
```

---

## Key Decisions Made

1. **Pass-rule architecture over weighted scoring.** The original weighted-average approach (Amount*0.40 + Check*0.25 + Date*0.25 + Desc*0.10) produced unintuitive confidence scores. A controller shouldn't see 82% on an exact-match same-day transaction. Research confirmed BlackLine/Oracle/Treasury Software all use deterministic rules, not weighted scoring. (Session: 2026-03-17)

2. **Amount must be exact, not "close enough."** Even a penny off is a red flag per Chris's domain expertise and industry best practices. $0.02+ is not a candidate at all. (Session: 2026-03-17)

3. **Check# is definitive regardless of date.** A check written in April clearing 30 days later in May is 100% the same transaction when check# + amount match. Date gaps on checks are normal business operations. (Session: 2026-03-17)

4. **Uniqueness drives confidence.** If only one GL entry exists at that exact amount, it IS the match. The real confidence question is "how many other transactions could this be?" not "how many scoring factors confirm it." (Session: 2026-03-17)

5. **Never auto-commit.** Everything goes to StagedMatches for controller review. AcceptAllHighConfidence auto-accepts at 85%+ but they're still visible and reversible. (Original design principle)

6. **macOS compatibility required.** Chris uses a Mac. No VBScript.RegExp, no COM objects, class modules must not have VERSION 1.0 CLASS headers. (Session: 2026-03-17)

7. **Research before algorithm design.** Chris explicitly rejected a guessed algorithm. Always research best practices before designing domain-specific algorithms. (Session: 2026-03-17)

---

## What's Built
- BofA sectioned CSV parser (deposits, withdrawals, checks sections)
- R&R DMS 9-column XLSX parser (SRC-based type derivation)
- Truist CSV parser (built, not tested on real data)
- Pass-rule matching engine (Phases -1 through 4)
- CVR many-to-one and reverse split matching
- Staging/accept/reject workflow
- Manual match creation
- Prior-period GL fallback (import April, re-run)
- Complete audit trail logging
- Outstanding items import/export
- Dashboard with 5-step workflow
- Reset-and-rerun utility (ModResetAndRerun)
- Python reference implementation with 21 test scenarios

---

## What's NOT Built Yet
- UserForms (frmDashboard, frmMatchReview, frmManualMatch, frmCVRGrouping) -- designed but not wired up
- Worldpay mixed-sign CVR fragments (gross minus processing fees)
- ARP returned check parsing (check# embedded in description text)
- Business-day date conversion (WORKDAY.INTL equivalent)
- Many-to-many matching (payroll = 1 bank ACH to 8-12 DMS entries)
- Per-vendor fee tolerance configuration
- Multi-location deployment (tested on Toyota only)
- Power BI reporting layer (future state)

---

## Performance (Toyota May 2025 Real Data)

| Metric | Value |
|--------|-------|
| Bank transactions | 788 |
| Auto-accepted (85%+) | 738 (93.6%) |
| Total matched | 754 (95.7%) |
| Staged for review | 17 |
| Unmatched | 34 |
| Net difference | $0.00 |
| False positives | 0 |

Chris's manual approach: 98.1% match rate but 619 items (78.6%) needed manual review.
Our tool: 95.7% match rate with only 17 items (2.2%) needing review.

---

## Locations

| Code | Store |
|------|-------|
| CAD | Jim Coleman Cadillac |
| TOY | Jim Coleman Toyota |
| NIB | Jim Coleman Nissan/Infiniti of Bethesda |
| NSS | Jim Coleman Nissan Silver Spring |
| NEC | Jim Coleman Nissan Ellicott City |
| HON | Jim Coleman Honda |
| JLR | Jim Coleman Jaguar/Land Rover |

---

## Related Projects
- ICC (Inventory Command Center) -- separate project, same dealership group
