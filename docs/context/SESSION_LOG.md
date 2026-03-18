# ABR — Session Log
*Running history of every working session -- never delete entries*

---

## Session: 2026-03-17 — Algorithm Redesign and Live Testing

**Focus:** Rewrote the matching algorithm from weighted scoring to pass-rule architecture, debugged VBA issues, tested against real Toyota May 2025 data.

**What We Built/Changed:**
- Rewrote ModMatchEngine.bas: pass-rule architecture (Phase -1 through Phase 4)
- Phase -1: DMS self-canceling pair detection (voids/reversals)
- Phase 0 Rule 0: Check# confirmed + exact amount = 100%
- Phase 0 Rule 1: Unique exact amount = 95% (with date corroboration) or 85% (without)
- Phase 1: Scored matching for duplicate-amount groups with margin-based confidence
- Phase 4: Near-amount matching ($0.01 tolerance, 55%, always review)
- Revised CVR scoring: sum accuracy = hard gate, date clustering = 60%, fragment penalty = 30%
- Fixed duplicate CVR staging bug (best-subset-only per target)
- Added audit trail columns to Reconciled sheet (breakdown, check match, amount diff)
- Created ModResetAndRerun.bas for one-shot reset and re-run
- Added .gitignore, README.md, updated CLAUDE.md
- Created GitHub repo and pushed all commits

**Decisions Made:**
- Pass-rule architecture over weighted scoring (research-backed: BlackLine, Oracle, Treasury Software)
- Amount must be exact for auto-matching; penny off = red flag, never auto-accepted
- Check# is definitive regardless of date gap (prior-period checks = 100% when check# confirms)
- Uniqueness is the real confidence driver, not date proximity
- CVR max fragments capped at 4 (was 6) to avoid coincidental sums
- Date is for disambiguation only, not confidence scoring

**Bugs Fixed:**
- `Dim As New` object reuse bug in RunScoredMatching Phase 1 (third occurrence of this VBA bug)
- `Dim As New` bug also present in Rules 0, 1, and Phase 4 (fixed proactively)
- Duplicate April DMS import (2741 rows instead of 1784) breaking Rule 1 uniqueness check
- CVR staging emitting duplicate matches for same target (best-subset-only fix)
- Summary script checking wrong column for BankData IsMatched (col 9 vs col 10)
- Overflow error in match rate calculation (Long division, fixed with CDbl)

**Research Conducted:**
- Web research on bank reconciliation best practices (BlackLine, Oracle, Trintech, Treasury Software)
- Analyzed Chris's manual Excel reconciliation workbook (WORKDAY.INTL, composite keys, cascading strategies)
- Devil's advocate challenge agent identified 5 critical fixes
- QA agent analyzed real data: 155 check# matches, 418 unique amounts, 30 duplicate-amount groups
- All findings incorporated into final algorithm

**Results:**
- 738/788 auto-accepted (93.6%) at 85%+ confidence
- 754/788 total matched (95.7%)
- 17 staged for review, 34 unmatched
- $0.00 net reconciling difference
- Zero false positives in 9-item spot check

**Next Steps:**
- Wire up UserForms for controller-friendly interface
- Parse ARP returned check numbers from description text
- Mixed-sign CVR for Worldpay net-of-fee deposits
- Test with second location
- Review the 17 staged items and investigate the 34 unmatched

---

## Session: 2026-03-17 (Earlier) — Initial Build and Data Calibration

**Focus:** Built initial VBA modules, calibrated against real Toyota May 2025 data, fixed macOS compatibility issues.

**What We Built/Changed:**
- Rewrote ModImportBank for real BofA sectioned CSV format (D-Mon dates in Checks section)
- Rewrote ModImportDMS for real R&R 9-column XLSX format
- Fixed VBScript.RegExp crashes (macOS incompatible) -- rewrote with native VBA string ops
- Fixed VERSION 1.0 CLASS compile errors on macOS
- Fixed `Dim As New` object reuse bug in LoadBankTransactions and LoadDMSTransactions
- Python reference implementation validated at 95.7% match rate

**Bugs Fixed:**
- VBScript.RegExp CreateObject crash on macOS (3 locations)
- VERSION 1.0 CLASS header causing compile errors on macOS
- Attribute VB_Name stripped then restored (needed for correct module naming on import)
- Ambiguous name detected: Class_Initialize (code pasted twice)
- GetOpenFilename sandbox failure on macOS (workaround: pass file paths directly)
- Zero matches due to Dim As New object reuse in both loading functions

**Next Steps:**
- Algorithm redesign (led to the pass-rule rewrite above)
