# ABR — Session Log
*Running history of every working session -- never delete entries*

---

## Session: 2026-03-20 — Honda Feb 2026 Testing and Multi-Agent QA

**Focus:** Tested matching engine against Honda Feb 2026 real data (second location). Built new bank parser, ran multi-agent QA/accounting/code review, implemented three fixes.

**What We Built/Changed:**
- New `ParseBofABAI` parser for BofA BAI flat columnar CSV (14-column format with BAI codes)
- Format auto-detection updated: BOFA sectioned vs BOFA_BAI vs TRUIST
- Python reference implementation updated to pass-rule architecture (was still using old weighted scoring)
- Sweep/securities exclusion: BAI 501/233 filtered from matching pool as reconciling items
- Outstanding deposit protection: BATCH deposits can't match bank txns >3 days earlier
- Check-number veto in Phase 1: bank+DMS both have check# but they differ = skip
- Prior outstanding checks import stub (VBA) + parser (Python) for bank rec format
- Fixed timestamp format bug across 6 VBA modules (`HH:MM:SS` → `h:mm:ss`)
- `is_reconciling_item` and `reconciling_type` fields added to Transaction dataclass

**Honda Feb 2026 Results (Final):**
- 556 bank transactions (538 matchable + 18 reconciling items excluded)
- 500 matches / 538 matchable = **95.7% bank match rate**
- 486/538 auto-accepted at 85%+ = **90.3% auto-accept rate**
- 213 check# confirmed matches at 100% avg confidence
- $0.07 net amount difference
- All 6 outstanding deposits correctly unmatched (PASS)
- 0 false positives on check-number mismatches (after Phase 1 veto fix)
- 23 unmatched bank: 18 prior-period checks + 5 ACH (2 Truist sweeps + 3 other)
- 154 unmatched DMS: 111 checks (73 on outstanding list) + 41 FINDEP + 2 BATCH deposits

**Decisions Made:**
- BAI format detection by "bai code" in header, checked before Truist (BAI header contains "debit/credit" which would false-match)
- Sweep/securities are "reconciling items" excluded from matching (no GL counterpart)
- BATCH deposits get a 3-day date guard to prevent false matches to earlier bank deposits
- Phase 1 scored matching must veto pairs where both sides have different check numbers
- BankSource stays "BOFA" for both sectioned and BAI formats (matching engine doesn't branch on it)

**Bugs Fixed:**
- Timestamp format `HH:MM:SS` showed month instead of minutes in all VBA modules (display bug)
- Phase 1 scored matching produced 8 false positives by pairing bank checks with wrong DMS checks at same amount (fixed with check-number veto)
- Outstanding deposits falsely matched when coincidentally same amount as earlier bank deposits (fixed with date guard)

**Multi-Agent QA Findings:**
- QA agent: verified all 207 check# matches, 232 unique amount matches, 3 near-amount matches; found the 8 Phase 1 false positives
- Accounting agent: all 13 of Jay's outstanding journal entries map 1:1 to our FINDEP items; Phase -1 missed some AHM ACH self-canceling pairs; Truist FL DR&CR debits are floorplan sweep transfers
- Senior dev: 18 findings, 0 actual critical bugs (check# stripping verified safe at DMS import), timestamp format bug confirmed, prior-period matching mutates input objects

**Next Steps:**
- Ask Jay for the January 2026 bank rec to match the 18 prior-period bank checks
- Improve Phase -1 to catch AHM ACH self-canceling pairs (reference matching too strict)
- Consider widening near-amount tolerance for ACH items ($0.03 Zurich mismatch)
- Handle Truist FL DR&CR sweep transfers as reconciling items
- Wire up UserForms
- Test with a third location

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
