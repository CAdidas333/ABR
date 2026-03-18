# ABR Tool — Project Context for Claude Code

## What This Project Is

Auto Bank Reconciliation (ABR) tool for Jim Coleman Automotive -- a 7-location
dealership group in the Maryland/DC area. Replaces a deceased developer's custom
Excel/VBA tool. Controllers use this to reconcile bank statements against the
Reynolds & Reynolds (R&R) DMS General Ledger every month.

**Owner:** Chris Whitney (controller / operator)
**Platform:** Excel/VBA (non-negotiable -- this is what controllers use daily)
**Banks:** Bank of America (primary, tested) + Truist (parser built, not yet tested)
**DMS:** Reynolds & Reynolds (R&R)
**Locations:** 7 Jim Coleman stores (CAD, TOY, NIB, NSS, NEC, HON, JLR)

## Current State

The tool is functional and tested against real Toyota May 2025 data:

- **738/788 bank transactions auto-accepted** (93.6% auto-accept rate)
- **95.7% total match rate** (754/788 matched)
- **$0.00 net reconciling difference** on all matched pairs
- **Zero false positives** in spot-check verification
- **17 items staged for manual review** (down from 619 in Chris's manual approach)

## Matching Algorithm

Pass-rule architecture (not weighted scoring). Researched against BlackLine, Oracle
Cash Management, Treasury Software, and real-world dealership reconciliation practices.

Key principles (from Chris's feedback and industry research):
- Amount must be EXACT. Penny off = red flag, never auto-accepted.
- Check# is definitive. Check# + exact amount = 100% regardless of date gap.
- Uniqueness drives confidence. One candidate at exact amount = it's the match.
- Date is for disambiguation only, not a confidence factor.
- Never auto-commit. Everything is staged for controller review.

See README.md for the full pipeline description.

## Critical VBA Patterns

**NEVER use `Dim x As New ClassName` inside a loop.** VBA reuses the same object
instance. This bug has bitten us THREE times:
1. LoadBankTransactions -- all 788 entries pointed to last row
2. LoadDMSTransactions -- same issue
3. RunScoredMatching Phase 1 -- all candidates pointed to last pair

Always: `Dim x As ClassName` before the loop, `Set x = New ClassName` inside.

**VBScript.RegExp is not available on macOS Excel.** Use native VBA string operations
(Mid, InStr, Replace, character-by-character walking).

**Class modules on macOS must NOT have VERSION 1.0 CLASS headers.** Strip them before
importing.

## File Organization

```
src/vba/modules/     -- 12 VBA standard modules
src/vba/classes/     -- 3 class modules (clsTransaction, clsMatchResult, clsMatchGroup)
src/vba/forms/       -- 4 UserForms (not yet wired up)
tests/               -- Python reference implementation + 21 test scenarios
build/               -- Build config and test data generator
```

## V2 Backlog

- Worldpay mixed-sign CVR (net-of-fee deposits)
- ARP returned check parsing (check# in description text)
- Business-day date conversion (WORKDAY.INTL equivalent)
- Many-to-many matching (payroll)
- Wire up UserForms (currently all operations via Immediate Window)
- Reversal pair detection improvements
- Hungarian algorithm for duplicate-amount optimal assignment
- Per-vendor fee tolerance configuration

## What NOT to Do

- Do not auto-commit any match. Everything goes through staging.
- Do not use VBScript.RegExp (macOS incompatible).
- Do not use `Dim As New` inside loops.
- Do not hardcode row ranges (always use GetLastRow/GetNextRow).
- Do not guess at algorithm design -- research first, design second.
- Do not import real client financial data into the repo (gitignored).
