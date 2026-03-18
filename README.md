# ABR — Auto Bank Reconciliation

Excel/VBA bank reconciliation tool for Jim Coleman Automotive, a 7-location dealership group in the Maryland/DC area. Matches Bank of America commercial bank statements against Reynolds & Reynolds DMS General Ledger exports.

## What It Does

Automates the monthly bank-to-GL reconciliation that controllers do manually. Imports bank statements and DMS GL exports, runs a multi-phase matching engine, and stages all proposed matches for controller review. Nothing is auto-committed — every match requires human approval.

**Performance on real data (Jim Coleman Toyota, May 2025):**

| Metric | Value |
|--------|-------|
| Bank transactions | 788 |
| Auto-accepted (85%+ confidence) | 738 (93.6%) |
| Staged for review | 17 |
| Unmatched (outstanding) | 34 |
| Match rate | 95.7% |
| Net reconciling difference | $0.00 |
| False positives | 0 |

For comparison, the controller's manual Excel-based reconciliation achieved 98.1% match rate but required manual review of 619 items (78.6%). This tool reduces manual review from 619 items to 17.

## Architecture

### Matching Pipeline

The matching engine uses a pass-rule architecture modeled after BlackLine, Oracle Cash Management, and Treasury Software best practices:

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

### Key Design Decisions

- **Amount is a binary gate, not a weighted factor.** Exact match or not a candidate. Penny off = major red flag, never auto-accepted.
- **Check number is definitive.** Check# confirmed + exact amount = 100% confidence regardless of date gap. Prior-period checks clearing 30-90 days later are correctly matched.
- **Uniqueness drives confidence.** If only one GL entry exists at that exact amount, it's the match. Date gaps from business-day processing are noise, not evidence of mismatch.
- **Date is for disambiguation, not confidence.** Only matters when multiple candidates exist at the same amount.
- **CVR sum accuracy is a hard gate.** Fragments must sum within $0.01 or it's not a match.

### Data Formats

**Bank of America** -- Sectioned CSV with 3 transaction groups (Deposits, Withdrawals, Checks). Checks use D-Mon date format (e.g., "16-May"). Statement year extracted from first row.

**R&R DMS GL** -- 9-column XLSX: SRC, Reference#, Date, Port, Control#, Debit Amount, Credit Amount, Name, Description. Transaction types derived from SRC code + reference pattern.

## Project Structure

```
src/vba/
  modules/
    ModMain.bas           -- 5-step workflow orchestrator
    ModMatchEngine.bas    -- Core matching algorithm (pass rules + scored matching)
    ModMatchCVR.bas       -- Many-to-one and reverse split matching
    ModImportBank.bas     -- BofA and Truist bank statement parsers
    ModImportDMS.bas      -- R&R DMS GL parser
    ModStagingManager.bas -- Match staging, accept/reject, manual match
    ModConfig.bas         -- Configuration from Config sheet
    ModHelpers.bas        -- Date parsing, string cleaning, Levenshtein distance
    ModAuditTrail.bas     -- Session and action logging
    ModOutstanding.bas    -- Outstanding items carry-forward
    ModExport.bas         -- Month-end finalization and export
    ModResetAndRerun.bas  -- One-shot reset and re-run utility
  classes/
    clsTransaction.cls    -- Transaction data object
    clsMatchResult.cls    -- Match proposal with confidence and breakdown
    clsMatchGroup.cls     -- CVR group representation
  forms/
    frmDashboard.frm      -- Main dashboard
    frmMatchReview.frm    -- Match review interface
    frmManualMatch.frm    -- Manual match creation
    frmCVRGrouping.frm    -- CVR group builder

build/
  build_config.json       -- Store names, column definitions
  generate_test_data.py   -- Test data generator

tests/
  matching_engine.py      -- Python reference implementation
  test_matching.py        -- Matching algorithm tests
  test_import_parsing.py  -- Parser tests
  test_cvr_matching.py    -- CVR matching tests
  test_confidence_scoring.py -- Confidence scoring tests
  data/                   -- Test scenarios (21 scenarios)
```

## Excel Workbook Sheets

| Sheet | Purpose |
|-------|---------|
| Dashboard | 5-step workflow with status indicators |
| BankData | Imported bank transactions |
| DMSData | Imported DMS GL transactions |
| StagedMatches | Proposed matches awaiting review |
| Reconciled | Accepted matches |
| Outstanding | Carry-forward items from prior periods |
| Config | Configurable parameters (thresholds, weights) |
| AuditLog | Complete action log |
| Lookups | Reference data |

## Setup

1. Open the `.xlsm` workbook in Excel (macOS or Windows)
2. Import all `.bas` modules into VBA Editor (Tools > Macros > Visual Basic Editor)
3. Import all `.cls` class modules
4. The modules are embedded in the workbook after first save -- no re-import needed

## Usage

From the Immediate Window (Cmd+G in VBA Editor):

```vb
' Full workflow:
ModMain.Step1_ImportBankStatement
ModMain.Step2_ImportDMSData
ModMain.Step3_RunAutoMatching

' Or with file paths (bypasses file picker on macOS):
ModImportBank.ImportBankStatement "/path/to/Bank Statement.csv"
ModImportDMS.ImportDMSExport "/path/to/DMS GL.xlsx"

' Reset and re-run everything:
ModResetAndRerun.ResetAndRerun

' Accept all high-confidence matches:
ModStagingManager.AcceptAllHighConfidence

' Manual match:
ModStagingManager.CreateManualMatch bankRowID, dmsRowID
```

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

## Known Limitations (V2 Backlog)

- Worldpay net-of-fee deposits need mixed-sign CVR fragments
- ARP returned checks have check# embedded in description, not parsed
- Payroll transactions have no DMS counterpart (external payroll provider)
- No business-day conversion (using calendar days with wider windows)
- No many-to-many matching (payroll = 1 bank ACH to 8-12 DMS journal entries)
- UserForms not yet wired up (all operations via Immediate Window)

## macOS Compatibility Notes

- `VBScript.RegExp` is not available -- all string operations use native VBA
- `GetOpenFilename` may not work in sandboxed macOS Excel -- pass file paths directly
- Class modules must NOT have `VERSION 1.0 CLASS` headers when importing
- Standard modules MUST have `Attribute VB_Name` for correct import naming
