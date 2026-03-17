# ABR — Auto Bank Reconciliation Tool

Excel/VBA bank reconciliation tool for Jim Coleman Automotive (7 locations).
Reconciles Bank of America and Truist bank statements against Reynolds & Reynolds DMS GL exports.

## What It Does

1. **Imports** bank statement CSVs (BofA or Truist) and R&R DMS GL exports
2. **Matches** transactions using weighted confidence scoring (amount, check number, date, description)
3. **Stages** all proposed matches for controller review — nothing is auto-committed
4. **Handles CVR** (Customer Vehicle Receivable) many-to-one matching where bank fragments sum to a single DMS entry
5. **Exports** reconciliation reports and outstanding items for carry-forward

## Key Design Decisions

- **Confidence scoring (0-100%):** Amount (40%) + Check# (25%) + Date proximity (25%) + Description (10%)
- **Check number veto:** Mismatched check numbers cap the score at 30, regardless of other factors
- **Never auto-commit:** Every match is staged for controller approval
- **CVR subset-sum solver:** Finds 2-6 bank fragments that sum to a DMS CVR entry, with 2-second timeout
- **Full audit trail:** Every action logged with timestamp, user, and session ID

## Project Structure

```
ABR/
├── src/vba/
│   ├── modules/        # 11 VBA modules (.bas)
│   ├── classes/        # 3 class modules (.cls)
│   └── forms/          # 4 UserForms (.frm)
├── build/
│   ├── build_config.json       # Sheet definitions, column specs
│   ├── generate_workbook.py    # Creates .xlsx with all 9 sheets
│   └── generate_test_data.py   # Generates 21 test scenarios
├── tests/
│   ├── matching_engine.py      # Python reference implementation
│   ├── test_*.py               # 91 tests (scoring, matching, CVR, parsing)
│   └── data/                   # Generated test data (65 files)
├── dist/                       # Generated workbooks (7 locations)
└── CLAUDE.md                   # Project context
```

## Deployment

1. Open a workbook from `dist/` (e.g., `ABR_JCC.xlsx`) in Excel
2. Save As → `.xlsm` (Excel Macro-Enabled Workbook)
3. Open VBA Editor (`Alt+F11`)
4. File → Import File → import each `.bas`, `.cls` from `src/vba/modules/` and `src/vba/classes/`
5. For UserForms: create forms in the VBA Editor and add controls as documented in each `.frm` file
6. Run `ModMain.Step1_ImportBankStatement` to begin

## Locations

| Code | Name | Bank |
|------|------|------|
| JCC | Jim Coleman Cadillac | Bank of America |
| JCH | Jim Coleman Honda | Bank of America |
| JCI | Jim Coleman Infiniti | Bank of America |
| JCL | Jim Coleman Lexus | Truist |
| JCV | Jim Coleman Volvo | Truist |
| JCA | Jim Coleman Acura | Bank of America |
| JCJ | Jim Coleman Jaguar Land Rover | Truist |

## Running Tests

```bash
pip3 install pytest
cd tests && python3 -m pytest -v
```

91 tests validate the matching algorithm, CVR logic, confidence scoring, and CSV parsing.
Full-month simulation (123 bank + 113 DMS transactions) achieves 89.4% match rate.

## Regenerating Workbooks

```bash
pip3 install openpyxl
python3 build/generate_workbook.py --all     # All 7 locations
python3 build/generate_workbook.py --location JCC  # Single location
```
