#!/bin/bash
# ============================================================================
# ABR Honda Setup — Opens workbook in Excel and imports VBA modules
#
# Usage: ./build/setup_honda.sh
# ============================================================================

set -e

PROJECT_DIR="$(cd "$(dirname "$0")/.." && pwd)"
DIST_DIR="$PROJECT_DIR/dist"
SRC_VBA="$PROJECT_DIR/src/vba"
WORKBOOK="$DIST_DIR/ABR_HON.xlsx"

echo "============================================"
echo "  ABR Honda Setup"
echo "============================================"

# Check workbook exists
if [ ! -f "$WORKBOOK" ]; then
    echo "Generating workbook..."
    python3 "$PROJECT_DIR/build/generate_workbook.py" --location HON --output "$DIST_DIR"
fi

echo ""
echo "Workbook: $WORKBOOK"
echo ""
echo "STEP 1: Opening workbook in Excel..."
echo "        When Excel opens, Save As -> .xlsm (macro-enabled)"
echo ""
open "$WORKBOOK"

echo "STEP 2: After saving as .xlsm, open VBA Editor:"
echo "        Mac: Fn + F11  (or Tools > Macro > Visual Basic Editor)"
echo ""
echo "STEP 3: Import VBA modules (File > Import File... for each):"
echo ""
echo "  Modules (12 files):"
for f in "$SRC_VBA/modules/"*.bas; do
    echo "    $(basename "$f")"
done
echo ""
echo "  Classes (3 files):"
for f in "$SRC_VBA/classes/"*.cls; do
    echo "    $(basename "$f")"
done
echo ""
echo "  Source folder: $SRC_VBA"
echo ""
echo "STEP 4: In the Immediate Window (Ctrl+G), run:"
echo ""
echo "  Import bank statement:"
echo "    ModImportBank.ImportBankStatement \"$PROJECT_DIR/docs/Honda Bkstmt.csv\""
echo ""
echo "  Import Feb GL:"
echo "    ModImportDMS.ImportDMSExport \"$PROJECT_DIR/docs/honda Feb GL.xlsx\""
echo ""
echo "  Import Jan GL (prior period):"
echo "    ModImportDMS.ImportDMSExport \"$PROJECT_DIR/docs/honda Jan GL.xlsx\""
echo ""
echo "  Run matching + auto-accept:"
echo "    ModResetAndRerun.ResetAndRerun"
echo ""
echo "============================================"
echo "  Or for one-shot (after importing VBA):"
echo "  Paste into Immediate Window:"
echo ""
echo "  ModImportBank.ImportBankStatement \"$PROJECT_DIR/docs/Honda Bkstmt.csv\""
echo "  ModImportDMS.ImportDMSExport \"$PROJECT_DIR/docs/honda Feb GL.xlsx\""
echo "  ModImportDMS.ImportDMSExport \"$PROJECT_DIR/docs/honda Jan GL.xlsx\""
echo "  ModMatchEngine.RunMatching ModImportBank.LoadBankTransactions(), ModImportDMS.LoadDMSTransactions()"
echo "  ModStagingManager.AcceptAllHighConfidence"
echo "============================================"
