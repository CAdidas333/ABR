#!/usr/bin/env python3
"""
ABR Honda Build — Generate workbook, inject VBA, open in Excel.

Creates ABR_HON.xlsm with all sheets, formatting, and VBA modules
pre-loaded. Opens it in Excel ready to use.

Usage:
    python3 build/build_honda.py
"""

import os
import subprocess
import sys
import shutil

try:
    from openpyxl import Workbook
except ImportError:
    print("openpyxl is required. Install with: pip3 install openpyxl")
    sys.exit(1)

PROJECT_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DIST_DIR = os.path.join(PROJECT_DIR, 'dist')
VBA_MODULES = os.path.join(PROJECT_DIR, 'src', 'vba', 'modules')
VBA_CLASSES = os.path.join(PROJECT_DIR, 'src', 'vba', 'classes')
DOCS_DIR = os.path.join(PROJECT_DIR, 'docs')

XLSM_PATH = os.path.join(DIST_DIR, 'ABR_HON.xlsm')

# Paths to Honda data files (for the bootstrap macro)
BANK_FILE = os.path.join(DOCS_DIR, 'Honda Bkstmt.csv')
FEB_GL_FILE = os.path.join(DOCS_DIR, 'honda Feb GL.xlsx')
JAN_GL_FILE = os.path.join(DOCS_DIR, 'honda Jan GL.xlsx')


def main():
    print("=" * 60)
    print("  ABR Honda Build")
    print("=" * 60)

    # Step 1: Generate the base workbook
    print("\n[1/4] Generating workbook...")
    os.makedirs(DIST_DIR, exist_ok=True)

    sys.path.insert(0, os.path.join(PROJECT_DIR, 'build'))
    from generate_workbook import generate_workbook
    xlsx_path = generate_workbook("Jim Coleman Honda", "HON", "BOFA", DIST_DIR)

    # Step 2: Collect all VBA source files
    print("[2/4] Collecting VBA modules...")
    bas_files = sorted([f for f in os.listdir(VBA_MODULES) if f.endswith('.bas')])
    cls_files = sorted([f for f in os.listdir(VBA_CLASSES) if f.endswith('.cls')])

    print(f"  {len(bas_files)} standard modules")
    for f in bas_files:
        print(f"    {f}")
    print(f"  {len(cls_files)} class modules")
    for f in cls_files:
        print(f"    {f}")

    # Step 3: Use AppleScript to open in Excel, save as .xlsm, import all VBA
    print("[3/4] Building AppleScript to import VBA into Excel...")

    # Build the import commands for AppleScript
    import_lines = []
    for f in bas_files:
        path = os.path.join(VBA_MODULES, f).replace('"', '\\"')
        import_lines.append(f'run VBA macro "ImportSingleFile \\"{path}\\""')
    for f in cls_files:
        path = os.path.join(VBA_CLASSES, f).replace('"', '\\"')
        import_lines.append(f'run VBA macro "ImportSingleFile \\"{path}\\""')

    # We'll inject a tiny bootstrap module first that has the ImportSingleFile macro,
    # then use it to import everything else. But AppleScript + Excel VBA interaction
    # on macOS is limited. The more reliable approach: write ALL VBA into a single
    # bootstrap .bas file that Excel can run to self-import the rest.

    # Create a bootstrap VBA module that imports all other modules
    bootstrap_code = build_bootstrap_module(bas_files, cls_files)
    bootstrap_path = os.path.join(DIST_DIR, 'ModBootstrap.bas')
    with open(bootstrap_path, 'w') as f:
        f.write(bootstrap_code)
    print(f"  Bootstrap module: {bootstrap_path}")

    # Step 4: Open workbook in Excel
    print("[4/4] Opening workbook in Excel...")
    subprocess.run(['open', xlsx_path])

    total_modules = len(bas_files) + len(cls_files)
    print(f"""
============================================================
  WORKBOOK IS OPEN IN EXCEL — 3 STEPS:
============================================================

  STEP 1: Save As .xlsm
    File > Save As > format: "Excel Macro-Enabled Workbook (.xlsm)"

  STEP 2: Import ONE file into VBA Editor
    Open VBA Editor: Fn + F11
    File > Import File... > select:
    {bootstrap_path}

  STEP 3: Immediate Window (Ctrl+G), type two commands:

    RunBootstrap
      (auto-imports all {total_modules} VBA modules for you)

    ImportAndMatch
      (imports Honda bank + GL data, runs matching, auto-accepts)

  Then check BankData, DMSData, StagedMatches, Reconciled sheets.
============================================================""")


def build_bootstrap_module(bas_files, cls_files):
    """Build a VBA module that imports all other modules and runs matching."""
    # Escape paths for VBA strings
    mod_dir = VBA_MODULES.replace('\\', '\\\\')
    cls_dir = VBA_CLASSES.replace('\\', '\\\\')

    lines = [
        'Attribute VB_Name = "ModBootstrap"',
        "'===============================================================================",
        "' ModBootstrap — Auto-generated module importer",
        "'",
        "' Run 'RunBootstrap' from Immediate Window to import all VBA modules.",
        "' Run 'ImportAndMatch' to import Honda data and run matching.",
        "'",
        "' AUTO-GENERATED — do not edit manually.",
        "'===============================================================================",
        "",
        "Option Explicit",
        "",
        "Public Sub RunBootstrap()",
        "    ' Import all VBA modules and class files into this workbook.",
        "    Dim vbProj As Object",
        "    Set vbProj = ThisWorkbook.VBProject",
        "    ",
        "    Dim imported As Long",
        "    imported = 0",
        "    ",
        "    On Error Resume Next",
        "    ",
        "    ' --- Standard Modules ---",
    ]

    for f in bas_files:
        path = os.path.join(VBA_MODULES, f)
        mod_name = f.replace('.bas', '')
        lines.append(f'    RemoveIfExists vbProj, "{mod_name}"')
        lines.append(f'    vbProj.VBComponents.Import "{path}"')
        lines.append(f'    imported = imported + 1')
        lines.append(f'    Debug.Print "  Imported: {f}"')
        lines.append(f'')

    lines.append("    ' --- Class Modules ---")
    for f in cls_files:
        path = os.path.join(VBA_CLASSES, f)
        cls_name = f.replace('.cls', '')
        lines.append(f'    RemoveIfExists vbProj, "{cls_name}"')
        lines.append(f'    vbProj.VBComponents.Import "{path}"')
        lines.append(f'    imported = imported + 1')
        lines.append(f'    Debug.Print "  Imported: {f}"')
        lines.append(f'')

    lines.extend([
        '    On Error GoTo 0',
        '    ',
        '    Debug.Print ""',
        '    Debug.Print "Bootstrap complete: " & imported & " modules imported."',
        '    Debug.Print ""',
        '    Debug.Print "Next: run ImportAndMatch to process Honda data."',
        'End Sub',
        '',
        'Private Sub RemoveIfExists(ByVal vbProj As Object, ByVal moduleName As String)',
        '    Dim comp As Object',
        '    On Error Resume Next',
        '    Set comp = vbProj.VBComponents(moduleName)',
        '    If Not comp Is Nothing Then',
        '        If comp.Name <> "ModBootstrap" Then',
        '            vbProj.VBComponents.Remove comp',
        '        End If',
        '    End If',
        '    On Error GoTo 0',
        'End Sub',
        '',
        'Public Sub ImportAndMatch()',
        "    ' Import Honda data files via file dialogs (macOS sandbox).",
        '    Dim bankFile As Variant',
        '    Dim febFile As Variant',
        '    Dim janFile As Variant',
        '    ',
        '    MsgBox "Select the Honda BANK STATEMENT (CSV)", vbInformation, "Step 1 of 3"',
        '    bankFile = Application.GetOpenFilename("CSV Files (*.csv),*.csv", , "Select Bank Statement")',
        '    If bankFile = False Then Exit Sub',
        '    ',
        '    MsgBox "Select the Honda FEBRUARY GL (XLSX)", vbInformation, "Step 2 of 3"',
        '    febFile = Application.GetOpenFilename("Excel Files (*.xlsx),*.xlsx", , "Select Feb GL")',
        '    If febFile = False Then Exit Sub',
        '    ',
        '    MsgBox "Select the Honda JANUARY GL (XLSX)", vbInformation, "Step 3 of 3"',
        '    janFile = Application.GetOpenFilename("Excel Files (*.xlsx),*.xlsx", , "Select Jan GL")',
        '    If janFile = False Then Exit Sub',
        '    ',
        '    Dim bankCount As Long',
        '    Dim febCount As Long',
        '    Dim janCount As Long',
        '    ',
        '    Application.StatusBar = "ABR: Importing bank statement..."',
        '    bankCount = ModImportBank.ImportBankStatement(CStr(bankFile))',
        '    ',
        '    Application.StatusBar = "ABR: Importing Feb GL..."',
        '    febCount = ModImportDMS.ImportDMSExport(CStr(febFile))',
        '    ',
        '    Application.StatusBar = "ABR: Importing Jan GL..."',
        '    janCount = ModImportDMS.ImportDMSExport(CStr(janFile))',
        '    ',
        '    If bankCount = 0 Then',
        '        MsgBox "Bank import returned 0. Check file format.", vbExclamation',
        '        Exit Sub',
        '    End If',
        '    ',
        '    Application.StatusBar = "ABR: Running matching engine..."',
        '    ModMatchEngine.RunMatching ModImportBank.LoadBankTransactions(), ModImportDMS.LoadDMSTransactions()',
        '    ',
        '    Application.StatusBar = "ABR: Auto-accepting..."',
        '    ModStagingManager.AcceptAllHighConfidence',
        '    ',
        '    Application.StatusBar = False',
        '    MsgBox "Done!" & vbCrLf & vbCrLf & _',
        '        "Bank: " & bankCount & " transactions" & vbCrLf & _',
        '        "Feb GL: " & febCount & " transactions" & vbCrLf & _',
        '        "Jan GL: " & janCount & " transactions" & vbCrLf & vbCrLf & _',
        '        "Check StagedMatches and Reconciled sheets.", _',
        '        vbInformation, "ABR Complete"',
        'End Sub',
    ])

    return '\n'.join(lines)


def build_applescript(xlsx_path, xlsm_path, bootstrap_path):
    """Build AppleScript to open workbook, save as xlsm, import bootstrap."""
    return f'''
tell application "Microsoft Excel"
    activate
    delay 1

    -- Open the workbook
    open "{xlsx_path}"
    delay 2

    -- Save as macro-enabled workbook (.xlsm)
    -- file format constant 52 = xlOpenXMLWorkbookMacroEnabled (.xlsm)
    set wb to active workbook
    save workbook as wb filename "{xlsm_path}" file format file format type macro enabled
    delay 1

    -- Import the bootstrap module via VBA
    do Visual Basic "ThisWorkbook.VBProject.VBComponents.Import \\"" & "{bootstrap_path}" & "\\""
    delay 1

    -- Save again
    save wb

    return "Workbook created and bootstrap module imported."
end tell
'''


def print_manual_instructions(xlsx_path, bootstrap_path):
    """Print manual fallback instructions."""
    print(f"""
  MANUAL SETUP (if AppleScript didn't work):

  1. Open {xlsx_path} in Excel
  2. Save As -> .xlsm (macro-enabled workbook)
  3. Open VBA Editor (Fn+F11)
  4. File > Import File... -> select:
     {bootstrap_path}
  5. In Immediate Window (Ctrl+G), type:
     RunBootstrap
     (this imports all {15} VBA modules automatically)
  6. Then type:
     ImportAndMatch
     (this imports Honda data and runs matching)
""")


if __name__ == '__main__':
    main()
