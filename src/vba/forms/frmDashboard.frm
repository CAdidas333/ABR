VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDashboard
   Caption         =   "ABR - Auto Bank Reconciliation"
   ClientHeight    =   7200
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6000
   OleObjectBlob   =   "frmDashboard.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmDashboard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'===============================================================================
' frmDashboard — Main Menu UserForm
'
' 5-step workflow with buttons and status indicators.
' This is the primary entry point for controllers.
'===============================================================================

Option Explicit

Private Sub UserForm_Initialize()
    Me.Caption = "ABR - " & ModConfig.LocationName()

    ' Update button states based on data availability
    Dim wsBank As Worksheet, wsDMS As Worksheet
    Set wsBank = ThisWorkbook.Sheets("BankData")
    Set wsDMS = ThisWorkbook.Sheets("DMSData")

    Dim bankCount As Long, dmsCount As Long
    bankCount = WorksheetFunction.Max(0, ModHelpers.GetLastRow(wsBank, 1) - 1)
    dmsCount = WorksheetFunction.Max(0, ModHelpers.GetLastRow(wsDMS, 1) - 1)

    lblBankCount.Caption = bankCount & " bank transactions loaded"
    lblDMSCount.Caption = dmsCount & " DMS transactions loaded"
    lblStagedCount.Caption = ModStagingManager.GetStagedCount() & " matches staged"
    lblReconCount.Caption = ModStagingManager.GetAcceptedCount() & " matches reconciled"
End Sub

Private Sub btnImportBank_Click()
    Me.Hide
    ModMain.Step1_ImportBankStatement
    Me.Show
    UserForm_Initialize  ' Refresh counts
End Sub

Private Sub btnImportDMS_Click()
    Me.Hide
    ModMain.Step2_ImportDMSData
    Me.Show
    UserForm_Initialize
End Sub

Private Sub btnRunMatching_Click()
    Me.Hide
    ModMain.Step3_RunAutoMatching
    Me.Show
    UserForm_Initialize
End Sub

Private Sub btnReviewMatches_Click()
    Me.Hide
    ModMain.Step4_ReviewMatches
    ' Don't re-show — user is now on the StagedMatches sheet
    Unload Me
End Sub

Private Sub btnFinalize_Click()
    Me.Hide
    ModMain.Step5_FinalizeAndExport
    Me.Show
    UserForm_Initialize
End Sub

Private Sub btnAcceptHighConf_Click()
    Me.Hide
    ModMain.AcceptAllHighConfidence
    Me.Show
    UserForm_Initialize
End Sub

Private Sub btnManualMatch_Click()
    Me.Hide
    ModMain.CreateManualMatchUI
    Me.Show
    UserForm_Initialize
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

'===============================================================================
' NOTE: This form requires controls to be created in the VBA Editor.
'
' Required controls (create via UserForm designer):
'   Labels:  lblBankCount, lblDMSCount, lblStagedCount, lblReconCount
'   Buttons: btnImportBank, btnImportDMS, btnRunMatching, btnReviewMatches,
'            btnFinalize, btnAcceptHighConf, btnManualMatch, btnClose
'
' Layout suggestion:
'   Row 1: Title label "ABR - Auto Bank Reconciliation"
'   Row 2: Location label (set in Initialize)
'   Row 3-4: Step 1 button + bank count label
'   Row 5-6: Step 2 button + DMS count label
'   Row 7-8: Step 3 button (Run Matching)
'   Row 9-10: Step 4 button (Review) + staged count label
'   Row 11-12: Step 5 button (Finalize) + recon count label
'   Row 13: Accept High Confidence button
'   Row 14: Manual Match button
'   Row 15: Close button
'===============================================================================
