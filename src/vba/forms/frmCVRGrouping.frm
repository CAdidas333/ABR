VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCVRGrouping
   Caption         =   "CVR Group Matching"
   ClientHeight    =   7200
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8400
   OleObjectBlob   =   "frmCVRGrouping.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmCVRGrouping"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'===============================================================================
' frmCVRGrouping — CVR Many-to-One Group Selection
'
' Shows a target DMS transaction at the top and a list of candidate bank
' transactions below with checkboxes. Running total shows sum vs target
' with a visual indicator (green = match, red = mismatch).
'===============================================================================

Option Explicit

Private mTargetDMSTxnID As Long
Private mTargetAmount As Currency

Public Sub SetTarget(ByVal dmsTxnID As Long, ByVal targetAmount As Currency, _
                     ByVal description As String)
    mTargetDMSTxnID = dmsTxnID
    mTargetAmount = targetAmount

    lblTargetDesc.Caption = description
    lblTargetAmount.Caption = Format(targetAmount, "$#,##0.00")

    LoadCandidates
End Sub

Private Sub LoadCandidates()
    ' Load unmatched bank transactions that could be fragments
    lstCandidates.Clear
    lstCandidates.ColumnCount = 4
    lstCandidates.ColumnWidths = "40;70;150;80"
    lstCandidates.MultiSelect = fmMultiSelectMulti

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("BankData")

    Dim lastRow As Long
    lastRow = ModHelpers.GetLastRow(ws, 1)

    Dim rowIdx As Long
    rowIdx = 0

    Dim i As Long
    For i = 2 To lastRow
        If ws.Cells(i, 10).Value = False Then  ' Not matched
            Dim amt As Currency
            amt = CCur(ws.Cells(i, 5).Value)

            ' Same sign and smaller than target
            If (amt > 0) = (mTargetAmount > 0) And Abs(amt) < Abs(mTargetAmount) Then
                lstCandidates.AddItem ""
                lstCandidates.List(rowIdx, 0) = CStr(ws.Cells(i, 1).Value)
                lstCandidates.List(rowIdx, 1) = Format(ws.Cells(i, 2).Value, "MM/DD/YY")
                lstCandidates.List(rowIdx, 2) = Left(CStr(ws.Cells(i, 4).Value), 40)
                lstCandidates.List(rowIdx, 3) = Format(amt, "$#,##0.00")
                rowIdx = rowIdx + 1
            End If
        End If
    Next i
End Sub

Private Sub lstCandidates_Click()
    ' Update running total when selection changes
    Dim total As Currency
    total = 0

    Dim i As Long
    For i = 0 To lstCandidates.ListCount - 1
        If lstCandidates.Selected(i) Then
            total = total + ModHelpers.NormalizeCurrency( _
                Replace(Replace(lstCandidates.List(i, 3), "$", ""), ",", ""))
        End If
    Next i

    lblRunningTotal.Caption = Format(total, "$#,##0.00")
    lblVariance.Caption = Format(total - mTargetAmount, "$#,##0.00")

    ' Color indicator
    If Abs(total - mTargetAmount) <= ModConfig.CVRTolerance() Then
        lblVariance.ForeColor = RGB(39, 118, 39)   ' Green
        lblMatchStatus.Caption = "MATCH"
        lblMatchStatus.ForeColor = RGB(39, 118, 39)
    Else
        lblVariance.ForeColor = RGB(192, 0, 0)      ' Red
        lblMatchStatus.Caption = "NO MATCH"
        lblMatchStatus.ForeColor = RGB(192, 0, 0)
    End If
End Sub

Private Sub btnPropose_Click()
    ' Stage the selected group as a CVR match
    Dim total As Currency
    total = 0
    Dim bankIDs As String
    bankIDs = ""
    Dim count As Long
    count = 0

    Dim i As Long
    For i = 0 To lstCandidates.ListCount - 1
        If lstCandidates.Selected(i) Then
            If bankIDs <> "" Then bankIDs = bankIDs & ","
            bankIDs = bankIDs & lstCandidates.List(i, 0)
            total = total + ModHelpers.NormalizeCurrency( _
                Replace(Replace(lstCandidates.List(i, 3), "$", ""), ",", ""))
            count = count + 1
        End If
    Next i

    If count < 2 Then
        MsgBox "Select at least 2 transactions to group.", vbExclamation
        Exit Sub
    End If

    If Abs(total - mTargetAmount) > ModConfig.CVRTolerance() Then
        Dim resp As VbMsgBoxResult
        resp = MsgBox("Selected items do not sum to the target amount." & vbCrLf & _
                      "Variance: " & Format(total - mTargetAmount, "$#,##0.00") & vbCrLf & _
                      "Propose anyway?", vbYesNo + vbExclamation, "Variance Warning")
        If resp = vbNo Then Exit Sub
    End If

    ' Create CVR match result
    Dim result As New clsMatchResult
    result.MatchID = ModStagingManager.GetNextMatchID()
    result.MatchType = "MANY_TO_ONE_BANK"
    result.BankTransactionIDs = bankIDs
    result.DMSTransactionIDs = CStr(mTargetDMSTxnID)
    result.BankAmount = total
    result.DMSAmount = mTargetAmount
    result.AmountDifference = total - mTargetAmount
    result.ConfidenceScore = 100  ' Manual grouping = 100%
    result.BankDescription = "Manual CVR group (" & count & " items)"
    result.DMSDescription = lblTargetDesc.Caption
    result.ScoreBreakdown = "Manual CVR grouping by controller"

    ModStagingManager.StageMatch result
    ModAuditTrail.LogCVRGroup result.MatchID, count, total

    MsgBox "CVR group match staged (ID: " & result.MatchID & ")", vbInformation
    Unload Me
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

'===============================================================================
' Required controls:
'   Labels:   lblTargetDesc, lblTargetAmount, lblRunningTotal, lblVariance,
'             lblMatchStatus
'   ListBox:  lstCandidates (MultiSelect = fmMultiSelectMulti)
'   Buttons:  btnPropose, btnClose
'===============================================================================
