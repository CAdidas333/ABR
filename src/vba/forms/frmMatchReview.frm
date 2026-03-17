VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMatchReview
   Caption         =   "Review Staged Matches"
   ClientHeight    =   8400
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10800
   OleObjectBlob   =   "frmMatchReview.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmMatchReview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'===============================================================================
' frmMatchReview — Staged Match Review UserForm
'
' Displays staged matches in a ListView for controller review.
' Accept/Reject per row or in bulk. Color-coded by confidence band.
'===============================================================================

Option Explicit

Private Sub UserForm_Initialize()
    LoadStagedMatches
    UpdateCounts
End Sub

Private Sub LoadStagedMatches()
    ' Load staged matches into the ListView
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("StagedMatches")

    Dim lastRow As Long
    lastRow = ModHelpers.GetLastRow(ws, 1)

    lstMatches.Clear
    lstMatches.ColumnCount = 8
    lstMatches.ColumnWidths = "40;80;50;80;100;80;100;80"

    Dim i As Long
    Dim rowIdx As Long
    rowIdx = 0

    For i = 2 To lastRow
        If ws.Cells(i, 16).Value = "STAGED" Then  ' Status = STAGED
            lstMatches.AddItem ""
            lstMatches.List(rowIdx, 0) = CStr(ws.Cells(i, 1).Value)   ' Match ID
            lstMatches.List(rowIdx, 1) = CStr(ws.Cells(i, 2).Value)   ' Match Type
            lstMatches.List(rowIdx, 2) = Format(ws.Cells(i, 3).Value * 100, "0.0") & "%"  ' Confidence
            lstMatches.List(rowIdx, 3) = Format(ws.Cells(i, 5).Value, "MM/DD/YY")  ' Bank Date
            lstMatches.List(rowIdx, 4) = Left(CStr(ws.Cells(i, 6).Value), 30)  ' Bank Desc
            lstMatches.List(rowIdx, 5) = Format(ws.Cells(i, 7).Value, "$#,##0.00")  ' Bank Amt
            lstMatches.List(rowIdx, 6) = Left(CStr(ws.Cells(i, 10).Value), 30) ' DMS Desc
            lstMatches.List(rowIdx, 7) = Format(ws.Cells(i, 11).Value, "$#,##0.00") ' DMS Amt
            rowIdx = rowIdx + 1
        End If
    Next i
End Sub

Private Sub UpdateCounts()
    Dim stagedCount As Long
    stagedCount = ModStagingManager.GetStagedCount()
    Dim acceptedCount As Long
    acceptedCount = ModStagingManager.GetAcceptedCount()

    lblStatus.Caption = stagedCount & " matches pending review | " & _
                        acceptedCount & " reconciled"
End Sub

Private Sub btnAcceptSelected_Click()
    ' Accept the currently selected match
    If lstMatches.ListIndex < 0 Then
        MsgBox "Please select a match to accept.", vbExclamation
        Exit Sub
    End If

    Dim matchID As Long
    matchID = CLng(lstMatches.List(lstMatches.ListIndex, 0))

    ModStagingManager.AcceptMatch matchID

    LoadStagedMatches
    UpdateCounts
    ModMain.UpdateDashboardStats
End Sub

Private Sub btnRejectSelected_Click()
    ' Reject the currently selected match
    If lstMatches.ListIndex < 0 Then
        MsgBox "Please select a match to reject.", vbExclamation
        Exit Sub
    End If

    Dim reason As String
    reason = InputBox("Rejection reason (optional):", "Reject Match")

    Dim matchID As Long
    matchID = CLng(lstMatches.List(lstMatches.ListIndex, 0))

    ModStagingManager.RejectMatch matchID, reason

    LoadStagedMatches
    UpdateCounts
    ModMain.UpdateDashboardStats
End Sub

Private Sub btnAcceptAllHigh_Click()
    ModStagingManager.AcceptAllHighConfidence
    LoadStagedMatches
    UpdateCounts
    ModMain.UpdateDashboardStats
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

'===============================================================================
' Required controls:
'   ListBox:  lstMatches (MultiSelect = fmMultiSelectSingle)
'   Labels:   lblStatus
'   Buttons:  btnAcceptSelected, btnRejectSelected, btnAcceptAllHigh, btnClose
'===============================================================================
