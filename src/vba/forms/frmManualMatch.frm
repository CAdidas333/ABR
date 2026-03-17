VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmManualMatch
   Caption         =   "Manual Match"
   ClientHeight    =   7800
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10800
   OleObjectBlob   =   "frmManualMatch.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmManualMatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'===============================================================================
' frmManualMatch — Side-by-Side Manual Matching
'
' Shows unmatched bank transactions on the left and unmatched DMS transactions
' on the right. Controller selects one from each side and clicks Match.
'===============================================================================

Option Explicit

Private Sub UserForm_Initialize()
    LoadUnmatchedBank
    LoadUnmatchedDMS
End Sub

Private Sub LoadUnmatchedBank()
    lstBank.Clear
    lstBank.ColumnCount = 4
    lstBank.ColumnWidths = "40;70;150;80"

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("BankData")

    Dim lastRow As Long
    lastRow = ModHelpers.GetLastRow(ws, 1)

    Dim rowIdx As Long
    rowIdx = 0

    Dim i As Long
    For i = 2 To lastRow
        If ws.Cells(i, 10).Value = False Then  ' Not matched
            lstBank.AddItem ""
            lstBank.List(rowIdx, 0) = CStr(ws.Cells(i, 1).Value)   ' Row ID
            lstBank.List(rowIdx, 1) = Format(ws.Cells(i, 2).Value, "MM/DD/YY")
            lstBank.List(rowIdx, 2) = Left(CStr(ws.Cells(i, 4).Value), 40)
            lstBank.List(rowIdx, 3) = Format(ws.Cells(i, 5).Value, "$#,##0.00")
            rowIdx = rowIdx + 1
        End If
    Next i

    lblBankCount.Caption = rowIdx & " unmatched bank transactions"
End Sub

Private Sub LoadUnmatchedDMS()
    lstDMS.Clear
    lstDMS.ColumnCount = 4
    lstDMS.ColumnWidths = "40;70;150;80"

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("DMSData")

    Dim lastRow As Long
    lastRow = ModHelpers.GetLastRow(ws, 1)

    Dim rowIdx As Long
    rowIdx = 0

    Dim i As Long
    For i = 2 To lastRow
        If ws.Cells(i, 9).Value = False Then  ' Not matched
            lstDMS.AddItem ""
            lstDMS.List(rowIdx, 0) = CStr(ws.Cells(i, 1).Value)   ' Row ID
            lstDMS.List(rowIdx, 1) = Format(ws.Cells(i, 2).Value, "MM/DD/YY")
            lstDMS.List(rowIdx, 2) = Left(CStr(ws.Cells(i, 3).Value), 40)
            lstDMS.List(rowIdx, 3) = Format(ws.Cells(i, 5).Value, "$#,##0.00")
            rowIdx = rowIdx + 1
        End If
    Next i

    lblDMSCount.Caption = rowIdx & " unmatched DMS transactions"
End Sub

Private Sub btnMatch_Click()
    If lstBank.ListIndex < 0 Then
        MsgBox "Select a bank transaction.", vbExclamation
        Exit Sub
    End If
    If lstDMS.ListIndex < 0 Then
        MsgBox "Select a DMS transaction.", vbExclamation
        Exit Sub
    End If

    Dim bankID As Long
    bankID = CLng(lstBank.List(lstBank.ListIndex, 0))
    Dim dmsID As Long
    dmsID = CLng(lstDMS.List(lstDMS.ListIndex, 0))

    Dim matchID As Long
    matchID = ModStagingManager.CreateManualMatch(bankID, dmsID)

    If matchID > 0 Then
        MsgBox "Match created (ID: " & matchID & ")", vbInformation
        LoadUnmatchedBank
        LoadUnmatchedDMS
        ModMain.UpdateDashboardStats
    Else
        MsgBox "Could not create match.", vbExclamation
    End If
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

'===============================================================================
' Required controls:
'   ListBoxes: lstBank (left side), lstDMS (right side)
'   Labels:    lblBankCount, lblDMSCount
'   Buttons:   btnMatch, btnClose
'===============================================================================
