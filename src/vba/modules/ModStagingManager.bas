Attribute VB_Name = "ModStagingManager"
'===============================================================================
' ModStagingManager — Match Staging and Review Workflow
'
' All proposed matches flow through this module. No match is ever committed
' without going through Stage → Accept/Reject. This is the enforcement
' layer for the "never auto-commit" rule.
'===============================================================================

Option Explicit

Private Const STAGED_SHEET As String = "StagedMatches"
Private Const RECONCILED_SHEET As String = "Reconciled"

' StagedMatches column positions
Private Const COL_MATCH_ID As Long = 1
Private Const COL_MATCH_TYPE As Long = 2
Private Const COL_CONFIDENCE As Long = 3
Private Const COL_BANK_IDS As Long = 4
Private Const COL_BANK_DATE As Long = 5
Private Const COL_BANK_DESC As Long = 6
Private Const COL_BANK_AMOUNT As Long = 7
Private Const COL_DMS_IDS As Long = 8
Private Const COL_DMS_DATE As Long = 9
Private Const COL_DMS_DESC As Long = 10
Private Const COL_DMS_AMOUNT As Long = 11
Private Const COL_AMOUNT_DIFF As Long = 12
Private Const COL_DATE_DIFF As Long = 13
Private Const COL_CHECK_MATCH As Long = 14
Private Const COL_BREAKDOWN As Long = 15
Private Const COL_STATUS As Long = 16
Private Const COL_ACTION_TS As Long = 17
Private Const COL_ACTION_BY As Long = 18
Private Const COL_REJECT_REASON As Long = 19

' ---------------------------------------------------------------------------
' Stage a Match
' ---------------------------------------------------------------------------

Public Sub StageMatch(ByVal match As clsMatchResult)
    ' Write a proposed match to the StagedMatches sheet. Status = STAGED.
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(STAGED_SHEET)

    Dim nextRow As Long
    nextRow = ModHelpers.GetNextRow(ws, COL_MATCH_ID)

    ws.Cells(nextRow, COL_MATCH_ID).Value = match.MatchID
    ws.Cells(nextRow, COL_MATCH_TYPE).Value = match.MatchType
    ws.Cells(nextRow, COL_CONFIDENCE).Value = match.ConfidenceScore / 100#
    ws.Cells(nextRow, COL_CONFIDENCE).NumberFormat = "0.0%"
    ws.Cells(nextRow, COL_BANK_IDS).Value = match.BankTransactionIDs
    ws.Cells(nextRow, COL_BANK_DATE).Value = match.BankDate
    ws.Cells(nextRow, COL_BANK_DATE).NumberFormat = "MM/DD/YYYY"
    ws.Cells(nextRow, COL_BANK_DESC).Value = match.BankDescription
    ws.Cells(nextRow, COL_BANK_AMOUNT).Value = match.BankAmount
    ws.Cells(nextRow, COL_BANK_AMOUNT).NumberFormat = "#,##0.00"
    ws.Cells(nextRow, COL_DMS_IDS).Value = match.DMSTransactionIDs
    ws.Cells(nextRow, COL_DMS_DATE).Value = match.DMSDate
    ws.Cells(nextRow, COL_DMS_DATE).NumberFormat = "MM/DD/YYYY"
    ws.Cells(nextRow, COL_DMS_DESC).Value = match.DMSDescription
    ws.Cells(nextRow, COL_DMS_AMOUNT).Value = match.DMSAmount
    ws.Cells(nextRow, COL_DMS_AMOUNT).NumberFormat = "#,##0.00"
    ws.Cells(nextRow, COL_AMOUNT_DIFF).Value = match.AmountDifference
    ws.Cells(nextRow, COL_AMOUNT_DIFF).NumberFormat = "#,##0.00"
    ws.Cells(nextRow, COL_DATE_DIFF).Value = match.DateDifference
    ws.Cells(nextRow, COL_CHECK_MATCH).Value = match.CheckNumberMatch
    ws.Cells(nextRow, COL_BREAKDOWN).Value = match.ScoreBreakdown
    ws.Cells(nextRow, COL_STATUS).Value = "STAGED"
    ws.Cells(nextRow, COL_ACTION_TS).Value = Now
    ws.Cells(nextRow, COL_ACTION_TS).NumberFormat = "MM/DD/YYYY HH:MM:SS"
End Sub

' ---------------------------------------------------------------------------
' Accept / Reject Matches
' ---------------------------------------------------------------------------

Public Sub AcceptMatch(ByVal matchID As Long)
    ' Accept a staged match: move from StagedMatches to Reconciled.
    Dim wsStaged As Worksheet, wsRecon As Worksheet
    Set wsStaged = ThisWorkbook.Sheets(STAGED_SHEET)
    Set wsRecon = ThisWorkbook.Sheets(RECONCILED_SHEET)

    Dim matchRow As Long
    matchRow = FindMatchRow(wsStaged, matchID)
    If matchRow = 0 Then Exit Sub

    ' Copy to Reconciled
    Dim nextReconRow As Long
    nextReconRow = ModHelpers.GetNextRow(wsRecon, 1)

    ' Reconciled columns: Match ID, Type, Confidence, Bank IDs, Bank Date,
    ' Bank Desc, Bank Amount, DMS IDs, DMS Date, DMS Desc, DMS Amount,
    ' Reconciled Timestamp, Reconciled By
    wsRecon.Cells(nextReconRow, 1).Value = wsStaged.Cells(matchRow, COL_MATCH_ID).Value
    wsRecon.Cells(nextReconRow, 2).Value = wsStaged.Cells(matchRow, COL_MATCH_TYPE).Value
    wsRecon.Cells(nextReconRow, 3).Value = wsStaged.Cells(matchRow, COL_CONFIDENCE).Value
    wsRecon.Cells(nextReconRow, 3).NumberFormat = "0.0%"
    wsRecon.Cells(nextReconRow, 4).Value = wsStaged.Cells(matchRow, COL_BANK_IDS).Value
    wsRecon.Cells(nextReconRow, 5).Value = wsStaged.Cells(matchRow, COL_BANK_DATE).Value
    wsRecon.Cells(nextReconRow, 5).NumberFormat = "MM/DD/YYYY"
    wsRecon.Cells(nextReconRow, 6).Value = wsStaged.Cells(matchRow, COL_BANK_DESC).Value
    wsRecon.Cells(nextReconRow, 7).Value = wsStaged.Cells(matchRow, COL_BANK_AMOUNT).Value
    wsRecon.Cells(nextReconRow, 7).NumberFormat = "#,##0.00"
    wsRecon.Cells(nextReconRow, 8).Value = wsStaged.Cells(matchRow, COL_DMS_IDS).Value
    wsRecon.Cells(nextReconRow, 9).Value = wsStaged.Cells(matchRow, COL_DMS_DATE).Value
    wsRecon.Cells(nextReconRow, 9).NumberFormat = "MM/DD/YYYY"
    wsRecon.Cells(nextReconRow, 10).Value = wsStaged.Cells(matchRow, COL_DMS_DESC).Value
    wsRecon.Cells(nextReconRow, 11).Value = wsStaged.Cells(matchRow, COL_DMS_AMOUNT).Value
    wsRecon.Cells(nextReconRow, 11).NumberFormat = "#,##0.00"
    wsRecon.Cells(nextReconRow, 12).Value = Now
    wsRecon.Cells(nextReconRow, 12).NumberFormat = "MM/DD/YYYY HH:MM:SS"
    wsRecon.Cells(nextReconRow, 13).Value = ModHelpers.GetCurrentUserName()

    ' Update staged status
    wsStaged.Cells(matchRow, COL_STATUS).Value = "ACCEPTED"
    wsStaged.Cells(matchRow, COL_ACTION_TS).Value = Now
    wsStaged.Cells(matchRow, COL_ACTION_BY).Value = ModHelpers.GetCurrentUserName()

    ' Log
    ModAuditTrail.LogMatchAccepted matchID
End Sub

Public Sub RejectMatch(ByVal matchID As Long, Optional ByVal reason As String = "")
    ' Reject a staged match: update status and clear match flags on source data.
    Dim wsStaged As Worksheet
    Set wsStaged = ThisWorkbook.Sheets(STAGED_SHEET)

    Dim matchRow As Long
    matchRow = FindMatchRow(wsStaged, matchID)
    If matchRow = 0 Then Exit Sub

    ' Update staged status
    wsStaged.Cells(matchRow, COL_STATUS).Value = "REJECTED"
    wsStaged.Cells(matchRow, COL_ACTION_TS).Value = Now
    wsStaged.Cells(matchRow, COL_ACTION_BY).Value = ModHelpers.GetCurrentUserName()
    If reason <> "" Then
        wsStaged.Cells(matchRow, COL_REJECT_REASON).Value = reason
    End If

    ' Clear match flags on source sheets
    Dim bankIDs As String
    bankIDs = CStr(wsStaged.Cells(matchRow, COL_BANK_IDS).Value)
    Dim dmsIDs As String
    dmsIDs = CStr(wsStaged.Cells(matchRow, COL_DMS_IDS).Value)

    ' Clear bank match status
    Dim bankIDArr() As String
    bankIDArr = Split(bankIDs, ",")
    Dim i As Long
    For i = LBound(bankIDArr) To UBound(bankIDArr)
        If Trim(bankIDArr(i)) <> "" Then
            ModImportBank.ClearBankMatchStatus CLng(Trim(bankIDArr(i)))
        End If
    Next i

    ' Clear DMS match status
    Dim dmsIDArr() As String
    dmsIDArr = Split(dmsIDs, ",")
    For i = LBound(dmsIDArr) To UBound(dmsIDArr)
        If Trim(dmsIDArr(i)) <> "" Then
            ModImportDMS.ClearDMSMatchStatus CLng(Trim(dmsIDArr(i)))
        End If
    Next i

    ' Log
    ModAuditTrail.LogMatchRejected matchID, reason
End Sub

' ---------------------------------------------------------------------------
' Bulk Operations
' ---------------------------------------------------------------------------

Public Sub AcceptAllHighConfidence()
    ' Accept all staged matches with confidence >= high threshold.
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(STAGED_SHEET)

    Dim threshold As Double
    threshold = ModConfig.HighConfidenceThreshold() / 100#

    Dim lastRow As Long
    lastRow = ModHelpers.GetLastRow(ws, COL_MATCH_ID)

    Dim acceptCount As Long
    acceptCount = 0

    Dim i As Long
    For i = 2 To lastRow
        If ws.Cells(i, COL_STATUS).Value = "STAGED" Then
            If CDbl(ws.Cells(i, COL_CONFIDENCE).Value) >= threshold Then
                AcceptMatch CLng(ws.Cells(i, COL_MATCH_ID).Value)
                acceptCount = acceptCount + 1
            End If
        End If
    Next i

    MsgBox acceptCount & " high-confidence matches accepted.", _
           vbInformation, "Bulk Accept Complete"
End Sub

Public Sub AcceptSelected(ByVal matchIDs As String)
    ' Accept a comma-separated list of match IDs.
    Dim ids() As String
    ids = Split(matchIDs, ",")

    Dim i As Long
    For i = LBound(ids) To UBound(ids)
        If Trim(ids(i)) <> "" Then
            AcceptMatch CLng(Trim(ids(i)))
        End If
    Next i
End Sub

' ---------------------------------------------------------------------------
' Manual Match
' ---------------------------------------------------------------------------

Public Function CreateManualMatch(ByVal bankTxnID As Long, _
                                   ByVal dmsTxnID As Long) As Long
    ' Create a manual match between a bank and DMS transaction.
    ' Returns the new match ID.

    Dim bankTxns As Collection
    Set bankTxns = ModImportBank.LoadBankTransactions()

    Dim dmsTxns As Collection
    Set dmsTxns = ModImportDMS.LoadDMSTransactions()

    ' Find the transactions
    Dim bankTxn As clsTransaction, dmsTxn As clsTransaction
    Dim found As Boolean

    found = False
    Dim i As Long
    For i = 1 To bankTxns.Count
        Set bankTxn = bankTxns(i)
        If bankTxn.TransactionID = bankTxnID Then
            found = True
            Exit For
        End If
    Next i
    If Not found Then
        CreateManualMatch = 0
        Exit Function
    End If

    found = False
    For i = 1 To dmsTxns.Count
        Set dmsTxn = dmsTxns(i)
        If dmsTxn.TransactionID = dmsTxnID Then
            found = True
            Exit For
        End If
    Next i
    If Not found Then
        CreateManualMatch = 0
        Exit Function
    End If

    ' Create match result
    Dim match As New clsMatchResult
    match.MatchID = GetNextMatchID()
    match.BankTransactionIDs = CStr(bankTxnID)
    match.DMSTransactionIDs = CStr(dmsTxnID)
    match.ConfidenceScore = 100  ' Manual = 100% by definition
    match.MatchType = "MANUAL"
    match.BankAmount = bankTxn.Amount
    match.DMSAmount = dmsTxn.Amount
    match.BankDescription = bankTxn.Description
    match.DMSDescription = dmsTxn.Description
    match.BankDate = bankTxn.TransactionDate
    match.DMSDate = dmsTxn.TransactionDate
    match.AmountDifference = bankTxn.Amount - dmsTxn.Amount
    match.DateDifference = ModHelpers.DateDiffDays(bankTxn.TransactionDate, _
                                                    dmsTxn.TransactionDate)
    match.ScoreBreakdown = "Manual match by controller"
    match.Status = "ACCEPTED"  ' Manual matches are immediately accepted

    ' Write directly to Reconciled (skip staging for manual matches)
    Dim wsRecon As Worksheet
    Set wsRecon = ThisWorkbook.Sheets(RECONCILED_SHEET)

    Dim nextRow As Long
    nextRow = ModHelpers.GetNextRow(wsRecon, 1)

    wsRecon.Cells(nextRow, 1).Value = match.MatchID
    wsRecon.Cells(nextRow, 2).Value = "MANUAL"
    wsRecon.Cells(nextRow, 3).Value = 1#  ' 100%
    wsRecon.Cells(nextRow, 3).NumberFormat = "0.0%"
    wsRecon.Cells(nextRow, 4).Value = CStr(bankTxnID)
    wsRecon.Cells(nextRow, 5).Value = bankTxn.TransactionDate
    wsRecon.Cells(nextRow, 5).NumberFormat = "MM/DD/YYYY"
    wsRecon.Cells(nextRow, 6).Value = bankTxn.Description
    wsRecon.Cells(nextRow, 7).Value = bankTxn.Amount
    wsRecon.Cells(nextRow, 7).NumberFormat = "#,##0.00"
    wsRecon.Cells(nextRow, 8).Value = CStr(dmsTxnID)
    wsRecon.Cells(nextRow, 9).Value = dmsTxn.TransactionDate
    wsRecon.Cells(nextRow, 9).NumberFormat = "MM/DD/YYYY"
    wsRecon.Cells(nextRow, 10).Value = dmsTxn.Description
    wsRecon.Cells(nextRow, 11).Value = dmsTxn.Amount
    wsRecon.Cells(nextRow, 11).NumberFormat = "#,##0.00"
    wsRecon.Cells(nextRow, 12).Value = Now
    wsRecon.Cells(nextRow, 12).NumberFormat = "MM/DD/YYYY HH:MM:SS"
    wsRecon.Cells(nextRow, 13).Value = ModHelpers.GetCurrentUserName()

    ' Update source sheets
    ModImportBank.UpdateBankMatchStatus bankTxnID, match.MatchID, "MANUAL", 100
    ModImportDMS.UpdateDMSMatchStatus dmsTxnID, match.MatchID, "MANUAL", 100

    ' Log
    ModAuditTrail.LogManualMatch match.MatchID, bankTxn.Description, dmsTxn.Description

    CreateManualMatch = match.MatchID
End Function

' ---------------------------------------------------------------------------
' Utility Functions
' ---------------------------------------------------------------------------

Public Function GetNextMatchID() As Long
    ' Get the next available match ID.
    Dim wsStaged As Worksheet, wsRecon As Worksheet
    Set wsStaged = ThisWorkbook.Sheets(STAGED_SHEET)
    Set wsRecon = ThisWorkbook.Sheets(RECONCILED_SHEET)

    Dim maxID As Long
    maxID = 0

    ' Check staged
    Dim lastRow As Long
    lastRow = ModHelpers.GetLastRow(wsStaged, COL_MATCH_ID)
    If lastRow > 1 Then
        Dim i As Long
        For i = 2 To lastRow
            If Not IsEmpty(wsStaged.Cells(i, COL_MATCH_ID).Value) Then
                If CLng(wsStaged.Cells(i, COL_MATCH_ID).Value) > maxID Then
                    maxID = CLng(wsStaged.Cells(i, COL_MATCH_ID).Value)
                End If
            End If
        Next i
    End If

    ' Check reconciled
    lastRow = ModHelpers.GetLastRow(wsRecon, 1)
    If lastRow > 1 Then
        For i = 2 To lastRow
            If Not IsEmpty(wsRecon.Cells(i, 1).Value) Then
                If CLng(wsRecon.Cells(i, 1).Value) > maxID Then
                    maxID = CLng(wsRecon.Cells(i, 1).Value)
                End If
            End If
        Next i
    End If

    GetNextMatchID = maxID + 1
End Function

Private Function FindMatchRow(ByVal ws As Worksheet, ByVal matchID As Long) As Long
    ' Find the row for a given match ID.
    Dim lastRow As Long
    lastRow = ModHelpers.GetLastRow(ws, COL_MATCH_ID)

    Dim i As Long
    For i = 2 To lastRow
        If Not IsEmpty(ws.Cells(i, COL_MATCH_ID).Value) Then
            If CLng(ws.Cells(i, COL_MATCH_ID).Value) = matchID Then
                FindMatchRow = i
                Exit Function
            End If
        End If
    Next i

    FindMatchRow = 0
End Function

' ---------------------------------------------------------------------------
' Statistics
' ---------------------------------------------------------------------------

Public Function GetStagedCount() As Long
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(STAGED_SHEET)
    Dim lastRow As Long
    lastRow = ModHelpers.GetLastRow(ws, COL_MATCH_ID)
    Dim count As Long
    count = 0
    Dim i As Long
    For i = 2 To lastRow
        If ws.Cells(i, COL_STATUS).Value = "STAGED" Then
            count = count + 1
        End If
    Next i
    GetStagedCount = count
End Function

Public Function GetAcceptedCount() As Long
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(RECONCILED_SHEET)
    GetAcceptedCount = WorksheetFunction.Max(0, ModHelpers.GetLastRow(ws, 1) - 1)
End Function
