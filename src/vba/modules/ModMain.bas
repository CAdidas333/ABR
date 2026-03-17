Attribute VB_Name = "ModMain"
'===============================================================================
' ModMain — Main Orchestrator
'
' Entry point for the ABR tool. Manages the 5-step reconciliation workflow
' and coordinates all modules.
'
' Workflow:
'   Step 1: Import Bank Statement
'   Step 2: Import DMS Data
'   Step 3: Run Auto-Matching (suggestion, NOT commitment)
'   Step 4: Review & Confirm Matches
'   Step 5: Finalize & Export
'===============================================================================

Option Explicit

' ---------------------------------------------------------------------------
' Step 1: Import Bank Statement
' ---------------------------------------------------------------------------

Public Sub Step1_ImportBankStatement()
    On Error GoTo HandleError

    ModAuditTrail.StartSession

    Dim txnCount As Long
    txnCount = ModImportBank.ImportBankStatement()

    If txnCount > 0 Then
        UpdateDashboardStatus 1, "COMPLETE", txnCount & " transactions imported"
        MsgBox txnCount & " bank transactions imported successfully.", _
               vbInformation, "Import Complete"
    Else
        MsgBox "No transactions were imported.", vbExclamation, "Import"
    End If

    UpdateDashboardStats
    Exit Sub

HandleError:
    MsgBox "Error importing bank statement:" & vbCrLf & Err.Description, _
           vbCritical, "Import Error"
End Sub

' ---------------------------------------------------------------------------
' Step 2: Import DMS Data
' ---------------------------------------------------------------------------

Public Sub Step2_ImportDMSData()
    On Error GoTo HandleError

    Dim txnCount As Long
    txnCount = ModImportDMS.ImportDMSExport()

    If txnCount > 0 Then
        UpdateDashboardStatus 2, "COMPLETE", txnCount & " transactions imported"
        MsgBox txnCount & " DMS transactions imported successfully.", _
               vbInformation, "Import Complete"
    Else
        MsgBox "No transactions were imported.", vbExclamation, "Import"
    End If

    UpdateDashboardStats
    Exit Sub

HandleError:
    MsgBox "Error importing DMS data:" & vbCrLf & Err.Description, _
           vbCritical, "Import Error"
End Sub

' ---------------------------------------------------------------------------
' Step 3: Run Auto-Matching
' ---------------------------------------------------------------------------

Public Sub Step3_RunAutoMatching()
    On Error GoTo HandleError

    ' Verify data exists
    Dim bankCount As Long, dmsCount As Long
    Dim wsBank As Worksheet, wsDMS As Worksheet
    Set wsBank = ThisWorkbook.Sheets("BankData")
    Set wsDMS = ThisWorkbook.Sheets("DMSData")

    bankCount = WorksheetFunction.Max(0, ModHelpers.GetLastRow(wsBank, 1) - 1)
    dmsCount = WorksheetFunction.Max(0, ModHelpers.GetLastRow(wsDMS, 1) - 1)

    If bankCount = 0 Or dmsCount = 0 Then
        MsgBox "Please import both bank statement and DMS data before running matching.", _
               vbExclamation, "Missing Data"
        Exit Sub
    End If

    Dim response As VbMsgBoxResult
    response = MsgBox("Run auto-matching on " & bankCount & " bank and " & _
                      dmsCount & " DMS transactions?" & vbCrLf & vbCrLf & _
                      "All matches will be STAGED for your review — nothing is " & _
                      "committed automatically.", _
                      vbYesNo + vbQuestion, "Run Auto-Matching")
    If response = vbNo Then Exit Sub

    Application.ScreenUpdating = False

    ' Load transactions
    Dim bankTxns As Collection
    Set bankTxns = ModImportBank.LoadBankTransactions()

    Dim dmsTxns As Collection
    Set dmsTxns = ModImportDMS.LoadDMSTransactions()

    ' Phase 1: 1:1 matching
    Application.StatusBar = "ABR: Phase 1 — Running 1:1 matching..."
    ModMatchEngine.RunMatching bankTxns, dmsTxns

    ' Reload to get updated match status
    Set bankTxns = ModImportBank.LoadBankTransactions()
    Set dmsTxns = ModImportDMS.LoadDMSTransactions()

    ' Get unmatched collections
    Dim unmatchedBank As New Collection
    Dim unmatchedDMS As New Collection

    Dim txn As clsTransaction
    Dim i As Long
    For i = 1 To bankTxns.Count
        Set txn = bankTxns(i)
        If Not txn.IsMatched Then unmatchedBank.Add txn
    Next i
    For i = 1 To dmsTxns.Count
        Set txn = dmsTxns(i)
        If Not txn.IsMatched Then unmatchedDMS.Add txn
    Next i

    ' Phase 2: CVR many-to-one matching
    Application.StatusBar = "ABR: Phase 2 — Running CVR matching..."
    ModMatchCVR.RunCVRMatching unmatchedBank, unmatchedDMS

    ' Phase 3: Reverse split matching
    Application.StatusBar = "ABR: Phase 3 — Running reverse split matching..."
    ModMatchCVR.RunReverseSplitMatching unmatchedBank, unmatchedDMS

    Application.ScreenUpdating = True
    Application.StatusBar = False

    ' Update dashboard
    UpdateDashboardStatus 3, "COMPLETE", "Matching complete"
    UpdateDashboardStats

    ' Show results
    Dim stagedCount As Long
    stagedCount = ModStagingManager.GetStagedCount()

    MsgBox "Auto-matching complete!" & vbCrLf & vbCrLf & _
           "Matches staged for review: " & stagedCount & vbCrLf & vbCrLf & _
           "Please go to Step 4 to review and confirm matches." & vbCrLf & _
           "(No matches have been committed — they all need your approval.)", _
           vbInformation, "Matching Complete"

    ' Navigate to StagedMatches sheet
    ThisWorkbook.Sheets("StagedMatches").Activate
    Exit Sub

HandleError:
    Application.ScreenUpdating = True
    Application.StatusBar = False
    MsgBox "Error during auto-matching:" & vbCrLf & Err.Description, _
           vbCritical, "Matching Error"
End Sub

' ---------------------------------------------------------------------------
' Step 4: Review & Confirm Matches
' ---------------------------------------------------------------------------

Public Sub Step4_ReviewMatches()
    ' Navigate to the StagedMatches sheet for review.
    ' The controller reviews staged matches and accepts/rejects each one.
    ThisWorkbook.Sheets("StagedMatches").Activate

    Dim stagedCount As Long
    stagedCount = ModStagingManager.GetStagedCount()

    If stagedCount = 0 Then
        MsgBox "No staged matches to review." & vbCrLf & _
               "Run auto-matching first (Step 3) or create manual matches.", _
               vbInformation, "Review"
    Else
        MsgBox stagedCount & " matches awaiting your review." & vbCrLf & vbCrLf & _
               "Available actions:" & vbCrLf & _
               "  - Accept All High Confidence (green rows)" & vbCrLf & _
               "  - Accept individual matches" & vbCrLf & _
               "  - Reject matches with optional reason" & vbCrLf & _
               "  - Create manual matches for unmatched items", _
               vbInformation, "Review Matches"
    End If
End Sub

Public Sub AcceptSelectedMatches()
    ' Accept all currently selected match rows on the StagedMatches sheet.
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("StagedMatches")

    If ActiveSheet.Name <> "StagedMatches" Then
        MsgBox "Please select match rows on the StagedMatches sheet.", _
               vbExclamation, "Wrong Sheet"
        Exit Sub
    End If

    Dim sel As Range
    Set sel = Selection

    Dim acceptCount As Long
    acceptCount = 0

    Dim row As Range
    For Each row In sel.Rows
        Dim matchRow As Long
        matchRow = row.row

        If matchRow >= 2 Then  ' Skip header
            If ws.Cells(matchRow, 16).Value = "STAGED" Then  ' Status column
                Dim matchID As Long
                matchID = CLng(ws.Cells(matchRow, 1).Value)
                ModStagingManager.AcceptMatch matchID
                acceptCount = acceptCount + 1
            End If
        End If
    Next row

    If acceptCount > 0 Then
        UpdateDashboardStats
        MsgBox acceptCount & " matches accepted.", vbInformation, "Accept Complete"
    End If
End Sub

Public Sub RejectSelectedMatches()
    ' Reject all currently selected match rows.
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("StagedMatches")

    If ActiveSheet.Name <> "StagedMatches" Then
        MsgBox "Please select match rows on the StagedMatches sheet.", _
               vbExclamation, "Wrong Sheet"
        Exit Sub
    End If

    Dim reason As String
    reason = InputBox("Enter rejection reason (optional):", "Reject Matches")

    Dim sel As Range
    Set sel = Selection

    Dim rejectCount As Long
    rejectCount = 0

    Dim row As Range
    For Each row In sel.Rows
        Dim matchRow As Long
        matchRow = row.row

        If matchRow >= 2 Then
            If ws.Cells(matchRow, 16).Value = "STAGED" Then
                Dim matchID As Long
                matchID = CLng(ws.Cells(matchRow, 1).Value)
                ModStagingManager.RejectMatch matchID, reason
                rejectCount = rejectCount + 1
            End If
        End If
    Next row

    If rejectCount > 0 Then
        UpdateDashboardStats
        MsgBox rejectCount & " matches rejected.", vbInformation, "Reject Complete"
    End If
End Sub

Public Sub AcceptAllHighConfidence()
    ' Accept all high-confidence staged matches.
    Dim stagedCount As Long
    stagedCount = ModStagingManager.GetStagedCount()

    If stagedCount = 0 Then
        MsgBox "No staged matches to accept.", vbInformation, "Accept"
        Exit Sub
    End If

    Dim response As VbMsgBoxResult
    response = MsgBox("Accept all HIGH confidence matches (>=" & _
                      Format(ModConfig.HighConfidenceThreshold(), "0") & "%)?" & vbCrLf & vbCrLf & _
                      "You should review medium and low confidence matches individually.", _
                      vbYesNo + vbQuestion, "Accept High Confidence")
    If response = vbNo Then Exit Sub

    ModStagingManager.AcceptAllHighConfidence
    UpdateDashboardStats
End Sub

Public Sub CreateManualMatchUI()
    ' Prompt for bank and DMS transaction IDs and create a manual match.
    Dim bankID As String
    bankID = InputBox("Enter Bank Transaction Row ID:", "Manual Match")
    If bankID = "" Then Exit Sub

    Dim dmsID As String
    dmsID = InputBox("Enter DMS Transaction Row ID:", "Manual Match")
    If dmsID = "" Then Exit Sub

    On Error GoTo HandleError
    Dim matchID As Long
    matchID = ModStagingManager.CreateManualMatch(CLng(bankID), CLng(dmsID))

    If matchID > 0 Then
        UpdateDashboardStats
        MsgBox "Manual match created (ID: " & matchID & ").", _
               vbInformation, "Manual Match"
    Else
        MsgBox "Could not create match. Verify the transaction IDs.", _
               vbExclamation, "Manual Match"
    End If
    Exit Sub

HandleError:
    MsgBox "Error creating manual match:" & vbCrLf & Err.Description, _
           vbCritical, "Error"
End Sub

' ---------------------------------------------------------------------------
' Step 5: Finalize & Export
' ---------------------------------------------------------------------------

Public Sub Step5_FinalizeAndExport()
    UpdateDashboardStatus 5, "IN PROGRESS", "Finalizing..."
    ModExport.FinalizeMonth
    UpdateDashboardStatus 5, "COMPLETE", "Finalized"
    UpdateDashboardStats

    ModAuditTrail.EndSession "Reconciliation session completed"
End Sub

' ---------------------------------------------------------------------------
' Dashboard Updates
' ---------------------------------------------------------------------------

Private Sub UpdateDashboardStatus(ByVal stepNum As Long, ByVal status As String, _
                                   ByVal detail As String)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Dashboard")

    Dim statusRow As Long
    statusRow = 8 + (stepNum - 1) * 2  ' Rows 8, 10, 12, 14, 16

    ws.Cells(statusRow, 4).Value = "[ " & status & " ]"

    Select Case status
        Case "COMPLETE"
            ws.Cells(statusRow, 4).Font.Color = RGB(39, 118, 39)  ' Green
        Case "IN PROGRESS"
            ws.Cells(statusRow, 4).Font.Color = RGB(196, 128, 0)  ' Orange
        Case "NOT STARTED"
            ws.Cells(statusRow, 4).Font.Color = RGB(128, 128, 128)  ' Gray
        Case Else
            ws.Cells(statusRow, 4).Font.Color = RGB(128, 128, 128)
    End Select
End Sub

Public Sub UpdateDashboardStats()
    ' Refresh all dashboard statistics.
    On Error Resume Next

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Dashboard")

    Dim wsBank As Worksheet, wsDMS As Worksheet
    Set wsBank = ThisWorkbook.Sheets("BankData")
    Set wsDMS = ThisWorkbook.Sheets("DMSData")

    Dim totalBank As Long, totalDMS As Long
    totalBank = WorksheetFunction.Max(0, ModHelpers.GetLastRow(wsBank, 1) - 1)
    totalDMS = WorksheetFunction.Max(0, ModHelpers.GetLastRow(wsDMS, 1) - 1)

    ' Count matched
    Dim matchedBank As Long, matchedDMS As Long
    matchedBank = 0: matchedDMS = 0
    Dim lastRow As Long, i As Long

    lastRow = ModHelpers.GetLastRow(wsBank, 1)
    For i = 2 To lastRow
        If wsBank.Cells(i, 10).Value = True Then matchedBank = matchedBank + 1
    Next i

    lastRow = ModHelpers.GetLastRow(wsDMS, 1)
    For i = 2 To lastRow
        If wsDMS.Cells(i, 9).Value = True Then matchedDMS = matchedDMS + 1
    Next i

    Dim stagedCount As Long
    stagedCount = ModStagingManager.GetStagedCount()

    Dim acceptedCount As Long
    acceptedCount = ModStagingManager.GetAcceptedCount()

    ' Write to dashboard (stats start at row 22)
    Dim statRow As Long
    statRow = 22

    ws.Cells(statRow, 3).Value = totalBank
    ws.Cells(statRow + 1, 3).Value = totalDMS
    ws.Cells(statRow + 2, 3).Value = acceptedCount  ' 1:1 reconciled
    ws.Cells(statRow + 3, 3).Value = 0  ' CVR/Split (would need to count separately)
    ws.Cells(statRow + 4, 3).Value = stagedCount
    ws.Cells(statRow + 5, 3).Value = totalBank - matchedBank
    ws.Cells(statRow + 6, 3).Value = totalDMS - matchedDMS

    If (totalBank + totalDMS) > 0 Then
        ws.Cells(statRow + 7, 3).Value = _
            Format((matchedBank + matchedDMS) / (totalBank + totalDMS) * 100, "0.0") & "%"
    Else
        ws.Cells(statRow + 7, 3).Value = "0.0%"
    End If

    ' Last session info
    ws.Cells(statRow + 10, 3).Value = Format(Now, "MM/DD/YYYY HH:MM:SS")
    ws.Cells(statRow + 11, 3).Value = ModHelpers.GetCurrentUserName()
    ws.Cells(statRow + 12, 3).Value = ModConfig.GetConfigValue("CurrentMonth")

    On Error GoTo 0
End Sub

' ---------------------------------------------------------------------------
' Quick Navigation
' ---------------------------------------------------------------------------

Public Sub GoToDashboard()
    ThisWorkbook.Sheets("Dashboard").Activate
End Sub

Public Sub GoToBankData()
    ThisWorkbook.Sheets("BankData").Activate
End Sub

Public Sub GoToDMSData()
    ThisWorkbook.Sheets("DMSData").Activate
End Sub

Public Sub GoToStagedMatches()
    ThisWorkbook.Sheets("StagedMatches").Activate
End Sub

Public Sub GoToReconciled()
    ThisWorkbook.Sheets("Reconciled").Activate
End Sub

Public Sub GoToAuditLog()
    ThisWorkbook.Sheets("AuditLog").Activate
End Sub
