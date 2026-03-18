Attribute VB_Name = "ModResetAndRerun"
'===============================================================================
' ModResetAndRerun — One-shot reset and re-run for clean matching
'
' Call ResetAndRerun from the Immediate Window to:
'   1. Delete duplicate April DMS rows (keep May + one April = 1784 rows)
'   2. Clear all match flags and staged/reconciled data
'   3. Re-run the full matching pipeline
'   4. Auto-accept high confidence matches
'   5. Run CVR matching on remaining unmatched
'   6. Print summary
'===============================================================================

Option Explicit

Public Sub ResetAndRerun()
    Application.ScreenUpdating = False
    Application.StatusBar = "ABR: Resetting and re-running..."

    ' --- Step 1: Fix duplicate April DMS import ---
    Dim wsDMS As Worksheet
    Set wsDMS = ThisWorkbook.Sheets("DMSData")
    Dim dmsLastRow As Long
    dmsLastRow = wsDMS.Cells(wsDMS.Rows.Count, 1).End(xlUp).Row

    ' Expected: 1784 data rows (827 May + 957 April) + 1 header = 1785
    ' If more than 1785 rows, delete the excess (duplicate April)
    If dmsLastRow > 1785 Then
        wsDMS.Rows("1786:" & dmsLastRow).Delete
        Debug.Print "Deleted duplicate DMS rows 1786-" & dmsLastRow
    End If

    dmsLastRow = wsDMS.Cells(wsDMS.Rows.Count, 1).End(xlUp).Row
    Debug.Print "DMS rows after cleanup: " & dmsLastRow - 1

    ' --- Step 2: Clear ALL match flags ---
    Dim wsBank As Worksheet
    Set wsBank = ThisWorkbook.Sheets("BankData")
    Dim bankLastRow As Long
    bankLastRow = wsBank.Cells(wsBank.Rows.Count, 1).End(xlUp).Row

    ' BankData: columns I-M (9-13) = IsMatched, MatchID, MatchType, Confidence
    If bankLastRow > 1 Then
        wsBank.Range(wsBank.Cells(2, 9), wsBank.Cells(bankLastRow, 13)).ClearContents
    End If
    Debug.Print "Cleared BankData match flags (" & bankLastRow - 1 & " rows)"

    ' DMSData: columns I-L (9-12) = IsMatched, MatchID, MatchType, Confidence
    dmsLastRow = wsDMS.Cells(wsDMS.Rows.Count, 1).End(xlUp).Row
    If dmsLastRow > 1 Then
        wsDMS.Range(wsDMS.Cells(2, 9), wsDMS.Cells(dmsLastRow, 12)).ClearContents
    End If
    Debug.Print "Cleared DMSData match flags (" & dmsLastRow - 1 & " rows)"

    ' StagedMatches: clear all data rows
    Dim wsStaged As Worksheet
    Set wsStaged = ThisWorkbook.Sheets("StagedMatches")
    Dim stagedLastRow As Long
    stagedLastRow = wsStaged.Cells(wsStaged.Rows.Count, 1).End(xlUp).Row
    If stagedLastRow > 1 Then
        wsStaged.Range(wsStaged.Cells(2, 1), wsStaged.Cells(stagedLastRow, 19)).ClearContents
    End If
    Debug.Print "Cleared StagedMatches (" & stagedLastRow - 1 & " rows)"

    ' Reconciled: clear all data rows
    Dim wsRecon As Worksheet
    Set wsRecon = ThisWorkbook.Sheets("Reconciled")
    Dim reconLastRow As Long
    reconLastRow = wsRecon.Cells(wsRecon.Rows.Count, 1).End(xlUp).Row
    If reconLastRow > 1 Then
        wsRecon.Range(wsRecon.Cells(2, 1), wsRecon.Cells(reconLastRow, 16)).ClearContents
    End If
    Debug.Print "Cleared Reconciled (" & reconLastRow - 1 & " rows)"

    ' --- Step 3: Set Config properly ---
    ModConfig.SetConfigValue "HighConfidenceThreshold", "85"
    ModConfig.SetConfigValue "CurrentMonth", "2025-05"
    Debug.Print "Config: HighConfidenceThreshold=85, CurrentMonth=2025-05"

    ' --- Step 4: Run matching ---
    Debug.Print ""
    Debug.Print "=== RUNNING MATCHING PIPELINE ==="
    Dim bankTxns As Collection
    Set bankTxns = ModImportBank.LoadBankTransactions()
    Dim dmsTxns As Collection
    Set dmsTxns = ModImportDMS.LoadDMSTransactions()

    Debug.Print "Loaded: " & bankTxns.Count & " bank, " & dmsTxns.Count & " DMS"

    ModMatchEngine.RunMatching bankTxns, dmsTxns

    ' --- Step 5: Auto-accept high confidence ---
    Debug.Print ""
    Debug.Print "=== AUTO-ACCEPTING HIGH CONFIDENCE ==="
    ModStagingManager.AcceptAllHighConfidence

    ' --- Step 6: Run CVR on remaining unmatched ---
    Debug.Print ""
    Debug.Print "=== RUNNING CVR MATCHING ==="

    ' Reload with updated match status
    Set bankTxns = ModImportBank.LoadBankTransactions()
    Set dmsTxns = ModImportDMS.LoadDMSTransactions()

    ' Build unmatched collections
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

    Debug.Print "Unmatched: " & unmatchedBank.Count & " bank, " & unmatchedDMS.Count & " DMS"

    ModMatchCVR.RunCVRMatching unmatchedBank, unmatchedDMS
    ModMatchCVR.RunReverseSplitMatching unmatchedBank, unmatchedDMS

    ' --- Step 7: Print final summary ---
    Application.ScreenUpdating = True
    Application.StatusBar = False

    Debug.Print ""
    Debug.Print "========================================="
    Debug.Print "  FINAL RESULTS"
    Debug.Print "========================================="

    Dim totalBank As Long
    totalBank = bankLastRow - 1

    Dim matchedBank As Long
    matchedBank = 0
    bankLastRow = wsBank.Cells(wsBank.Rows.Count, 1).End(xlUp).Row
    For i = 2 To bankLastRow
        If wsBank.Cells(i, 10).Value = True Then matchedBank = matchedBank + 1
    Next i

    Dim reconCount As Long
    reconLastRow = wsRecon.Cells(wsRecon.Rows.Count, 1).End(xlUp).Row
    If reconLastRow > 1 Then reconCount = reconLastRow - 1 Else reconCount = 0

    Dim stagedCount As Long
    stagedCount = ModStagingManager.GetStagedCount()

    Debug.Print "  Bank transactions:    " & totalBank
    Debug.Print "  Reconciled (accepted):" & reconCount
    Debug.Print "  Still staged:         " & stagedCount
    Debug.Print "  Unmatched bank:       " & totalBank - matchedBank
    Debug.Print "  Match rate:           " & Format(matchedBank / totalBank * 100, "0.0") & "%"
    Debug.Print "========================================="

    MsgBox "Reset and re-run complete!" & vbCrLf & vbCrLf & _
        "Reconciled: " & reconCount & vbCrLf & _
        "Still staged: " & stagedCount & vbCrLf & _
        "Unmatched: " & totalBank - matchedBank & vbCrLf & _
        "Match rate: " & Format(matchedBank / totalBank * 100, "0.0") & "%", _
        vbInformation, "ABR Reset Complete"
End Sub
