Attribute VB_Name = "ModMatchEngine"
'===============================================================================
' ModMatchEngine — Core Matching Algorithm (v2)
'
' Pass-rule architecture based on real-world bank reconciliation best practices.
' Modeled after BlackLine pass rules, Oracle Cash Management, and
' Treasury Software matching approaches.
'
' Phase -1: Detect DMS self-canceling pairs (voids/reversals) and exclude
' Phase  0: Pass Rules — deterministic, no scoring
'           Rule 0: Check# confirmed + exact amount -> 100%
'           Rule 1: Unique exact amount + date corroboration -> 95%
'           Rule 1b: Unique exact amount, no date corroboration -> 85%
' Phase  1: Scored 1:1 — for duplicate-amount groups only
'           Score on date + description to discriminate candidates
' Phase  4: Near-Amount — $0.01 tolerance, always staged for review
'
' CVR (many-to-one) and Reverse Split are handled by ModMatchCVR.
' Prior-period fallback is orchestrated by ModMain.
'
' Critical rule: NOTHING is auto-committed. Every match is STAGED for review.
'===============================================================================

Option Explicit

' ---------------------------------------------------------------------------
' Main Entry Point
' ---------------------------------------------------------------------------

Public Sub RunMatching(ByVal bankTxns As Collection, ByVal dmsTxns As Collection)
    ' Run the full phased matching pipeline.
    ' Results are staged via ModStagingManager.

    Application.StatusBar = "ABR: Running matching..."
    Application.ScreenUpdating = False

    Dim matchID As Long
    matchID = ModStagingManager.GetNextMatchID()

    ' Track which transactions are assigned in this run
    Dim assignedBankIDs As New Collection
    Dim assignedDMSIDs As New Collection

    ' Phase -1: Identify DMS self-canceling pairs (voids/reversals)
    Application.StatusBar = "ABR: Phase -1 — Detecting reversals..."
    Dim excludedDMSIDs As New Collection
    DetectSelfCancelingPairs dmsTxns, excludedDMSIDs

    ' Phase 0, Rule 0: Check# confirmed + exact amount
    Application.StatusBar = "ABR: Phase 0 — Check# + amount pass rule..."
    Dim rule0Count As Long
    rule0Count = RunPassRuleCheckNumber(bankTxns, dmsTxns, excludedDMSIDs, _
                                         assignedBankIDs, assignedDMSIDs, matchID)

    ' Phase 0, Rule 1: Unique exact amount
    Application.StatusBar = "ABR: Phase 0 — Unique amount pass rule..."
    Dim rule1Count As Long
    rule1Count = RunPassRuleUniqueAmount(bankTxns, dmsTxns, excludedDMSIDs, _
                                          assignedBankIDs, assignedDMSIDs, matchID)

    ' Phase 1: Scored matching for duplicate amounts
    Application.StatusBar = "ABR: Phase 1 — Scoring duplicate amounts..."
    Dim phase1Count As Long
    phase1Count = RunScoredMatching(bankTxns, dmsTxns, excludedDMSIDs, _
                                     assignedBankIDs, assignedDMSIDs, matchID)

    ' Phase 4: Near-amount matching ($0.01 tolerance)
    Application.StatusBar = "ABR: Phase 4 — Near-amount matching..."
    Dim phase4Count As Long
    phase4Count = RunNearAmountMatching(bankTxns, dmsTxns, excludedDMSIDs, _
                                         assignedBankIDs, assignedDMSIDs, matchID)

    Dim totalMatched As Long
    totalMatched = rule0Count + rule1Count + phase1Count + phase4Count

    Application.StatusBar = "ABR: Matching complete. " & totalMatched & _
        " matches staged (Rule0:" & rule0Count & " Rule1:" & rule1Count & _
        " Scored:" & phase1Count & " Near:" & phase4Count & ")"
    Application.ScreenUpdating = True
End Sub

' ---------------------------------------------------------------------------
' Phase -1: Self-Canceling Pair Detection
' ---------------------------------------------------------------------------

Private Sub DetectSelfCancelingPairs(ByVal dmsTxns As Collection, _
                                     ByVal excludedIDs As Collection)
    ' Identify DMS self-canceling pairs: same absolute amount, opposite signs,
    ' same reference/check number, within 30 days. These are voids/reversals
    ' that net to zero and should be excluded from matching.

    Dim i As Long, j As Long
    Dim txnA As clsTransaction, txnB As clsTransaction

    For i = 1 To dmsTxns.Count
        Set txnA = dmsTxns(i)
        If txnA.IsMatched Then GoTo NextPairA
        If IsInCollection(excludedIDs, CStr(txnA.TransactionID)) Then GoTo NextPairA

        For j = i + 1 To dmsTxns.Count
            Set txnB = dmsTxns(j)
            If txnB.IsMatched Then GoTo NextPairB
            If IsInCollection(excludedIDs, CStr(txnB.TransactionID)) Then GoTo NextPairB

            ' Same absolute amount, opposite signs
            If Abs(txnA.Amount) = Abs(txnB.Amount) And txnA.Amount <> 0 Then
                If (txnA.Amount > 0) <> (txnB.Amount > 0) Then
                    ' Check for same reference
                    Dim sameRef As Boolean
                    sameRef = False

                    If txnA.CheckNumber <> "" And txnA.CheckNumber = txnB.CheckNumber Then
                        sameRef = True
                    ElseIf txnA.ReferenceNumber <> "" And _
                           txnA.ReferenceNumber = txnB.ReferenceNumber Then
                        sameRef = True
                    End If

                    ' Within 30 days
                    If sameRef Then
                        Dim daysDiff As Long
                        daysDiff = Abs(DateDiff("d", txnA.TransactionDate, _
                                                txnB.TransactionDate))
                        If daysDiff <= 30 Then
                            ' Exclude both from matching pool
                            excludedIDs.Add True, CStr(txnA.TransactionID)
                            excludedIDs.Add True, CStr(txnB.TransactionID)

                            ModAuditTrail.LogEvent "REVERSAL_PAIR", _
                                "Excluded DMS IDs " & txnA.TransactionID & " & " & _
                                txnB.TransactionID & " (self-canceling: " & _
                                Format(txnA.Amount, "$#,##0.00") & " / " & _
                                Format(txnB.Amount, "$#,##0.00") & ")"
                            GoTo NextPairA
                        End If
                    End If
                End If
            End If
NextPairB:
        Next j
NextPairA:
    Next i
End Sub

' ---------------------------------------------------------------------------
' Phase 0, Rule 0: Check# Confirmed + Exact Amount
' ---------------------------------------------------------------------------

Private Function RunPassRuleCheckNumber(ByVal bankTxns As Collection, _
                                         ByVal dmsTxns As Collection, _
                                         ByVal excludedDMSIDs As Collection, _
                                         ByVal assignedBankIDs As Collection, _
                                         ByVal assignedDMSIDs As Collection, _
                                         ByRef matchID As Long) As Long
    ' Match transactions where check numbers match AND amounts are exact.
    ' This is the most definitive match type — 100% confidence.

    Dim matchCount As Long
    matchCount = 0

    Dim bankTxn As clsTransaction
    Dim dmsTxn As clsTransaction
    Dim b As Long, d As Long

    For b = 1 To bankTxns.Count
        Set bankTxn = bankTxns(b)
        If bankTxn.IsMatched Then GoTo NextBankR0
        If IsInCollection(assignedBankIDs, CStr(bankTxn.TransactionID)) Then GoTo NextBankR0
        If bankTxn.CheckNumber = "" Then GoTo NextBankR0

        For d = 1 To dmsTxns.Count
            Set dmsTxn = dmsTxns(d)
            If dmsTxn.IsMatched Then GoTo NextDMSR0
            If IsInCollection(assignedDMSIDs, CStr(dmsTxn.TransactionID)) Then GoTo NextDMSR0
            If IsInCollection(excludedDMSIDs, CStr(dmsTxn.TransactionID)) Then GoTo NextDMSR0

            ' Get DMS check number
            Dim dmsCheck As String
            dmsCheck = dmsTxn.CheckNumber
            If dmsTxn.TypeCode = "CHK" And dmsCheck = "" Then
                dmsCheck = dmsTxn.ReferenceNumber
            End If
            If dmsCheck = "" Then GoTo NextDMSR0

            ' Check# must match
            If bankTxn.CheckNumber <> dmsCheck Then GoTo NextDMSR0

            ' Amount must be exact
            If bankTxn.Amount <> dmsTxn.Amount Then GoTo NextDMSR0

            ' MATCH — 100% confidence
            Dim daysDiff As Long
            daysDiff = Abs(DateDiff("d", bankTxn.TransactionDate, dmsTxn.TransactionDate))

            Dim result As New clsMatchResult
            result.MatchID = matchID
            matchID = matchID + 1
            result.BankTransactionIDs = CStr(bankTxn.TransactionID)
            result.DMSTransactionIDs = CStr(dmsTxn.TransactionID)
            result.ConfidenceScore = 100
            result.MatchType = "ONE_TO_ONE"
            result.AmountDifference = 0
            result.DateDifference = daysDiff
            result.BankAmount = bankTxn.Amount
            result.DMSAmount = dmsTxn.Amount
            result.BankDescription = bankTxn.Description
            result.DMSDescription = dmsTxn.Description
            result.BankDate = bankTxn.TransactionDate
            result.DMSDate = dmsTxn.TransactionDate
            result.CheckNumberMatch = "YES"
            result.ScoreBreakdown = "PASS RULE: Check# " & bankTxn.CheckNumber & _
                " confirmed + exact amount " & Format(bankTxn.Amount, "$#,##0.00") & _
                " | " & daysDiff & "d gap -> 100%"

            ' Stage and record
            ModStagingManager.StageMatch result
            assignedBankIDs.Add bankTxn.TransactionID, CStr(bankTxn.TransactionID)
            assignedDMSIDs.Add dmsTxn.TransactionID, CStr(dmsTxn.TransactionID)

            ModImportBank.UpdateBankMatchStatus bankTxn.TransactionID, _
                result.MatchID, result.MatchType, result.ConfidenceScore
            ModImportDMS.UpdateDMSMatchStatus dmsTxn.TransactionID, _
                result.MatchID, result.MatchType, result.ConfidenceScore

            ModAuditTrail.LogMatchProposed result.MatchID, result.MatchType, _
                result.ConfidenceScore, bankTxn.Description, dmsTxn.Description

            matchCount = matchCount + 1
            GoTo NextBankR0  ' This bank txn is matched, move on
NextDMSR0:
        Next d
NextBankR0:
    Next b

    RunPassRuleCheckNumber = matchCount
End Function

' ---------------------------------------------------------------------------
' Phase 0, Rule 1: Unique Exact Amount
' ---------------------------------------------------------------------------

Private Function RunPassRuleUniqueAmount(ByVal bankTxns As Collection, _
                                          ByVal dmsTxns As Collection, _
                                          ByVal excludedDMSIDs As Collection, _
                                          ByVal assignedBankIDs As Collection, _
                                          ByVal assignedDMSIDs As Collection, _
                                          ByRef matchID As Long) As Long
    ' For each unmatched bank transaction, count DMS candidates at exact amount.
    ' If exactly 1 candidate exists:
    '   - Within 5 calendar days: 95% confidence
    '   - Beyond 5 days: 85% confidence (staged for review, not auto-matched)

    Dim matchCount As Long
    matchCount = 0

    Dim bankTxn As clsTransaction
    Dim dmsTxn As clsTransaction
    Dim b As Long, d As Long

    For b = 1 To bankTxns.Count
        Set bankTxn = bankTxns(b)
        If bankTxn.IsMatched Then GoTo NextBankR1
        If IsInCollection(assignedBankIDs, CStr(bankTxn.TransactionID)) Then GoTo NextBankR1

        ' Count DMS candidates at this exact amount
        Dim candidateCount As Long
        candidateCount = 0
        Dim lastCandidate As clsTransaction
        Dim lastCandidateIdx As Long

        For d = 1 To dmsTxns.Count
            Set dmsTxn = dmsTxns(d)
            If dmsTxn.IsMatched Then GoTo NextDMSR1Count
            If IsInCollection(assignedDMSIDs, CStr(dmsTxn.TransactionID)) Then GoTo NextDMSR1Count
            If IsInCollection(excludedDMSIDs, CStr(dmsTxn.TransactionID)) Then GoTo NextDMSR1Count

            If dmsTxn.Amount = bankTxn.Amount Then
                candidateCount = candidateCount + 1
                Set lastCandidate = dmsTxn
                lastCandidateIdx = d
                If candidateCount > 1 Then GoTo NextBankR1  ' Multiple candidates = skip
            End If
NextDMSR1Count:
        Next d

        ' Exactly 1 candidate at this exact amount
        If candidateCount = 1 Then
            Dim daysDiff As Long
            daysDiff = Abs(DateDiff("d", bankTxn.TransactionDate, _
                                    lastCandidate.TransactionDate))

            ' Determine confidence based on date corroboration
            Dim confidence As Double
            Dim breakdown As String

            If daysDiff <= 5 Then
                confidence = 95
                breakdown = "PASS RULE: Unique amount " & _
                    Format(bankTxn.Amount, "$#,##0.00") & _
                    " (1 of 1 candidate) + " & daysDiff & "d date match -> 95%"
            Else
                confidence = 85
                breakdown = "PASS RULE: Unique amount " & _
                    Format(bankTxn.Amount, "$#,##0.00") & _
                    " (1 of 1 candidate) + " & daysDiff & "d gap (no date corroboration) -> 85%"
            End If

            Dim result As New clsMatchResult
            result.MatchID = matchID
            matchID = matchID + 1
            result.BankTransactionIDs = CStr(bankTxn.TransactionID)
            result.DMSTransactionIDs = CStr(lastCandidate.TransactionID)
            result.ConfidenceScore = confidence
            result.MatchType = "ONE_TO_ONE"
            result.AmountDifference = 0
            result.DateDifference = daysDiff
            result.BankAmount = bankTxn.Amount
            result.DMSAmount = lastCandidate.Amount
            result.BankDescription = bankTxn.Description
            result.DMSDescription = lastCandidate.Description
            result.BankDate = bankTxn.TransactionDate
            result.DMSDate = lastCandidate.TransactionDate
            result.CheckNumberMatch = "N/A"
            result.ScoreBreakdown = breakdown

            ' Stage and record
            ModStagingManager.StageMatch result
            assignedBankIDs.Add bankTxn.TransactionID, CStr(bankTxn.TransactionID)
            assignedDMSIDs.Add lastCandidate.TransactionID, CStr(lastCandidate.TransactionID)

            ModImportBank.UpdateBankMatchStatus bankTxn.TransactionID, _
                result.MatchID, result.MatchType, result.ConfidenceScore
            ModImportDMS.UpdateDMSMatchStatus lastCandidate.TransactionID, _
                result.MatchID, result.MatchType, result.ConfidenceScore

            ModAuditTrail.LogMatchProposed result.MatchID, result.MatchType, _
                result.ConfidenceScore, bankTxn.Description, lastCandidate.Description

            matchCount = matchCount + 1
        End If
NextBankR1:
    Next b

    RunPassRuleUniqueAmount = matchCount
End Function

' ---------------------------------------------------------------------------
' Phase 1: Scored Matching for Duplicate Amounts
' ---------------------------------------------------------------------------

Private Function RunScoredMatching(ByVal bankTxns As Collection, _
                                    ByVal dmsTxns As Collection, _
                                    ByVal excludedDMSIDs As Collection, _
                                    ByVal assignedBankIDs As Collection, _
                                    ByVal assignedDMSIDs As Collection, _
                                    ByRef matchID As Long) As Long
    ' For remaining unmatched transactions where multiple candidates exist
    ' at the same exact amount, score on date proximity and description
    ' similarity to discriminate candidates.

    Dim matchCount As Long
    matchCount = 0

    ' Generate all exact-amount candidate pairs
    Dim candidates As New Collection
    Dim bankTxn As clsTransaction
    Dim dmsTxn As clsTransaction
    Dim b As Long, d As Long

    For b = 1 To bankTxns.Count
        Set bankTxn = bankTxns(b)
        If bankTxn.IsMatched Then GoTo NextBankP1
        If IsInCollection(assignedBankIDs, CStr(bankTxn.TransactionID)) Then GoTo NextBankP1

        For d = 1 To dmsTxns.Count
            Set dmsTxn = dmsTxns(d)
            If dmsTxn.IsMatched Then GoTo NextDMSP1
            If IsInCollection(assignedDMSIDs, CStr(dmsTxn.TransactionID)) Then GoTo NextDMSP1
            If IsInCollection(excludedDMSIDs, CStr(dmsTxn.TransactionID)) Then GoTo NextDMSP1

            ' Must be exact amount match
            If bankTxn.Amount <> dmsTxn.Amount Then GoTo NextDMSP1

            ' Score this pair for discrimination
            Dim daysDiff As Long
            daysDiff = Abs(DateDiff("d", bankTxn.TransactionDate, dmsTxn.TransactionDate))

            Dim descSim As Double
            descSim = ScoreDescription(bankTxn.Description, dmsTxn.Description)

            ' Discrimination score: date proximity (70%) + description (30%)
            ' Date score: 100 for same day, decaying
            Dim dateScore As Double
            Select Case daysDiff
                Case 0: dateScore = 100
                Case 1: dateScore = 95
                Case 2: dateScore = 88
                Case 3: dateScore = 78
                Case 4: dateScore = 68
                Case 5: dateScore = 58
                Case 6 To 7: dateScore = 45
                Case 8 To 14: dateScore = 30
                Case 15 To 30: dateScore = 15
                Case Else: dateScore = 5
            End Select

            Dim pairScore As Double
            pairScore = dateScore * 0.7 + descSim * 0.3

            Dim result As New clsMatchResult
            result.BankTransactionIDs = CStr(bankTxn.TransactionID)
            result.DMSTransactionIDs = CStr(dmsTxn.TransactionID)
            result.ConfidenceScore = pairScore   ' Raw score for sorting
            result.MatchType = "ONE_TO_ONE"
            result.AmountDifference = 0
            result.DateDifference = daysDiff
            result.BankAmount = bankTxn.Amount
            result.DMSAmount = dmsTxn.Amount
            result.BankDescription = bankTxn.Description
            result.DMSDescription = dmsTxn.Description
            result.BankDate = bankTxn.TransactionDate
            result.DMSDate = dmsTxn.TransactionDate
            result.CheckNumberMatch = "N/A"

            candidates.Add result
NextDMSP1:
        Next d

        If b Mod 25 = 0 Then
            Application.StatusBar = "ABR: Phase 1 — Scoring " & _
                Format(b / bankTxns.Count * 100, "0") & "%..."
        End If
NextBankP1:
    Next b

    ' Sort candidates by raw score descending
    Dim sortedCandidates As Collection
    Set sortedCandidates = SortMatchResults(candidates)

    ' Greedy assignment with margin-based confidence
    ' First pass: build a map of candidate counts per bank ID for margin calc
    Dim candidateCountPerBank As New Collection
    Dim c As Long
    For c = 1 To sortedCandidates.Count
        Dim cand As clsMatchResult
        Set cand = sortedCandidates(c)
        Dim bKey As String
        bKey = cand.BankTransactionIDs

        If Not IsInCollection(candidateCountPerBank, bKey) Then
            candidateCountPerBank.Add 1, bKey
        End If
    Next c

    ' Second pass: assign with confidence based on margin
    For c = 1 To sortedCandidates.Count
        Set cand = sortedCandidates(c)

        Dim bankID As Long, dmsID As Long
        bankID = CLng(cand.BankTransactionIDs)
        dmsID = CLng(cand.DMSTransactionIDs)

        If IsInCollection(assignedBankIDs, CStr(bankID)) Then GoTo NextCandP1
        If IsInCollection(assignedDMSIDs, CStr(dmsID)) Then GoTo NextCandP1

        ' Find runner-up score for this bank transaction
        Dim runnerUpScore As Double
        runnerUpScore = -1
        Dim c2 As Long
        For c2 = c + 1 To sortedCandidates.Count
            Dim cand2 As clsMatchResult
            Set cand2 = sortedCandidates(c2)
            If cand2.BankTransactionIDs = CStr(bankID) Then
                If Not IsInCollection(assignedDMSIDs, cand2.DMSTransactionIDs) Then
                    runnerUpScore = cand2.ConfidenceScore
                    Exit For
                End If
            End If
        Next c2

        ' Calculate confidence from margin
        Dim margin As Double
        Dim finalConfidence As Double

        If runnerUpScore < 0 Then
            ' No runner-up (only 1 remaining candidate at this amount for this bank txn)
            ' This can happen when other candidates were consumed by prior assignments
            finalConfidence = 90
        Else
            margin = cand.ConfidenceScore - runnerUpScore
            If margin >= 20 Then
                finalConfidence = 90    ' Clear winner
            ElseIf margin >= 10 Then
                finalConfidence = 80    ' Good separation
            ElseIf margin >= 5 Then
                finalConfidence = 70    ' Modest separation
            Else
                finalConfidence = 60    ' Too close to call — review required
            End If
        End If

        cand.ConfidenceScore = finalConfidence
        cand.ScoreBreakdown = "SCORED: Amount " & Format(cand.BankAmount, "$#,##0.00") & _
            " | Date:" & cand.DateDifference & "d" & _
            " | Margin:" & Format(margin, "0") & _
            IIf(runnerUpScore < 0, " (no runner-up)", " vs runner-up:" & Format(runnerUpScore, "0")) & _
            " -> " & Format(finalConfidence, "0") & "%"

        cand.MatchID = matchID
        matchID = matchID + 1

        ' Stage and record
        ModStagingManager.StageMatch cand
        assignedBankIDs.Add bankID, CStr(bankID)
        assignedDMSIDs.Add dmsID, CStr(dmsID)

        ModImportBank.UpdateBankMatchStatus bankID, cand.MatchID, _
            cand.MatchType, cand.ConfidenceScore
        ModImportDMS.UpdateDMSMatchStatus dmsID, cand.MatchID, _
            cand.MatchType, cand.ConfidenceScore

        ModAuditTrail.LogMatchProposed cand.MatchID, cand.MatchType, _
            cand.ConfidenceScore, cand.BankDescription, cand.DMSDescription

        matchCount = matchCount + 1
NextCandP1:
    Next c

    RunScoredMatching = matchCount
End Function

' ---------------------------------------------------------------------------
' Phase 4: Near-Amount Matching ($0.01 Tolerance)
' ---------------------------------------------------------------------------

Private Function RunNearAmountMatching(ByVal bankTxns As Collection, _
                                        ByVal dmsTxns As Collection, _
                                        ByVal excludedDMSIDs As Collection, _
                                        ByVal assignedBankIDs As Collection, _
                                        ByVal assignedDMSIDs As Collection, _
                                        ByRef matchID As Long) As Long
    ' Match remaining transactions with $0.01 tolerance.
    ' NEVER auto-matched. Always staged for review at 55% confidence.

    Dim matchCount As Long
    matchCount = 0

    Dim bankTxn As clsTransaction
    Dim dmsTxn As clsTransaction
    Dim b As Long, d As Long

    For b = 1 To bankTxns.Count
        Set bankTxn = bankTxns(b)
        If bankTxn.IsMatched Then GoTo NextBankP4
        If IsInCollection(assignedBankIDs, CStr(bankTxn.TransactionID)) Then GoTo NextBankP4

        ' Find best near-amount candidate
        Dim bestCandidate As clsTransaction
        Set bestCandidate = Nothing
        Dim bestDaysDiff As Long
        bestDaysDiff = 9999

        For d = 1 To dmsTxns.Count
            Set dmsTxn = dmsTxns(d)
            If dmsTxn.IsMatched Then GoTo NextDMSP4
            If IsInCollection(assignedDMSIDs, CStr(dmsTxn.TransactionID)) Then GoTo NextDMSP4
            If IsInCollection(excludedDMSIDs, CStr(dmsTxn.TransactionID)) Then GoTo NextDMSP4

            Dim amtDiff As Currency
            amtDiff = Abs(bankTxn.Amount - dmsTxn.Amount)

            ' Must be $0.01 off (not exact — those were handled already)
            If amtDiff > 0 And amtDiff <= 0.01 Then
                Dim daysDiff As Long
                daysDiff = Abs(DateDiff("d", bankTxn.TransactionDate, dmsTxn.TransactionDate))
                If daysDiff < bestDaysDiff Then
                    bestDaysDiff = daysDiff
                    Set bestCandidate = dmsTxn
                End If
            End If
NextDMSP4:
        Next d

        If Not bestCandidate Is Nothing Then
            Dim result As New clsMatchResult
            result.MatchID = matchID
            matchID = matchID + 1
            result.BankTransactionIDs = CStr(bankTxn.TransactionID)
            result.DMSTransactionIDs = CStr(bestCandidate.TransactionID)
            result.ConfidenceScore = 55
            result.MatchType = "ONE_TO_ONE"
            result.AmountDifference = bankTxn.Amount - bestCandidate.Amount
            result.DateDifference = bestDaysDiff
            result.BankAmount = bankTxn.Amount
            result.DMSAmount = bestCandidate.Amount
            result.BankDescription = bankTxn.Description
            result.DMSDescription = bestCandidate.Description
            result.BankDate = bankTxn.TransactionDate
            result.DMSDate = bestCandidate.TransactionDate
            result.CheckNumberMatch = "N/A"
            result.ScoreBreakdown = "NEAR-AMOUNT: $" & _
                Format(Abs(bankTxn.Amount - bestCandidate.Amount), "0.00") & _
                " off (" & Format(bankTxn.Amount, "$#,##0.00") & " vs " & _
                Format(bestCandidate.Amount, "$#,##0.00") & ") | " & _
                bestDaysDiff & "d gap -> 55% (review required)"

            ModStagingManager.StageMatch result
            assignedBankIDs.Add bankTxn.TransactionID, CStr(bankTxn.TransactionID)
            assignedDMSIDs.Add bestCandidate.TransactionID, CStr(bestCandidate.TransactionID)

            ModImportBank.UpdateBankMatchStatus bankTxn.TransactionID, _
                result.MatchID, result.MatchType, result.ConfidenceScore
            ModImportDMS.UpdateDMSMatchStatus bestCandidate.TransactionID, _
                result.MatchID, result.MatchType, result.ConfidenceScore

            matchCount = matchCount + 1
        End If
NextBankP4:
    Next b

    RunNearAmountMatching = matchCount
End Function

' ---------------------------------------------------------------------------
' Description Similarity Scorer (used by Phase 1 as tiebreaker)
' ---------------------------------------------------------------------------

Public Function ScoreDescription(ByVal bankDesc As String, _
                                  ByVal dmsDesc As String) As Double
    ' Score description similarity (0-100). Used to discriminate between
    ' candidates at the same amount.

    Dim cleanBank As String, cleanDMS As String
    cleanBank = ModHelpers.CleanDescription(bankDesc)
    cleanDMS = ModHelpers.CleanDescription(dmsDesc)

    If cleanBank = "" Or cleanDMS = "" Then
        ScoreDescription = 50#  ' Neutral
        Exit Function
    End If

    Dim maxLen As Long
    maxLen = WorksheetFunction.Max(Len(cleanBank), Len(cleanDMS))
    If maxLen = 0 Then
        ScoreDescription = 50#
        Exit Function
    End If

    Dim dist As Long
    dist = ModHelpers.LevenshteinDistance(cleanBank, cleanDMS)

    Dim similarity As Double
    similarity = (1# - (CDbl(dist) / CDbl(maxLen))) * 100#

    ' Bonus for shared significant words
    Dim bankWords() As String, dmsWords() As String
    bankWords = Split(cleanBank, " ")
    dmsWords = Split(cleanDMS, " ")

    Dim significantShared As Long
    significantShared = 0

    Dim i As Long, j As Long
    For i = LBound(bankWords) To UBound(bankWords)
        If Len(bankWords(i)) >= 4 Then
            For j = LBound(dmsWords) To UBound(dmsWords)
                If bankWords(i) = dmsWords(j) Then
                    significantShared = significantShared + 1
                    Exit For
                End If
            Next j
        End If
    Next i

    If significantShared >= 2 Then
        similarity = similarity + 20
    ElseIf significantShared = 1 Then
        similarity = similarity + 10
    End If

    ' Cap at 100
    If similarity > 100 Then similarity = 100
    If similarity < 0 Then similarity = 0

    ScoreDescription = similarity
End Function

' ---------------------------------------------------------------------------
' Legacy Scoring Functions (kept for diagnostic use from Immediate Window)
' ---------------------------------------------------------------------------

Public Function ScoreMatch(ByVal bankTxn As clsTransaction, _
                            ByVal dmsTxn As clsTransaction) As clsMatchResult
    ' Legacy scorer for diagnostic use. Not used by the main pipeline.
    ' Returns a result with deduction-based confidence.

    Dim amtDiff As Currency
    amtDiff = Abs(bankTxn.Amount - dmsTxn.Amount)
    If amtDiff > 0.05 Then
        Set ScoreMatch = Nothing
        Exit Function
    End If

    Dim bankCheck As String
    bankCheck = bankTxn.CheckNumber
    Dim dmsCheck As String
    If dmsTxn.TypeCode = "CHK" Then
        dmsCheck = dmsTxn.CheckNumber
        If dmsCheck = "" Then dmsCheck = dmsTxn.ReferenceNumber
    Else
        dmsCheck = dmsTxn.CheckNumber
    End If

    Dim isVeto As Boolean
    Dim checkConfirmed As Boolean
    isVeto = False
    checkConfirmed = False
    If bankCheck <> "" And dmsCheck <> "" Then
        If bankCheck = dmsCheck Then
            checkConfirmed = True
        Else
            isVeto = True
        End If
    End If

    Dim daysDiff As Long
    daysDiff = Abs(DateDiff("d", bankTxn.TransactionDate, dmsTxn.TransactionDate))

    Dim total As Double
    total = 100#

    If amtDiff > 0.01 Then
        total = total - 40
    ElseIf amtDiff > 0 Then
        total = total - 30
    End If

    If isVeto Then
        total = 25#
        GoTo BuildLegacy
    End If

    If Not checkConfirmed Then
        Select Case daysDiff
            Case 0:
            Case 1 To 2: total = total - 3
            Case 3 To 5: total = total - 8
            Case 6 To 7: total = total - 15
            Case Else: total = total - 25
        End Select
    End If

    If total < 0 Then total = 0

BuildLegacy:
    total = Round(total, 2)

    Dim result As New clsMatchResult
    result.BankTransactionIDs = CStr(bankTxn.TransactionID)
    result.DMSTransactionIDs = CStr(dmsTxn.TransactionID)
    result.ConfidenceScore = total
    result.MatchType = "ONE_TO_ONE"
    result.AmountDifference = bankTxn.Amount - dmsTxn.Amount
    result.DateDifference = daysDiff
    result.BankAmount = bankTxn.Amount
    result.DMSAmount = dmsTxn.Amount
    result.BankDescription = bankTxn.Description
    result.DMSDescription = dmsTxn.Description
    result.BankDate = bankTxn.TransactionDate
    result.DMSDate = dmsTxn.TransactionDate

    If isVeto Then
        result.CheckNumberMatch = "NO"
    ElseIf checkConfirmed Then
        result.CheckNumberMatch = "YES"
    Else
        result.CheckNumberMatch = "N/A"
    End If

    result.ScoreBreakdown = "Legacy diagnostic scorer"
    Set ScoreMatch = result
End Function

' ---------------------------------------------------------------------------
' Sorting Helper
' ---------------------------------------------------------------------------

Private Function SortMatchResults(ByVal unsorted As Collection) As Collection
    ' Simple insertion sort by confidence score descending.
    Dim sorted As New Collection

    If unsorted.Count = 0 Then
        Set SortMatchResults = sorted
        Exit Function
    End If

    Dim i As Long
    For i = 1 To unsorted.Count
        Dim item As clsMatchResult
        Set item = unsorted(i)

        Dim inserted As Boolean
        inserted = False

        Dim j As Long
        For j = 1 To sorted.Count
            Dim existing As clsMatchResult
            Set existing = sorted(j)

            If item.ConfidenceScore > existing.ConfidenceScore Then
                If j = 1 Then
                    sorted.Add item, Before:=1
                Else
                    sorted.Add item, Before:=j
                End If
                inserted = True
                Exit For
            End If
        Next j

        If Not inserted Then
            sorted.Add item
        End If
    Next i

    Set SortMatchResults = sorted
End Function

' ---------------------------------------------------------------------------
' Collection Helper
' ---------------------------------------------------------------------------

Private Function IsInCollection(ByVal coll As Collection, ByVal key As String) As Boolean
    On Error GoTo NotFound
    Dim dummy As Variant
    dummy = coll(key)
    IsInCollection = True
    Exit Function
NotFound:
    IsInCollection = False
End Function
