'===============================================================================
' ModMatchCVR — CVR Many-to-One and Split Transaction Matching
'
' Handles the #1 user pain point: CVR (Customer Vehicle Receivable) transactions
' where a single DMS entry appears as multiple bank deposits.
'
' Also handles reverse splits: multiple DMS entries matching one bank deposit.
'
' Uses subset-sum search with timeout to prevent Excel freezing.
'===============================================================================

Option Explicit

' ---------------------------------------------------------------------------
' CVR Many-to-One Matching (Bank fragments → DMS lump sum)
' ---------------------------------------------------------------------------

Public Sub RunCVRMatching(ByVal unmatchedBank As Collection, _
                          ByVal unmatchedDMS As Collection)
    ' Find groups of bank transactions whose amounts sum to a single DMS entry.
    ' Only considers DMS entries with TypeCode "CVR" or amount > $5000.

    Application.StatusBar = "ABR: Running CVR matching..."

    Dim nextMatchID As Long
    nextMatchID = ModStagingManager.GetNextMatchID()

    Dim tolerance As Currency
    tolerance = ModConfig.CVRTolerance()

    Dim maxFragments As Long
    maxFragments = ModConfig.MaxCVRFragments()

    Dim maxCandidates As Long
    maxCandidates = ModConfig.MaxCVRCandidates()

    Dim timeoutSec As Double
    timeoutSec = ModConfig.CVRTimeoutSeconds()

    Dim dateWindow As Long
    dateWindow = ModConfig.DateWindowDays()

    ' Find CVR candidates on DMS side
    Dim dmsTxn As clsTransaction
    Dim i As Long

    For i = 1 To unmatchedDMS.Count
        Set dmsTxn = unmatchedDMS(i)
        If dmsTxn.IsMatched Then GoTo NextDMSCVR

        ' Only consider CVR type or large amounts
        If dmsTxn.TypeCode <> "CVR" And Abs(dmsTxn.Amount) <= 5000 Then
            GoTo NextDMSCVR
        End If

        ' Find bank candidates: unmatched, same sign, smaller, within date window
        Dim bankCandidates As New Collection
        Dim bankTxn As clsTransaction
        Dim j As Long

        For j = 1 To unmatchedBank.Count
            Set bankTxn = unmatchedBank(j)
            If bankTxn.IsMatched Then GoTo NextBankCandidate

            ' Same sign check
            If (bankTxn.Amount > 0) <> (dmsTxn.Amount > 0) Then GoTo NextBankCandidate

            ' Fragment must be smaller than the whole
            If Abs(bankTxn.Amount) >= Abs(dmsTxn.Amount) Then GoTo NextBankCandidate

            ' Within date window
            If ModHelpers.DateDiffDays(bankTxn.TransactionDate, _
                                       dmsTxn.TransactionDate) > dateWindow Then
                GoTo NextBankCandidate
            End If

            bankCandidates.Add bankTxn

            ' Limit candidates
            If bankCandidates.Count >= maxCandidates Then Exit For
NextBankCandidate:
        Next j

        If bankCandidates.Count < 2 Then GoTo NextDMSCVR

        ' Run subset-sum search
        Dim subsets As Collection
        Set subsets = FindSubsetSum(bankCandidates, dmsTxn.Amount, _
                                    tolerance, maxFragments, timeoutSec)

        ' Process found subsets
        Dim k As Long
        For k = 1 To subsets.Count
            Dim subset As Collection
            Set subset = subsets(k)

            Dim confidence As Double
            confidence = ScoreCVRGroup(subset, dmsTxn, tolerance)

            If confidence < ModConfig.LowConfidenceThreshold() Then GoTo NextSubset

            ' Build match result
            Dim result As New clsMatchResult
            result.MatchID = nextMatchID
            nextMatchID = nextMatchID + 1
            result.MatchType = "MANY_TO_ONE_BANK"
            result.ConfidenceScore = confidence
            result.DMSTransactionIDs = CStr(dmsTxn.TransactionID)
            result.DMSAmount = dmsTxn.Amount
            result.DMSDescription = dmsTxn.Description
            result.DMSDate = dmsTxn.TransactionDate

            ' Build bank IDs and calculate sum
            Dim bankIDs As String
            bankIDs = ""
            Dim groupSum As Currency
            groupSum = 0
            Dim desc As String
            desc = ""
            Dim frag As clsTransaction

            Dim m As Long
            For m = 1 To subset.Count
                Set frag = subset(m)
                If bankIDs <> "" Then bankIDs = bankIDs & ","
                bankIDs = bankIDs & CStr(frag.TransactionID)
                groupSum = groupSum + frag.Amount
                If desc <> "" Then desc = desc & " + "
                desc = desc & Format(frag.Amount, "$#,##0.00")
            Next m

            result.BankTransactionIDs = bankIDs
            result.BankAmount = groupSum
            result.BankDescription = "CVR Group: " & desc
            result.AmountDifference = groupSum - dmsTxn.Amount
            result.ScoreBreakdown = "CVR group: " & subset.Count & " fragments, " & _
                "sum=" & Format(groupSum, "$#,##0.00") & " vs target=" & _
                Format(dmsTxn.Amount, "$#,##0.00")

            ' Stage the match
            ModStagingManager.StageMatch result

            ' Log
            ModAuditTrail.LogCVRGroup result.MatchID, subset.Count, groupSum
            ModAuditTrail.LogMatchProposed result.MatchID, result.MatchType, _
                confidence, result.BankDescription, dmsTxn.Description
NextSubset:
        Next k

        Set bankCandidates = Nothing
NextDMSCVR:
    Next i

    Application.StatusBar = "ABR: CVR matching complete."
End Sub

' ---------------------------------------------------------------------------
' Reverse Split Matching (Multiple DMS → Single Bank)
' ---------------------------------------------------------------------------

Public Sub RunReverseSplitMatching(ByVal unmatchedBank As Collection, _
                                   ByVal unmatchedDMS As Collection)
    ' Find groups of DMS transactions that sum to a single bank transaction.

    Application.StatusBar = "ABR: Running reverse split matching..."

    Dim nextMatchID As Long
    nextMatchID = ModStagingManager.GetNextMatchID()

    Dim tolerance As Currency
    tolerance = ModConfig.CVRTolerance()

    Dim dateWindow As Long
    dateWindow = ModConfig.DateWindowDays()

    Dim bankTxn As clsTransaction
    Dim i As Long

    For i = 1 To unmatchedBank.Count
        Set bankTxn = unmatchedBank(i)
        If bankTxn.IsMatched Then GoTo NextBankSplit

        ' Only consider large amounts
        If Abs(bankTxn.Amount) <= 5000 Then GoTo NextBankSplit

        ' Find DMS candidates
        Dim dmsCandidates As New Collection
        Dim dmsTxn As clsTransaction
        Dim j As Long

        For j = 1 To unmatchedDMS.Count
            Set dmsTxn = unmatchedDMS(j)
            If dmsTxn.IsMatched Then GoTo NextDMSCandidate

            If (dmsTxn.Amount > 0) <> (bankTxn.Amount > 0) Then GoTo NextDMSCandidate
            If Abs(dmsTxn.Amount) >= Abs(bankTxn.Amount) Then GoTo NextDMSCandidate

            If ModHelpers.DateDiffDays(dmsTxn.TransactionDate, _
                                       bankTxn.TransactionDate) > dateWindow Then
                GoTo NextDMSCandidate
            End If

            dmsCandidates.Add dmsTxn
            If dmsCandidates.Count >= ModConfig.MaxCVRCandidates() Then Exit For
NextDMSCandidate:
        Next j

        If dmsCandidates.Count < 2 Then GoTo NextBankSplit

        Dim subsets As Collection
        Set subsets = FindSubsetSum(dmsCandidates, bankTxn.Amount, tolerance, _
                                    ModConfig.MaxCVRFragments(), _
                                    ModConfig.CVRTimeoutSeconds())

        Dim k As Long
        For k = 1 To subsets.Count
            Dim subset As Collection
            Set subset = subsets(k)

            Dim confidence As Double
            confidence = ScoreCVRGroup(subset, bankTxn, tolerance)

            If confidence < ModConfig.LowConfidenceThreshold() Then GoTo NextSubsetSplit

            Dim result As New clsMatchResult
            result.MatchID = nextMatchID
            nextMatchID = nextMatchID + 1
            result.MatchType = "MANY_TO_ONE_DMS"
            result.ConfidenceScore = confidence
            result.BankTransactionIDs = CStr(bankTxn.TransactionID)
            result.BankAmount = bankTxn.Amount
            result.BankDescription = bankTxn.Description
            result.BankDate = bankTxn.TransactionDate

            Dim dmsIDs As String
            dmsIDs = ""
            Dim groupSum As Currency
            groupSum = 0
            Dim desc As String
            desc = ""
            Dim frag As clsTransaction

            Dim m As Long
            For m = 1 To subset.Count
                Set frag = subset(m)
                If dmsIDs <> "" Then dmsIDs = dmsIDs & ","
                dmsIDs = dmsIDs & CStr(frag.TransactionID)
                groupSum = groupSum + frag.Amount
                If desc <> "" Then desc = desc & " + "
                desc = desc & Format(frag.Amount, "$#,##0.00")
            Next m

            result.DMSTransactionIDs = dmsIDs
            result.DMSAmount = groupSum
            result.DMSDescription = "Split: " & desc
            result.AmountDifference = groupSum - bankTxn.Amount
            result.ScoreBreakdown = "Reverse split: " & subset.Count & " DMS entries"

            ModStagingManager.StageMatch result
            ModAuditTrail.LogMatchProposed result.MatchID, result.MatchType, _
                confidence, bankTxn.Description, result.DMSDescription
NextSubsetSplit:
        Next k

        Set dmsCandidates = Nothing
NextBankSplit:
    Next i

    Application.StatusBar = "ABR: Reverse split matching complete."
End Sub

' ---------------------------------------------------------------------------
' Subset Sum Solver
' ---------------------------------------------------------------------------

Public Function FindSubsetSum(ByVal candidates As Collection, _
                               ByVal target As Currency, _
                               ByVal tolerance As Currency, _
                               ByVal maxDepth As Long, _
                               ByVal timeoutSec As Double) As Collection
    ' Find all subsets of candidates whose amounts sum to target within tolerance.
    ' Uses iterative deepening with combinations of size 2..maxDepth.
    ' Respects timeout to prevent Excel freezing.

    Dim results As New Collection
    Dim startTime As Double
    startTime = Timer

    Dim n As Long
    n = candidates.Count

    ' Convert to array for faster access
    Dim amounts() As Currency
    ReDim amounts(1 To n)
    Dim i As Long
    For i = 1 To n
        Dim txn As clsTransaction
        Set txn = candidates(i)
        amounts(i) = txn.Amount
    Next i

    ' Iterative deepening: try combinations of size 2, 3, ..., maxDepth
    Dim depth As Long
    For depth = 2 To WorksheetFunction.Min(maxDepth, n)
        ' Check timeout
        If Timer - startTime > timeoutSec Then Exit For

        ' Generate combinations of 'depth' items from n candidates
        Call FindCombinations(candidates, amounts, n, depth, target, _
                              tolerance, startTime, timeoutSec, results)
    Next depth

    Set FindSubsetSum = results
End Function

Private Sub FindCombinations(ByVal candidates As Collection, _
                              ByRef amounts() As Currency, _
                              ByVal n As Long, _
                              ByVal depth As Long, _
                              ByVal target As Currency, _
                              ByVal tolerance As Currency, _
                              ByVal startTime As Double, _
                              ByVal timeoutSec As Double, _
                              ByVal results As Collection)
    ' Generate all combinations of 'depth' items and check if they sum to target.
    ' Uses iterative approach for depth 2-3, recursive for larger.

    If depth = 2 Then
        Dim i As Long, j As Long
        For i = 1 To n - 1
            If Timer - startTime > timeoutSec Then Exit Sub
            For j = i + 1 To n
                If Abs(amounts(i) + amounts(j) - target) <= tolerance Then
                    Dim combo2 As New Collection
                    combo2.Add candidates(i)
                    combo2.Add candidates(j)
                    results.Add combo2
                End If
            Next j
        Next i

    ElseIf depth = 3 Then
        Dim a As Long, b As Long, c As Long
        For a = 1 To n - 2
            If Timer - startTime > timeoutSec Then Exit Sub
            For b = a + 1 To n - 1
                For c = b + 1 To n
                    If Abs(amounts(a) + amounts(b) + amounts(c) - target) <= tolerance Then
                        Dim combo3 As New Collection
                        combo3.Add candidates(a)
                        combo3.Add candidates(b)
                        combo3.Add candidates(c)
                        results.Add combo3
                    End If
                Next c
            Next b
        Next a

    Else
        ' For depth 4+, use recursive approach
        Dim indices() As Long
        ReDim indices(1 To depth)
        Call RecursiveCombinations(candidates, amounts, n, depth, 1, 1, _
                                   indices, target, tolerance, startTime, _
                                   timeoutSec, results)
    End If
End Sub

Private Sub RecursiveCombinations(ByVal candidates As Collection, _
                                   ByRef amounts() As Currency, _
                                   ByVal n As Long, _
                                   ByVal depth As Long, _
                                   ByVal pos As Long, _
                                   ByVal startIdx As Long, _
                                   ByRef indices() As Long, _
                                   ByVal target As Currency, _
                                   ByVal tolerance As Currency, _
                                   ByVal startTime As Double, _
                                   ByVal timeoutSec As Double, _
                                   ByVal results As Collection)
    ' Recursive combination generator with sum checking.
    If Timer - startTime > timeoutSec Then Exit Sub

    If pos > depth Then
        ' Check sum
        Dim total As Currency
        total = 0
        Dim k As Long
        For k = 1 To depth
            total = total + amounts(indices(k))
        Next k

        If Abs(total - target) <= tolerance Then
            Dim combo As New Collection
            For k = 1 To depth
                combo.Add candidates(indices(k))
            Next k
            results.Add combo
        End If
        Exit Sub
    End If

    Dim i As Long
    For i = startIdx To n - (depth - pos)
        indices(pos) = i
        RecursiveCombinations candidates, amounts, n, depth, pos + 1, i + 1, _
                              indices, target, tolerance, startTime, timeoutSec, results
    Next i
End Sub

' ---------------------------------------------------------------------------
' CVR Group Confidence Scoring
' ---------------------------------------------------------------------------

Public Function ScoreCVRGroup(ByVal group As Collection, _
                               ByVal targetTxn As clsTransaction, _
                               ByVal tolerance As Currency) As Double
    ' Compute confidence for a CVR group match.
    ' Weights: SumAccuracy (50%) + DateClustering (30%) + FragmentCount (20%)

    ' Sum accuracy score
    Dim groupSum As Currency
    groupSum = 0
    Dim minDate As Date, maxDate As Date
    Dim frag As clsTransaction

    Dim i As Long
    For i = 1 To group.Count
        Set frag = group(i)
        groupSum = groupSum + frag.Amount

        If i = 1 Then
            minDate = frag.TransactionDate
            maxDate = frag.TransactionDate
        Else
            If frag.TransactionDate < minDate Then minDate = frag.TransactionDate
            If frag.TransactionDate > maxDate Then maxDate = frag.TransactionDate
        End If
    Next i

    Dim variance As Currency
    variance = Abs(groupSum - targetTxn.Amount)

    Dim sumScore As Double
    If variance = 0 Then
        sumScore = 100#
    ElseIf variance <= 0.01 Then
        sumScore = 95#
    ElseIf variance <= tolerance Then
        sumScore = 80#
    Else
        sumScore = 0#
    End If

    ' Date clustering score
    Dim dateSpread As Long
    dateSpread = DateDiff("d", minDate, maxDate)

    Dim dateScore As Double
    If dateSpread <= 1 Then
        dateScore = 100#
    ElseIf dateSpread <= 3 Then
        dateScore = 80#
    ElseIf dateSpread <= 5 Then
        dateScore = 60#
    Else
        dateScore = 30#
    End If

    ' Fragment count score
    Dim fragScore As Double
    Select Case group.Count
        Case 2: fragScore = 100#
        Case 3: fragScore = 85#
        Case 4: fragScore = 65#
        Case 5: fragScore = 45#
        Case 6: fragScore = 45#
        Case Else: fragScore = 30#
    End Select

    ' Weighted composite
    ScoreCVRGroup = Round(sumScore * 0.5 + dateScore * 0.3 + fragScore * 0.2, 2)
End Function
