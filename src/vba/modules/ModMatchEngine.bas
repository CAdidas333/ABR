Attribute VB_Name = "ModMatchEngine"
'===============================================================================
' ModMatchEngine — Core Matching Algorithm
'
' Implements the weighted multi-factor confidence scoring engine and the
' greedy 1:1 matching procedure. This is the heart of ABR.
'
' Confidence Score = Amount * W_a + CheckNumber * W_c + Date * W_d + Desc * W_desc
'
' Critical rule: NOTHING is auto-committed. Every match is STAGED for review.
'===============================================================================

Option Explicit

' ---------------------------------------------------------------------------
' Factor Scoring Functions
' ---------------------------------------------------------------------------

Public Function ScoreAmount(ByVal bankAmount As Currency, _
                            ByVal dmsAmount As Currency) As Double
    ' Score amount match (0-100). Acts as a gate: > $0.05 diff = 0.
    Dim diff As Currency
    diff = Abs(bankAmount - dmsAmount)

    If diff = 0 Then
        ScoreAmount = 100#
    ElseIf diff <= 0.01 Then
        ScoreAmount = 98#
    ElseIf diff <= 0.05 Then
        ScoreAmount = 90#
    Else
        ScoreAmount = 0#  ' Gate: not a match candidate
    End If
End Function

Public Function ScoreCheckNumber(ByVal bankCheck As String, _
                                  ByVal dmsCheck As String, _
                                  ByRef isVeto As Boolean) As Double
    ' Score check number match (0-100).
    ' Sets isVeto = True if check numbers mismatch (caps total at 30).
    isVeto = False

    Dim bankClean As String, dmsClean As String
    bankClean = Trim(bankCheck)
    dmsClean = Trim(dmsCheck)

    If bankClean <> "" And dmsClean <> "" Then
        If bankClean = dmsClean Then
            ScoreCheckNumber = 100#
        Else
            ScoreCheckNumber = 0#
            isVeto = True   ' Hard veto: mismatched check numbers
        End If
    Else
        ' One or both missing — inconclusive, neutral score
        ScoreCheckNumber = 50#
    End If
End Function

Public Function ScoreDate(ByVal bankDate As Date, ByVal dmsDate As Date, _
                          Optional ByVal maxWindow As Long = 0) As Double
    ' Score date proximity (0-100). Beyond maxWindow = 0.
    If maxWindow = 0 Then
        maxWindow = ModConfig.DateWindowDays()
    End If

    Dim daysDiff As Long
    daysDiff = ModHelpers.DateDiffDays(bankDate, dmsDate)

    If daysDiff > maxWindow Then
        ScoreDate = 0#
        Exit Function
    End If

    ' Scoring curve
    Select Case daysDiff
        Case 0: ScoreDate = 100#
        Case 1: ScoreDate = 95#
        Case 2: ScoreDate = 85#
        Case 3: ScoreDate = 70#
        Case 4: ScoreDate = 55#
        Case 5: ScoreDate = 40#
        Case 6: ScoreDate = 25#
        Case 7: ScoreDate = 10#
        Case Else: ScoreDate = 0#
    End Select
End Function

Public Function ScoreDescription(ByVal bankDesc As String, _
                                  ByVal dmsDesc As String) As Double
    ' Score description similarity (0-100). Low-weight tiebreaker.
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

    ' CHECK/CHK keyword bonus
    Dim bankHasCheck As Boolean, dmsHasCheck As Boolean
    bankHasCheck = (InStr(cleanBank, "CHECK") > 0 Or InStr(cleanBank, "CHK") > 0)
    dmsHasCheck = (InStr(cleanDMS, "CHECK") > 0 Or InStr(cleanDMS, "CHK") > 0)
    If bankHasCheck And dmsHasCheck Then
        similarity = similarity + 10
    End If

    ' Cap at 100
    If similarity > 100 Then similarity = 100
    If similarity < 0 Then similarity = 0

    ScoreDescription = similarity
End Function

' ---------------------------------------------------------------------------
' Composite Match Scoring
' ---------------------------------------------------------------------------

Public Function ScoreMatch(ByVal bankTxn As clsTransaction, _
                            ByVal dmsTxn As clsTransaction) As clsMatchResult
    ' Compute confidence score for a bank-DMS transaction pair.
    ' Returns Nothing if not a viable match (amount gate fails).

    ' Amount gate
    Dim amountScore As Double
    amountScore = ScoreAmount(bankTxn.Amount, dmsTxn.Amount)
    If amountScore = 0 Then
        Set ScoreMatch = Nothing
        Exit Function
    End If

    ' Check number
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
    Dim checkScore As Double
    checkScore = ScoreCheckNumber(bankCheck, dmsCheck, isVeto)

    ' Date
    Dim dateScore As Double
    dateScore = ScoreDate(bankTxn.TransactionDate, dmsTxn.TransactionDate)

    ' Description
    Dim descScore As Double
    descScore = ScoreDescription(bankTxn.Description, dmsTxn.Description)

    ' Weighted composite
    Dim total As Double
    total = amountScore * ModConfig.AmountWeight() + _
            checkScore * ModConfig.CheckNumberWeight() + _
            dateScore * ModConfig.DateProximityWeight() + _
            descScore * ModConfig.DescriptionWeight()

    ' Check number veto — cap at 30
    If isVeto Then
        If total > 30 Then total = 30
    End If

    ' Round to 2 decimal places
    total = Round(total, 2)

    ' Build result
    Dim result As New clsMatchResult
    result.BankTransactionIDs = CStr(bankTxn.TransactionID)
    result.DMSTransactionIDs = CStr(dmsTxn.TransactionID)
    result.ConfidenceScore = total
    result.MatchType = "ONE_TO_ONE"
    result.AmountDifference = bankTxn.Amount - dmsTxn.Amount
    result.DateDifference = ModHelpers.DateDiffDays(bankTxn.TransactionDate, _
                                                     dmsTxn.TransactionDate)
    result.BankAmount = bankTxn.Amount
    result.DMSAmount = dmsTxn.Amount
    result.BankDescription = bankTxn.Description
    result.DMSDescription = dmsTxn.Description
    result.BankDate = bankTxn.TransactionDate
    result.DMSDate = dmsTxn.TransactionDate

    ' Check number match status
    If bankCheck <> "" And dmsCheck <> "" Then
        If isVeto Then
            result.CheckNumberMatch = "NO"
        Else
            result.CheckNumberMatch = "YES"
        End If
    Else
        result.CheckNumberMatch = "N/A"
    End If

    ' Score breakdown
    result.ScoreBreakdown = "Amount:" & Format(amountScore, "0") & _
        "*" & Format(ModConfig.AmountWeight(), "0.00") & _
        " Check:" & Format(checkScore, "0") & _
        "*" & Format(ModConfig.CheckNumberWeight(), "0.00") & _
        " Date:" & Format(dateScore, "0") & _
        "*" & Format(ModConfig.DateProximityWeight(), "0.00") & _
        " Desc:" & Format(descScore, "0") & _
        "*" & Format(ModConfig.DescriptionWeight(), "0.00")

    If isVeto Then
        result.ScoreBreakdown = result.ScoreBreakdown & " [CHECK# VETO: capped at 30]"
    End If

    Set ScoreMatch = result
End Function

' ---------------------------------------------------------------------------
' 1:1 Matching (Greedy Assignment)
' ---------------------------------------------------------------------------

Public Sub RunMatching(ByVal bankTxns As Collection, ByVal dmsTxns As Collection)
    ' Run the 1:1 matching pass using greedy assignment.
    ' Results are staged via ModStagingManager.

    Application.StatusBar = "ABR: Running auto-matching..."
    Application.ScreenUpdating = False

    ' Step 1: Generate all viable candidate matches
    Dim candidates As New Collection
    Dim bankTxn As clsTransaction
    Dim dmsTxn As clsTransaction

    Dim bankCount As Long, dmsCount As Long
    bankCount = bankTxns.Count
    dmsCount = dmsTxns.Count

    Dim processed As Long
    processed = 0

    Dim b As Long, d As Long
    For b = 1 To bankCount
        Set bankTxn = bankTxns(b)
        If bankTxn.IsMatched Then GoTo NextBankTxn

        For d = 1 To dmsCount
            Set dmsTxn = dmsTxns(d)
            If dmsTxn.IsMatched Then GoTo NextDMSTxn

            Dim result As clsMatchResult
            Set result = ScoreMatch(bankTxn, dmsTxn)

            If Not result Is Nothing Then
                If result.ConfidenceScore >= ModConfig.LowConfidenceThreshold() Then
                    candidates.Add result
                End If
            End If
NextDMSTxn:
        Next d

        processed = processed + 1
        If processed Mod 10 = 0 Then
            Application.StatusBar = "ABR: Scoring matches... " & _
                Format(processed / bankCount * 100, "0") & "%"
        End If
NextBankTxn:
    Next b

    ' Step 2: Sort candidates by confidence descending
    Dim sortedCandidates As Collection
    Set sortedCandidates = SortMatchResults(candidates)

    ' Step 3: Greedy assignment
    Dim matchedBankIDs As New Collection  ' Using as a set
    Dim matchedDMSIDs As New Collection
    Dim matchID As Long
    matchID = ModStagingManager.GetNextMatchID()

    Dim candidate As clsMatchResult
    Dim i As Long
    For i = 1 To sortedCandidates.Count
        Set candidate = sortedCandidates(i)

        Dim bankID As Long, dmsID As Long
        bankID = CLng(candidate.BankTransactionIDs)
        dmsID = CLng(candidate.DMSTransactionIDs)

        ' Check if either transaction already assigned
        If Not IsInCollection(matchedBankIDs, CStr(bankID)) And _
           Not IsInCollection(matchedDMSIDs, CStr(dmsID)) Then

            candidate.MatchID = matchID
            matchID = matchID + 1

            ' Stage the match
            ModStagingManager.StageMatch candidate

            ' Mark as assigned
            matchedBankIDs.Add bankID, CStr(bankID)
            matchedDMSIDs.Add dmsID, CStr(dmsID)

            ' Update source sheets
            ModImportBank.UpdateBankMatchStatus bankID, candidate.MatchID, _
                candidate.MatchType, candidate.ConfidenceScore
            ModImportDMS.UpdateDMSMatchStatus dmsID, candidate.MatchID, _
                candidate.MatchType, candidate.ConfidenceScore

            ' Log the proposal
            ModAuditTrail.LogMatchProposed candidate.MatchID, candidate.MatchType, _
                candidate.ConfidenceScore, candidate.BankDescription, _
                candidate.DMSDescription
        End If
    Next i

    Application.StatusBar = "ABR: Matching complete. " & _
        matchedBankIDs.Count & " matches staged for review."
    Application.ScreenUpdating = True
End Sub

' ---------------------------------------------------------------------------
' Sorting Helper
' ---------------------------------------------------------------------------

Private Function SortMatchResults(ByVal unsorted As Collection) As Collection
    ' Simple insertion sort by confidence score descending.
    ' Adequate for typical reconciliation sizes (< 20,000 candidates).
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
