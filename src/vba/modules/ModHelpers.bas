Attribute VB_Name = "ModHelpers"
'===============================================================================
' ModHelpers — Utility Functions
'
' Date parsing, string utilities, currency normalization, check number
' extraction. Zero dependencies — everything else uses this module.
'===============================================================================

Option Explicit

' ---------------------------------------------------------------------------
' Check Number Extraction
' ---------------------------------------------------------------------------

Public Function ExtractCheckNumber(ByVal desc As String) As String
    ' Extract check number from bank description field.
    ' Patterns: CHECK #NNNN, CHK #NNNN, CHECK NNNN, CK #NNNN

    Dim upperDesc As String
    upperDesc = UCase(Trim(desc))

    Dim patterns As Variant
    patterns = Array("CHECK\s*#?\s*(\d{3,8})", _
                     "CHK\s*#?\s*(\d{3,8})", _
                     "CK\s*#?\s*(\d{3,8})")

    Dim regEx As Object
    Set regEx = CreateObject("VBScript.RegExp")
    regEx.IgnoreCase = True
    regEx.Global = False

    Dim i As Long
    For i = LBound(patterns) To UBound(patterns)
        regEx.Pattern = patterns(i)
        If regEx.Test(upperDesc) Then
            Dim matches As Object
            Set matches = regEx.Execute(upperDesc)
            If matches.Count > 0 Then
                ExtractCheckNumber = matches(0).SubMatches(0)
                Exit Function
            End If
        End If
    Next i

    ExtractCheckNumber = ""
End Function

' ---------------------------------------------------------------------------
' String Utilities
' ---------------------------------------------------------------------------

Public Function CleanDescription(ByVal desc As String) As String
    ' Normalize a description for comparison.
    Dim cleaned As String
    cleaned = UCase(Trim(desc))

    ' Collapse multiple spaces
    Dim regEx As Object
    Set regEx = CreateObject("VBScript.RegExp")
    regEx.Global = True

    regEx.Pattern = "\s+"
    cleaned = regEx.Replace(cleaned, " ")

    ' Remove common noise words
    Dim noiseWords As Variant
    noiseWords = Array("THE", "A", "AN", "FOR", "OF", "TO", "IN", "ON", "AT")

    Dim i As Long
    For i = LBound(noiseWords) To UBound(noiseWords)
        regEx.Pattern = "\b" & noiseWords(i) & "\b"
        cleaned = regEx.Replace(cleaned, "")
    Next i

    ' Final cleanup
    regEx.Pattern = "\s+"
    cleaned = Trim(regEx.Replace(cleaned, " "))

    CleanDescription = cleaned
End Function

Public Function LevenshteinDistance(ByVal s1 As String, ByVal s2 As String) As Long
    ' Compute Levenshtein edit distance between two strings.
    Dim len1 As Long, len2 As Long
    len1 = Len(s1)
    len2 = Len(s2)

    If len1 = 0 Then
        LevenshteinDistance = len2
        Exit Function
    End If
    If len2 = 0 Then
        LevenshteinDistance = len1
        Exit Function
    End If

    ' Use two-row approach for memory efficiency
    Dim prevRow() As Long
    Dim currRow() As Long
    ReDim prevRow(0 To len2)
    ReDim currRow(0 To len2)

    Dim i As Long, j As Long
    For j = 0 To len2
        prevRow(j) = j
    Next j

    For i = 1 To len1
        currRow(0) = i
        For j = 1 To len2
            Dim cost As Long
            If Mid(s1, i, 1) = Mid(s2, j, 1) Then
                cost = 0
            Else
                cost = 1
            End If

            Dim ins As Long, del As Long, sub_ As Long
            ins = prevRow(j) + 1
            del = currRow(j - 1) + 1
            sub_ = prevRow(j - 1) + cost

            currRow(j) = WorksheetFunction.Min(ins, del, sub_)
        Next j

        ' Swap rows
        Dim temp() As Long
        ReDim temp(0 To len2)
        For j = 0 To len2
            temp(j) = currRow(j)
        Next j
        For j = 0 To len2
            prevRow(j) = temp(j)
        Next j
    Next i

    LevenshteinDistance = prevRow(len2)
End Function

' ---------------------------------------------------------------------------
' Currency Utilities
' ---------------------------------------------------------------------------

Public Function NormalizeCurrency(ByVal value As Variant) As Currency
    ' Convert any amount representation to Currency type.
    On Error GoTo HandleError

    If IsNull(value) Or IsEmpty(value) Then
        NormalizeCurrency = 0
        Exit Function
    End If

    Dim strVal As String
    strVal = CStr(value)

    ' Remove currency symbols and commas
    strVal = Replace(strVal, "$", "")
    strVal = Replace(strVal, ",", "")
    strVal = Replace(strVal, " ", "")

    ' Handle parentheses for negatives: (1234.56) → -1234.56
    If Left(strVal, 1) = "(" And Right(strVal, 1) = ")" Then
        strVal = "-" & Mid(strVal, 2, Len(strVal) - 2)
    End If

    NormalizeCurrency = CCur(strVal)
    Exit Function

HandleError:
    NormalizeCurrency = 0
End Function

' ---------------------------------------------------------------------------
' Date Utilities
' ---------------------------------------------------------------------------

Public Function ParseDateFlexible(ByVal dateStr As String) As Date
    ' Flexible date parser supporting MM/DD/YYYY and YYYY-MM-DD formats.
    On Error GoTo TryAlternate

    Dim cleaned As String
    cleaned = Trim(dateStr)

    ' Try standard VBA parsing first
    ParseDateFlexible = CDate(cleaned)
    Exit Function

TryAlternate:
    On Error GoTo HandleError

    ' Try YYYY-MM-DD format
    If Len(cleaned) >= 10 And Mid(cleaned, 5, 1) = "-" Then
        Dim yr As Long, mo As Long, dy As Long
        yr = CLng(Left(cleaned, 4))
        mo = CLng(Mid(cleaned, 6, 2))
        dy = CLng(Mid(cleaned, 9, 2))
        ParseDateFlexible = DateSerial(yr, mo, dy)
        Exit Function
    End If

HandleError:
    ParseDateFlexible = CDate("1/1/1900")  ' Sentinel value
End Function

Public Function DateDiffDays(ByVal d1 As Date, ByVal d2 As Date) As Long
    ' Absolute date difference in days.
    DateDiffDays = Abs(DateDiff("d", d1, d2))
End Function

' ---------------------------------------------------------------------------
' System Utilities
' ---------------------------------------------------------------------------

Public Function GetCurrentUserName() As String
    ' Get the current Windows username for audit trail.
    GetCurrentUserName = Environ("USERNAME")
    If GetCurrentUserName = "" Then
        GetCurrentUserName = Environ("USER")  ' macOS fallback
    End If
    If GetCurrentUserName = "" Then
        GetCurrentUserName = "Unknown"
    End If
End Function

Public Function GenerateSessionID() As String
    ' Generate a unique session ID based on timestamp.
    GenerateSessionID = Format(Now, "YYYYMMDD_HHMMSS") & "_" & _
                        Format(Int(Rnd * 10000), "0000")
End Function

Public Function FormatCurrencyDisplay(ByVal amount As Currency) As String
    ' Standard currency display format.
    FormatCurrencyDisplay = Format(amount, "$#,##0.00")
End Function

' ---------------------------------------------------------------------------
' Sheet Utilities
' ---------------------------------------------------------------------------

Public Function GetLastRow(ByVal ws As Worksheet, Optional ByVal col As Long = 1) As Long
    ' Find the last used row in a column.
    If WorksheetFunction.CountA(ws.Columns(col)) = 0 Then
        GetLastRow = 1  ' Header row only
    Else
        GetLastRow = ws.Cells(ws.Rows.Count, col).End(xlUp).Row
    End If
End Function

Public Function GetNextRow(ByVal ws As Worksheet, Optional ByVal col As Long = 1) As Long
    ' Get the next empty row (after last data row).
    GetNextRow = GetLastRow(ws, col) + 1
End Function
