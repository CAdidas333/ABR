Attribute VB_Name = "ModImportBank"
'===============================================================================
' ModImportBank — Bank Statement Import
'
' Parsers for Bank of America and Truist CSV export formats.
' Auto-detects format by reading the first row.
'
' Bank of America format is a sectioned CSV (no header row):
'   Row 1: "Statement Information,..." — skip
'   Row 2: "Account Summary,..." — skip
'   Rows 3+: Transaction rows in sections:
'     - "Deposits and other credits"  (positive amounts)
'     - "Withdrawals and other Debits" (negative amounts)
'     - "Checks" (negative amounts, D-Mon date format, check number field)
'     - "Daily Ledger Balances" — skip entirely
'
' Writes parsed transactions to the BankData sheet.
'===============================================================================

Option Explicit

Private Const BANK_SHEET As String = "BankData"

' BankData column positions
Private Const COL_ROW_ID As Long = 1
Private Const COL_TXN_DATE As Long = 2
Private Const COL_POST_DATE As Long = 3
Private Const COL_DESC As Long = 4
Private Const COL_AMOUNT As Long = 5
Private Const COL_CHECK_NUM As Long = 6
Private Const COL_BALANCE As Long = 7
Private Const COL_BANK_SRC As Long = 8
Private Const COL_IMPORT_TS As Long = 9
Private Const COL_IS_MATCHED As Long = 10
Private Const COL_MATCH_ID As Long = 11
Private Const COL_MATCH_TYPE As Long = 12
Private Const COL_CONFIDENCE As Long = 13

' Section type constants for BofA sectioned CSV
Private Const SEC_STMT_INFO As String = "statement information"
Private Const SEC_ACCT_SUMMARY As String = "account summary"
Private Const SEC_DEPOSITS As String = "deposits and other credits"
Private Const SEC_WITHDRAWALS As String = "withdrawals and other debits"
Private Const SEC_CHECKS As String = "checks"
Private Const SEC_DAILY_BALANCES As String = "daily ledger balances"

' ---------------------------------------------------------------------------
' Public Entry Point
' ---------------------------------------------------------------------------

Public Function ImportBankStatement(Optional ByVal filePath As String = "") As Long
    ' Import a bank statement CSV file.
    ' Returns the number of transactions imported.
    ' If filePath is empty, prompts user to select a file.

    If filePath = "" Then
        filePath = Application.GetOpenFilename( _
            FileFilter:="CSV Files (*.csv),*.csv,All Files (*.*),*.*", _
            Title:="Select Bank Statement File")
        If filePath = "False" Or filePath = "" Then
            ImportBankStatement = 0
            Exit Function
        End If
    End If

    ' Detect format
    Dim bankFormat As String
    bankFormat = DetectBankFormat(filePath)

    Dim txnCount As Long
    Select Case bankFormat
        Case "BOFA"
            txnCount = ParseBankOfAmerica(filePath)
        Case "TRUIST"
            txnCount = ParseTruist(filePath)
        Case Else
            MsgBox "Unable to detect bank statement format." & vbCrLf & _
                   "Expected Bank of America or Truist CSV format.", _
                   vbExclamation, "Import Error"
            ImportBankStatement = 0
            Exit Function
    End Select

    ' Log the import
    ModAuditTrail.LogImport "BANK", filePath, txnCount

    ImportBankStatement = txnCount
End Function

' ---------------------------------------------------------------------------
' Format Detection
' ---------------------------------------------------------------------------

Public Function DetectBankFormat(ByVal filePath As String) As String
    ' Auto-detect bank CSV format by reading the first row.
    ' BofA sectioned CSV starts with "Statement Information,..."
    ' Truist has a header row with Debit/Credit columns.
    Dim fileNum As Integer
    fileNum = FreeFile
    Open filePath For Input As #fileNum

    Dim headerLine As String
    Line Input #fileNum, headerLine
    Close #fileNum

    Dim headerLower As String
    headerLower = LCase(Trim(headerLine))

    ' BofA sectioned format: first line starts with "Statement Information"
    If Left(headerLower, Len(SEC_STMT_INFO)) = SEC_STMT_INFO Then
        DetectBankFormat = "BOFA"
    ElseIf InStr(headerLower, "debit") > 0 And InStr(headerLower, "credit") > 0 Then
        DetectBankFormat = "TRUIST"
    Else
        DetectBankFormat = "UNKNOWN"
    End If
End Function

' ---------------------------------------------------------------------------
' Bank of America Parser — Sectioned CSV Format
' ---------------------------------------------------------------------------

Private Function ParseBankOfAmerica(ByVal filePath As String) As Long
    ' Parse BofA sectioned CSV.
    '
    ' The file has NO header row. Each row's first field is the section type.
    ' We extract the statement year from the first row's date range, then
    ' parse transaction rows from three sections:
    '   - Deposits and other credits
    '   - Withdrawals and other Debits
    '   - Checks (uses D-Mon date format)
    ' We skip Statement Information, Account Summary, and Daily Ledger Balances.

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(BANK_SHEET)

    Dim startRow As Long
    startRow = ModHelpers.GetNextRow(ws, COL_ROW_ID)

    Dim fileNum As Integer
    fileNum = FreeFile
    Open filePath For Input As #fileNum

    ' --- Read the first line to extract the statement year ---
    Dim firstLine As String
    Line Input #fileNum, firstLine

    Dim stmtYear As Long
    stmtYear = ExtractStatementYear(firstLine)
    If stmtYear = 0 Then stmtYear = Year(Now)  ' Fallback

    ' Close and reopen to process all lines from the top
    Close #fileNum
    fileNum = FreeFile
    Open filePath For Input As #fileNum

    Dim rowID As Long
    If startRow <= 2 Then
        rowID = 1
    Else
        rowID = ws.Cells(startRow - 1, COL_ROW_ID).Value + 1
    End If

    Dim txnCount As Long
    txnCount = 0

    Dim importTimestamp As Date
    importTimestamp = Now

    Do While Not EOF(fileNum)
        Dim dataLine As String
        Line Input #fileNum, dataLine

        If Trim(dataLine) = "" Then GoTo NextBofALine

        ' Parse the line into fields
        Dim fields() As String
        fields = ParseCSVLine(dataLine)

        If UBound(fields) < 1 Then GoTo NextBofALine

        ' Determine section type from first field
        Dim sectionType As String
        sectionType = LCase(Trim(fields(0)))

        ' Skip non-transaction sections
        If sectionType = SEC_STMT_INFO Then GoTo NextBofALine
        If sectionType = SEC_ACCT_SUMMARY Then GoTo NextBofALine
        If sectionType = SEC_DAILY_BALANCES Then GoTo NextBofALine

        ' Route to appropriate section handler
        Dim txnDate As Date
        Dim desc As String
        Dim amount As Currency
        Dim checkNum As String
        Dim parsed As Boolean
        parsed = False

        If sectionType = SEC_DEPOSITS Then
            ' Deposits and other credits
            ' Fields: Type(0), Date(1), DepositID(2), Amount(3), Description(4), RefNum(5)
            If UBound(fields) < 4 Then GoTo NextBofALine

            On Error GoTo SkipBofALine
            txnDate = ParseBofADate(Trim(fields(1)), stmtYear)
            amount = Abs(ParseBofAAmount(fields(3)))  ' Ensure positive
            desc = CleanCSVField(fields(4))
            checkNum = ""
            On Error GoTo 0
            parsed = True

        ElseIf sectionType = SEC_WITHDRAWALS Then
            ' Withdrawals and other Debits
            ' Fields: Type(0), Date(1), empty(2), Amount(3), Description(4), RefNum(5)
            If UBound(fields) < 4 Then GoTo NextBofALine

            On Error GoTo SkipBofALine
            txnDate = ParseBofADate(Trim(fields(1)), stmtYear)
            amount = -Abs(ParseBofAAmount(fields(3)))  ' Ensure negative
            desc = CleanCSVField(fields(4))
            checkNum = ""
            On Error GoTo 0
            parsed = True

        ElseIf sectionType = SEC_CHECKS Then
            ' Checks
            ' Fields: Type(0), Date(1), CheckNumber(2), Amount(3), empty(4), RefNum(5)
            ' Date is D-Mon format (e.g., "16-May", "2-May")
            ' Check number may have asterisk suffix (e.g., "230280*")
            If UBound(fields) < 3 Then GoTo NextBofALine

            On Error GoTo SkipBofALine
            txnDate = ParseDMonDate(Trim(fields(1)), stmtYear)
            amount = -Abs(ParseBofAAmount(fields(3)))  ' Checks are always negative
            checkNum = StripAsterisk(Trim(fields(2)))
            ' Checks usually don't have a description in field(4); build one
            If UBound(fields) >= 4 And Trim(fields(4)) <> "" Then
                desc = CleanCSVField(fields(4))
            Else
                desc = "Check #" & checkNum
            End If
            On Error GoTo 0
            parsed = True
        End If

        If Not parsed Then GoTo NextBofALine

        ' Write to sheet
        ws.Cells(startRow, COL_ROW_ID).Value = rowID
        ws.Cells(startRow, COL_TXN_DATE).Value = txnDate
        ws.Cells(startRow, COL_TXN_DATE).NumberFormat = "MM/DD/YYYY"
        ws.Cells(startRow, COL_POST_DATE).Value = txnDate  ' BofA doesn't separate post date
        ws.Cells(startRow, COL_POST_DATE).NumberFormat = "MM/DD/YYYY"
        ws.Cells(startRow, COL_DESC).Value = desc
        ws.Cells(startRow, COL_AMOUNT).Value = amount
        ws.Cells(startRow, COL_AMOUNT).NumberFormat = "#,##0.00"
        ws.Cells(startRow, COL_CHECK_NUM).Value = checkNum
        ws.Cells(startRow, COL_BALANCE).Value = ""  ' Sectioned CSV has no per-txn running balance
        ws.Cells(startRow, COL_BANK_SRC).Value = "BOFA"
        ws.Cells(startRow, COL_IMPORT_TS).Value = importTimestamp
        ws.Cells(startRow, COL_IMPORT_TS).NumberFormat = "MM/DD/YYYY HH:MM:SS"
        ws.Cells(startRow, COL_IS_MATCHED).Value = False

        rowID = rowID + 1
        startRow = startRow + 1
        txnCount = txnCount + 1
        GoTo NextBofALine

SkipBofALine:
        On Error GoTo 0

NextBofALine:
    Loop

    Close #fileNum
    ParseBankOfAmerica = txnCount
End Function

' ---------------------------------------------------------------------------
' BofA Helper: Extract Statement Year
' ---------------------------------------------------------------------------

Private Function ExtractStatementYear(ByVal stmtInfoLine As String) As Long
    ' Extract the year from the Statement Information row.
    ' Expected format: "Statement Information,acct#,May 1, 2025 to May 31, 2025,..."
    ' We look for a 4-digit year in the date range portion.

    Dim fields() As String
    fields = ParseCSVLine(stmtInfoLine)

    ' The date range is typically in field index 2
    ' e.g. "May 1, 2025 to May 31, 2025"
    ' But fields may shift due to commas in the date. Scan all fields for a year.
    Dim i As Long
    For i = 0 To UBound(fields)
        Dim fieldVal As String
        fieldVal = Trim(fields(i))
        ' Look for a 4-digit number that looks like a year (2000-2099)
        Dim yr As Long
        yr = ExtractYearFromString(fieldVal)
        If yr >= 2000 And yr <= 2099 Then
            ExtractStatementYear = yr
            Exit Function
        End If
    Next i

    ExtractStatementYear = 0
End Function

Private Function ExtractYearFromString(ByVal s As String) As Long
    ' Find a 4-digit year (2000-2099) in a string.
    Dim i As Long
    For i = 1 To Len(s) - 3
        Dim chunk As String
        chunk = Mid(s, i, 4)
        If IsNumeric(chunk) Then
            Dim val As Long
            val = CLng(chunk)
            If val >= 2000 And val <= 2099 Then
                ' Make sure it's not part of a longer number
                Dim charBefore As String, charAfter As String
                charBefore = ""
                charAfter = ""
                If i > 1 Then charBefore = Mid(s, i - 1, 1)
                If i + 4 <= Len(s) Then charAfter = Mid(s, i + 4, 1)
                If Not IsNumeric(charBefore) And Not IsNumeric(charAfter) Then
                    ExtractYearFromString = val
                    Exit Function
                End If
            End If
        End If
    Next i
    ExtractYearFromString = 0
End Function

' ---------------------------------------------------------------------------
' BofA Helper: Parse M/D/YYYY Date
' ---------------------------------------------------------------------------

Private Function ParseBofADate(ByVal dateStr As String, ByVal stmtYear As Long) As Date
    ' Parse a date in M/D/YYYY format (e.g. "5/1/2025").
    ' Falls back to VBA CDate if direct parsing fails.
    Dim parts() As String
    parts = Split(dateStr, "/")
    If UBound(parts) = 2 Then
        Dim m As Long, d As Long, y As Long
        m = CLng(Trim(parts(0)))
        d = CLng(Trim(parts(1)))
        y = CLng(Trim(parts(2)))
        ParseBofADate = DateSerial(y, m, d)
    Else
        ' Fallback — try VBA's built-in parser
        ParseBofADate = CDate(dateStr)
    End If
End Function

' ---------------------------------------------------------------------------
' BofA Helper: Parse D-Mon Date (for Checks section)
' ---------------------------------------------------------------------------

Private Function ParseDMonDate(ByVal dateStr As String, ByVal stmtYear As Long) As Date
    ' Parse a date in D-Mon format (e.g. "16-May", "2-May").
    ' Uses the statement year since the year is not in this format.
    Dim parts() As String
    parts = Split(dateStr, "-")
    If UBound(parts) = 1 Then
        Dim d As Long
        d = CLng(Trim(parts(0)))
        Dim monStr As String
        monStr = Trim(parts(1))
        Dim m As Long
        m = MonthNameToNumber(monStr)
        If m > 0 Then
            ParseDMonDate = DateSerial(stmtYear, m, d)
            Exit Function
        End If
    End If
    ' Fallback — try VBA's built-in parser, though it may not have a year
    ParseDMonDate = CDate(dateStr & " " & stmtYear)
End Function

Private Function MonthNameToNumber(ByVal monStr As String) As Long
    ' Convert a 3-letter month abbreviation to its number (1-12).
    Select Case LCase(Left(monStr, 3))
        Case "jan": MonthNameToNumber = 1
        Case "feb": MonthNameToNumber = 2
        Case "mar": MonthNameToNumber = 3
        Case "apr": MonthNameToNumber = 4
        Case "may": MonthNameToNumber = 5
        Case "jun": MonthNameToNumber = 6
        Case "jul": MonthNameToNumber = 7
        Case "aug": MonthNameToNumber = 8
        Case "sep": MonthNameToNumber = 9
        Case "oct": MonthNameToNumber = 10
        Case "nov": MonthNameToNumber = 11
        Case "dec": MonthNameToNumber = 12
        Case Else: MonthNameToNumber = 0
    End Select
End Function

' ---------------------------------------------------------------------------
' BofA Helper: Parse Amount with Commas
' ---------------------------------------------------------------------------

Private Function ParseBofAAmount(ByVal amtStr As String) As Currency
    ' Parse an amount that may be quoted and contain commas.
    ' Examples: "216,268.90", "-330,406.45", -41, "-2,684.99"
    Dim cleaned As String
    cleaned = Trim(amtStr)
    ' Remove surrounding quotes
    If Left(cleaned, 1) = """" And Right(cleaned, 1) = """" Then
        cleaned = Mid(cleaned, 2, Len(cleaned) - 2)
    End If
    ' Remove commas
    cleaned = Replace(cleaned, ",", "")
    ' Remove any dollar signs
    cleaned = Replace(cleaned, "$", "")
    cleaned = Trim(cleaned)
    ParseBofAAmount = CCur(cleaned)
End Function

' ---------------------------------------------------------------------------
' BofA Helper: Strip Asterisk from Check Numbers
' ---------------------------------------------------------------------------

Private Function StripAsterisk(ByVal checkNum As String) As String
    ' Remove trailing asterisk from check numbers (e.g. "230280*" -> "230280").
    If Right(checkNum, 1) = "*" Then
        StripAsterisk = Left(checkNum, Len(checkNum) - 1)
    Else
        StripAsterisk = checkNum
    End If
End Function

' ---------------------------------------------------------------------------
' Truist Parser
' ---------------------------------------------------------------------------

Private Function ParseTruist(ByVal filePath As String) As Long
    ' Parse Truist CSV: Date, Description, Debit, Credit, Balance

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(BANK_SHEET)

    Dim startRow As Long
    startRow = ModHelpers.GetNextRow(ws, COL_ROW_ID)

    Dim fileNum As Integer
    fileNum = FreeFile
    Open filePath For Input As #fileNum

    ' Skip header
    Dim headerLine As String
    Line Input #fileNum, headerLine

    Dim rowID As Long
    If startRow <= 2 Then
        rowID = 1
    Else
        rowID = ws.Cells(startRow - 1, COL_ROW_ID).Value + 1
    End If

    Dim txnCount As Long
    txnCount = 0

    Dim importTimestamp As Date
    importTimestamp = Now

    Do While Not EOF(fileNum)
        Dim dataLine As String
        Line Input #fileNum, dataLine

        If Trim(dataLine) = "" Then GoTo NextLineTruist

        Dim fields() As String
        fields = ParseCSVLine(dataLine)

        If UBound(fields) < 4 Then GoTo NextLineTruist

        Dim txnDate As Date
        Dim desc As String
        Dim amount As Currency
        Dim balance As Currency
        Dim checkNum As String

        On Error GoTo SkipLineTruist
        txnDate = ModHelpers.ParseDateFlexible(Trim(fields(0)))
        desc = CleanCSVField(fields(1))

        ' Truist uses separate Debit/Credit columns
        Dim debitStr As String, creditStr As String
        debitStr = Trim(fields(2))
        creditStr = Trim(fields(3))

        If debitStr <> "" Then
            amount = -Abs(ModHelpers.NormalizeCurrency(debitStr))
        ElseIf creditStr <> "" Then
            amount = Abs(ModHelpers.NormalizeCurrency(creditStr))
        Else
            GoTo SkipLineTruist
        End If

        If UBound(fields) >= 4 Then
            balance = ModHelpers.NormalizeCurrency(fields(4))
        End If
        On Error GoTo 0

        checkNum = ModHelpers.ExtractCheckNumber(desc)

        ws.Cells(startRow, COL_ROW_ID).Value = rowID
        ws.Cells(startRow, COL_TXN_DATE).Value = txnDate
        ws.Cells(startRow, COL_TXN_DATE).NumberFormat = "MM/DD/YYYY"
        ws.Cells(startRow, COL_POST_DATE).Value = txnDate
        ws.Cells(startRow, COL_POST_DATE).NumberFormat = "MM/DD/YYYY"
        ws.Cells(startRow, COL_DESC).Value = desc
        ws.Cells(startRow, COL_AMOUNT).Value = amount
        ws.Cells(startRow, COL_AMOUNT).NumberFormat = "#,##0.00"
        ws.Cells(startRow, COL_CHECK_NUM).Value = checkNum
        ws.Cells(startRow, COL_BALANCE).Value = balance
        ws.Cells(startRow, COL_BALANCE).NumberFormat = "#,##0.00"
        ws.Cells(startRow, COL_BANK_SRC).Value = "TRUIST"
        ws.Cells(startRow, COL_IMPORT_TS).Value = importTimestamp
        ws.Cells(startRow, COL_IMPORT_TS).NumberFormat = "MM/DD/YYYY HH:MM:SS"
        ws.Cells(startRow, COL_IS_MATCHED).Value = False

        rowID = rowID + 1
        startRow = startRow + 1
        txnCount = txnCount + 1
        GoTo NextLineTruist

SkipLineTruist:
        On Error GoTo 0

NextLineTruist:
    Loop

    Close #fileNum
    ParseTruist = txnCount
End Function

' ---------------------------------------------------------------------------
' CSV Parsing Helpers
' ---------------------------------------------------------------------------

Private Function ParseCSVLine(ByVal csvLine As String) As String()
    ' Parse a CSV line handling quoted fields with commas.
    Dim result() As String
    Dim fieldCount As Long
    fieldCount = 0

    Dim inQuotes As Boolean
    inQuotes = False

    Dim currentField As String
    currentField = ""

    Dim i As Long
    For i = 1 To Len(csvLine)
        Dim ch As String
        ch = Mid(csvLine, i, 1)

        If ch = """" Then
            inQuotes = Not inQuotes
        ElseIf ch = "," And Not inQuotes Then
            ReDim Preserve result(0 To fieldCount)
            result(fieldCount) = currentField
            fieldCount = fieldCount + 1
            currentField = ""
        Else
            currentField = currentField & ch
        End If
    Next i

    ' Add last field
    ReDim Preserve result(0 To fieldCount)
    result(fieldCount) = currentField

    ParseCSVLine = result
End Function

Private Function CleanCSVField(ByVal field As String) As String
    ' Remove surrounding quotes and trim whitespace.
    Dim cleaned As String
    cleaned = Trim(field)
    If Left(cleaned, 1) = """" And Right(cleaned, 1) = """" Then
        cleaned = Mid(cleaned, 2, Len(cleaned) - 2)
    End If
    CleanCSVField = Trim(cleaned)
End Function

' ---------------------------------------------------------------------------
' Load Bank Transactions into Collection
' ---------------------------------------------------------------------------

Public Function LoadBankTransactions() As Collection
    ' Load all bank transactions from BankData sheet into a Collection of clsTransaction.
    Dim txns As New Collection
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(BANK_SHEET)

    Dim lastRow As Long
    lastRow = ModHelpers.GetLastRow(ws, COL_ROW_ID)

    If lastRow <= 1 Then
        Set LoadBankTransactions = txns
        Exit Function
    End If

    Dim i As Long
    For i = 2 To lastRow
        If ws.Cells(i, COL_ROW_ID).Value = "" Then GoTo NextRow

        Dim txn As New clsTransaction
        txn.TransactionID = CLng(ws.Cells(i, COL_ROW_ID).Value)
        txn.Source = "BANK"
        txn.TransactionDate = CDate(ws.Cells(i, COL_TXN_DATE).Value)
        txn.Description = CStr(ws.Cells(i, COL_DESC).Value)
        txn.Amount = CCur(ws.Cells(i, COL_AMOUNT).Value)
        txn.CheckNumber = CStr(ws.Cells(i, COL_CHECK_NUM).Value)
        txn.BankSource = CStr(ws.Cells(i, COL_BANK_SRC).Value)
        txn.IsMatched = CBool(ws.Cells(i, COL_IS_MATCHED).Value)
        txn.SheetRow = i

        If Not IsEmpty(ws.Cells(i, COL_MATCH_ID).Value) Then
            If ws.Cells(i, COL_MATCH_ID).Value <> "" Then
                txn.MatchID = CLng(ws.Cells(i, COL_MATCH_ID).Value)
            End If
        End If

        txns.Add txn
NextRow:
    Next i

    Set LoadBankTransactions = txns
End Function

' ---------------------------------------------------------------------------
' Update Match Status on Sheet
' ---------------------------------------------------------------------------

Public Sub UpdateBankMatchStatus(ByVal txnID As Long, ByVal matchID As Long, _
                                 ByVal matchType As String, ByVal confidence As Double)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(BANK_SHEET)

    Dim lastRow As Long
    lastRow = ModHelpers.GetLastRow(ws, COL_ROW_ID)

    Dim i As Long
    For i = 2 To lastRow
        If CLng(ws.Cells(i, COL_ROW_ID).Value) = txnID Then
            ws.Cells(i, COL_IS_MATCHED).Value = True
            ws.Cells(i, COL_MATCH_ID).Value = matchID
            ws.Cells(i, COL_MATCH_TYPE).Value = matchType
            ws.Cells(i, COL_CONFIDENCE).Value = confidence / 100#
            ws.Cells(i, COL_CONFIDENCE).NumberFormat = "0.0%"
            Exit Sub
        End If
    Next i
End Sub

Public Sub ClearBankMatchStatus(ByVal txnID As Long)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(BANK_SHEET)

    Dim lastRow As Long
    lastRow = ModHelpers.GetLastRow(ws, COL_ROW_ID)

    Dim i As Long
    For i = 2 To lastRow
        If CLng(ws.Cells(i, COL_ROW_ID).Value) = txnID Then
            ws.Cells(i, COL_IS_MATCHED).Value = False
            ws.Cells(i, COL_MATCH_ID).Value = ""
            ws.Cells(i, COL_MATCH_TYPE).Value = ""
            ws.Cells(i, COL_CONFIDENCE).Value = ""
            Exit Sub
        End If
    Next i
End Sub
