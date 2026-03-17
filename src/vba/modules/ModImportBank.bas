Attribute VB_Name = "ModImportBank"
'===============================================================================
' ModImportBank — Bank Statement Import
'
' Parsers for Bank of America and Truist CSV export formats.
' Auto-detects format by reading the header row.
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
    ' Auto-detect bank CSV format by reading the header row.
    Dim fileNum As Integer
    fileNum = FreeFile
    Open filePath For Input As #fileNum

    Dim headerLine As String
    Line Input #fileNum, headerLine
    Close #fileNum

    headerLine = LCase(headerLine)

    If InStr(headerLine, "debit") > 0 And InStr(headerLine, "credit") > 0 Then
        DetectBankFormat = "TRUIST"
    ElseIf InStr(headerLine, "amount") > 0 And InStr(headerLine, "running balance") > 0 Then
        DetectBankFormat = "BOFA"
    ElseIf InStr(headerLine, "amount") > 0 Then
        DetectBankFormat = "BOFA"  ' Default guess
    Else
        DetectBankFormat = "UNKNOWN"
    End If
End Function

' ---------------------------------------------------------------------------
' Bank of America Parser
' ---------------------------------------------------------------------------

Private Function ParseBankOfAmerica(ByVal filePath As String) As Long
    ' Parse BofA CSV: Date, Description, Amount, Running Balance

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

        If Trim(dataLine) = "" Then GoTo NextLine

        ' Parse CSV fields
        Dim fields() As String
        fields = ParseCSVLine(dataLine)

        If UBound(fields) < 3 Then GoTo NextLine

        Dim txnDate As Date
        Dim desc As String
        Dim amount As Currency
        Dim balance As Currency
        Dim checkNum As String

        On Error GoTo SkipLine
        txnDate = ModHelpers.ParseDateFlexible(Trim(fields(0)))
        desc = CleanCSVField(fields(1))
        amount = ModHelpers.NormalizeCurrency(fields(2))
        If UBound(fields) >= 3 Then
            balance = ModHelpers.NormalizeCurrency(fields(3))
        End If
        On Error GoTo 0

        ' Extract check number from description
        checkNum = ModHelpers.ExtractCheckNumber(desc)

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
        ws.Cells(startRow, COL_BALANCE).Value = balance
        ws.Cells(startRow, COL_BALANCE).NumberFormat = "#,##0.00"
        ws.Cells(startRow, COL_BANK_SRC).Value = "BOFA"
        ws.Cells(startRow, COL_IMPORT_TS).Value = importTimestamp
        ws.Cells(startRow, COL_IMPORT_TS).NumberFormat = "MM/DD/YYYY HH:MM:SS"
        ws.Cells(startRow, COL_IS_MATCHED).Value = False

        rowID = rowID + 1
        startRow = startRow + 1
        txnCount = txnCount + 1
        GoTo NextLine

SkipLine:
        On Error GoTo 0

NextLine:
    Loop

    Close #fileNum
    ParseBankOfAmerica = txnCount
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
