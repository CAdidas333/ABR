Attribute VB_Name = "ModImportDMS"
'===============================================================================
' ModImportDMS — R&R DMS GL Export Import
'
' Parses Reynolds & Reynolds DMS GL export CSV files.
' Writes parsed transactions to the DMSData sheet.
'===============================================================================

Option Explicit

Private Const DMS_SHEET As String = "DMSData"

' DMSData column positions
Private Const COL_ROW_ID As Long = 1
Private Const COL_GL_DATE As Long = 2
Private Const COL_DESC As Long = 3
Private Const COL_REF_NUM As Long = 4
Private Const COL_AMOUNT As Long = 5
Private Const COL_TYPE_CODE As Long = 6
Private Const COL_GL_ACCT As Long = 7
Private Const COL_IMPORT_TS As Long = 8
Private Const COL_IS_MATCHED As Long = 9
Private Const COL_MATCH_ID As Long = 10
Private Const COL_MATCH_TYPE As Long = 11
Private Const COL_CONFIDENCE As Long = 12

' ---------------------------------------------------------------------------
' Public Entry Point
' ---------------------------------------------------------------------------

Public Function ImportDMSExport(Optional ByVal filePath As String = "") As Long
    ' Import an R&R DMS GL export CSV file.
    ' Returns the number of transactions imported.

    If filePath = "" Then
        filePath = Application.GetOpenFilename( _
            FileFilter:="CSV Files (*.csv),*.csv,All Files (*.*),*.*", _
            Title:="Select R&R DMS GL Export File")
        If filePath = "False" Or filePath = "" Then
            ImportDMSExport = 0
            Exit Function
        End If
    End If

    Dim txnCount As Long
    txnCount = ParseDMSFile(filePath)

    ' Log the import
    ModAuditTrail.LogImport "DMS", filePath, txnCount

    ImportDMSExport = txnCount
End Function

' ---------------------------------------------------------------------------
' DMS Parser
' ---------------------------------------------------------------------------

Private Function ParseDMSFile(ByVal filePath As String) As Long
    ' Parse R&R DMS CSV: GL Date, Description, Reference, Amount, Type Code

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(DMS_SHEET)

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

        If Trim(dataLine) = "" Then GoTo NextLineDMS

        Dim fields() As String
        fields = ParseCSVLineDMS(dataLine)

        If UBound(fields) < 4 Then GoTo NextLineDMS

        Dim glDate As Date
        Dim desc As String
        Dim refNum As String
        Dim amount As Currency
        Dim typeCode As String

        On Error GoTo SkipLineDMS
        glDate = ModHelpers.ParseDateFlexible(Trim(fields(0)))
        desc = CleanCSVFieldDMS(fields(1))
        refNum = Trim(CleanCSVFieldDMS(fields(2)))
        amount = ModHelpers.NormalizeCurrency(fields(3))
        typeCode = UCase(Trim(CleanCSVFieldDMS(fields(4))))
        On Error GoTo 0

        ' Write to sheet
        ws.Cells(startRow, COL_ROW_ID).Value = rowID
        ws.Cells(startRow, COL_GL_DATE).Value = glDate
        ws.Cells(startRow, COL_GL_DATE).NumberFormat = "MM/DD/YYYY"
        ws.Cells(startRow, COL_DESC).Value = desc
        ws.Cells(startRow, COL_REF_NUM).Value = refNum
        ws.Cells(startRow, COL_AMOUNT).Value = amount
        ws.Cells(startRow, COL_AMOUNT).NumberFormat = "#,##0.00"
        ws.Cells(startRow, COL_TYPE_CODE).Value = typeCode
        ws.Cells(startRow, COL_IMPORT_TS).Value = importTimestamp
        ws.Cells(startRow, COL_IMPORT_TS).NumberFormat = "MM/DD/YYYY HH:MM:SS"
        ws.Cells(startRow, COL_IS_MATCHED).Value = False

        rowID = rowID + 1
        startRow = startRow + 1
        txnCount = txnCount + 1
        GoTo NextLineDMS

SkipLineDMS:
        On Error GoTo 0

NextLineDMS:
    Loop

    Close #fileNum
    ParseDMSFile = txnCount
End Function

' ---------------------------------------------------------------------------
' CSV Parsing Helpers
' ---------------------------------------------------------------------------

Private Function ParseCSVLineDMS(ByVal csvLine As String) As String()
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

    ReDim Preserve result(0 To fieldCount)
    result(fieldCount) = currentField
    ParseCSVLineDMS = result
End Function

Private Function CleanCSVFieldDMS(ByVal field As String) As String
    Dim cleaned As String
    cleaned = Trim(field)
    If Left(cleaned, 1) = """" And Right(cleaned, 1) = """" Then
        cleaned = Mid(cleaned, 2, Len(cleaned) - 2)
    End If
    CleanCSVFieldDMS = Trim(cleaned)
End Function

' ---------------------------------------------------------------------------
' Load DMS Transactions into Collection
' ---------------------------------------------------------------------------

Public Function LoadDMSTransactions() As Collection
    ' Load all DMS transactions from DMSData sheet into a Collection of clsTransaction.
    Dim txns As New Collection
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(DMS_SHEET)

    Dim lastRow As Long
    lastRow = ModHelpers.GetLastRow(ws, COL_ROW_ID)

    If lastRow <= 1 Then
        Set LoadDMSTransactions = txns
        Exit Function
    End If

    Dim i As Long
    For i = 2 To lastRow
        If ws.Cells(i, COL_ROW_ID).Value = "" Then GoTo NextRowDMS

        Dim txn As New clsTransaction
        txn.TransactionID = CLng(ws.Cells(i, COL_ROW_ID).Value)
        txn.Source = "DMS"
        txn.TransactionDate = CDate(ws.Cells(i, COL_GL_DATE).Value)
        txn.Description = CStr(ws.Cells(i, COL_DESC).Value)
        txn.ReferenceNumber = CStr(ws.Cells(i, COL_REF_NUM).Value)
        txn.Amount = CCur(ws.Cells(i, COL_AMOUNT).Value)
        txn.TypeCode = CStr(ws.Cells(i, COL_TYPE_CODE).Value)
        txn.IsMatched = CBool(ws.Cells(i, COL_IS_MATCHED).Value)
        txn.SheetRow = i

        ' For CHK type, extract check number from reference
        If txn.TypeCode = "CHK" Then
            Dim regEx As Object
            Set regEx = CreateObject("VBScript.RegExp")
            regEx.Pattern = "(\d{3,8})"
            If regEx.Test(txn.ReferenceNumber) Then
                Dim matches As Object
                Set matches = regEx.Execute(txn.ReferenceNumber)
                txn.CheckNumber = matches(0).SubMatches(0)
            End If
        End If

        If Not IsEmpty(ws.Cells(i, COL_MATCH_ID).Value) Then
            If ws.Cells(i, COL_MATCH_ID).Value <> "" Then
                txn.MatchID = CLng(ws.Cells(i, COL_MATCH_ID).Value)
            End If
        End If

        txns.Add txn
NextRowDMS:
    Next i

    Set LoadDMSTransactions = txns
End Function

' ---------------------------------------------------------------------------
' Update Match Status on Sheet
' ---------------------------------------------------------------------------

Public Sub UpdateDMSMatchStatus(ByVal txnID As Long, ByVal matchID As Long, _
                                ByVal matchType As String, ByVal confidence As Double)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(DMS_SHEET)

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

Public Sub ClearDMSMatchStatus(ByVal txnID As Long)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(DMS_SHEET)

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
