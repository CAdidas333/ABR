Attribute VB_Name = "ModImportDMS"
'===============================================================================
' ModImportDMS — R&R DMS GL Export Import
'
' Parses Reynolds & Reynolds DMS GL export XLSX files.
' The R&R export has 9 columns:
'   1=SRC, 2=Reference#, 3=Date, 4=Port, 5=Control#,
'   6=Debit Amount, 7=Credit Amount, 8=Name, 9=Description
'
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

' R&R source XLSX column positions
Private Const SRC_COL_SRC As Long = 1
Private Const SRC_COL_REF As Long = 2
Private Const SRC_COL_DATE As Long = 3
Private Const SRC_COL_PORT As Long = 4
Private Const SRC_COL_CTRL As Long = 5
Private Const SRC_COL_DEBIT As Long = 6
Private Const SRC_COL_CREDIT As Long = 7
Private Const SRC_COL_NAME As Long = 8
Private Const SRC_COL_DESC As Long = 9

' ---------------------------------------------------------------------------
' Public Entry Point
' ---------------------------------------------------------------------------

Public Function ImportDMSExport(Optional ByVal filePath As String = "") As Long
    ' Import an R&R DMS GL export XLSX file.
    ' Returns the number of transactions imported.

    If filePath = "" Then
        filePath = Application.GetOpenFilename( _
            FileFilter:="Excel Files (*.xlsx),*.xlsx,All Files (*.*),*.*", _
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
' DMS Parser — reads XLSX via Workbooks.Open
' ---------------------------------------------------------------------------

Private Function ParseDMSFile(ByVal filePath As String) As Long
    ' Parse R&R DMS XLSX with 9 columns:
    '   SRC, Reference#, Date, Port, Control#, Debit, Credit, Name, Description
    '
    ' Combines Debit/Credit into signed amount and derives TypeCode from
    ' SRC value and Reference# pattern.

    Dim wsDest As Worksheet
    Set wsDest = ThisWorkbook.Sheets(DMS_SHEET)

    Dim startRow As Long
    startRow = ModHelpers.GetNextRow(wsDest, COL_ROW_ID)

    ' Open the source workbook read-only
    Dim wbSrc As Workbook
    Application.ScreenUpdating = False
    Set wbSrc = Workbooks.Open(Filename:=filePath, ReadOnly:=True, UpdateLinks:=0)

    Dim wsSrc As Worksheet
    Set wsSrc = wbSrc.ActiveSheet

    ' Find last row in source
    Dim lastSrcRow As Long
    lastSrcRow = wsSrc.Cells(wsSrc.Rows.Count, SRC_COL_SRC).End(xlUp).Row

    ' Determine starting row ID
    Dim rowID As Long
    If startRow <= 2 Then
        rowID = 1
    Else
        rowID = wsDest.Cells(startRow - 1, COL_ROW_ID).Value + 1
    End If

    Dim txnCount As Long
    txnCount = 0

    Dim importTimestamp As Date
    importTimestamp = Now

    ' Row 1 is typically a header; start from row 2
    Dim r As Long
    For r = 2 To lastSrcRow
        ' Read source columns
        Dim srcCode As Variant
        srcCode = wsSrc.Cells(r, SRC_COL_SRC).Value

        ' Skip blank rows
        If IsEmpty(srcCode) Or Trim(CStr(srcCode)) = "" Then GoTo NextRowSrc

        Dim refRaw As String
        refRaw = Trim(CStr(wsSrc.Cells(r, SRC_COL_REF).Value))

        Dim glDate As Date
        On Error GoTo SkipRowSrc
        glDate = CDate(wsSrc.Cells(r, SRC_COL_DATE).Value)
        On Error GoTo 0

        Dim portCode As String
        portCode = Trim(CStr(wsSrc.Cells(r, SRC_COL_PORT).Value))

        Dim ctrlNum As String
        ctrlNum = Trim(CStr(wsSrc.Cells(r, SRC_COL_CTRL).Value))

        ' --- Combine Debit/Credit into signed amount ---
        Dim amount As Currency
        Dim debitVal As Variant
        Dim creditVal As Variant
        debitVal = wsSrc.Cells(r, SRC_COL_DEBIT).Value
        creditVal = wsSrc.Cells(r, SRC_COL_CREDIT).Value

        If Not IsEmpty(debitVal) And debitVal <> "" Then
            amount = CCur(debitVal)
        ElseIf Not IsEmpty(creditVal) And creditVal <> "" Then
            amount = CCur(creditVal)   ' Credit values are already negative
        Else
            amount = 0
        End If

        Dim nameField As String
        nameField = Trim(CStr(wsSrc.Cells(r, SRC_COL_NAME).Value))

        Dim descField As String
        descField = Trim(CStr(wsSrc.Cells(r, SRC_COL_DESC).Value))

        ' --- Build combined description ---
        Dim fullDesc As String
        If nameField <> "" And descField <> "" Then
            fullDesc = nameField & " - " & descField
        ElseIf nameField <> "" Then
            fullDesc = nameField
        Else
            fullDesc = descField
        End If

        ' --- Derive TypeCode from SRC + Reference# pattern ---
        Dim typeCode As String
        Dim srcNum As Long
        srcNum = CLng(srcCode)

        typeCode = DeriveTypeCode(srcNum, refRaw)

        ' --- Extract clean reference number ---
        Dim refNum As String
        refNum = CleanReference(srcNum, refRaw)

        ' --- Write to DMSData sheet ---
        wsDest.Cells(startRow, COL_ROW_ID).Value = rowID
        wsDest.Cells(startRow, COL_GL_DATE).Value = glDate
        wsDest.Cells(startRow, COL_GL_DATE).NumberFormat = "MM/DD/YYYY"
        wsDest.Cells(startRow, COL_DESC).Value = fullDesc
        wsDest.Cells(startRow, COL_REF_NUM).Value = refNum
        wsDest.Cells(startRow, COL_AMOUNT).Value = amount
        wsDest.Cells(startRow, COL_AMOUNT).NumberFormat = "#,##0.00"
        wsDest.Cells(startRow, COL_TYPE_CODE).Value = typeCode
        wsDest.Cells(startRow, COL_GL_ACCT).Value = portCode
        wsDest.Cells(startRow, COL_IMPORT_TS).Value = importTimestamp
        wsDest.Cells(startRow, COL_IMPORT_TS).NumberFormat = "MM/DD/YYYY HH:MM:SS"
        wsDest.Cells(startRow, COL_IS_MATCHED).Value = False

        rowID = rowID + 1
        startRow = startRow + 1
        txnCount = txnCount + 1
        GoTo NextRowSrc

SkipRowSrc:
        On Error GoTo 0

NextRowSrc:
    Next r

    ' Close source workbook without saving
    wbSrc.Close SaveChanges:=False
    Application.ScreenUpdating = True

    ParseDMSFile = txnCount
End Function

' ---------------------------------------------------------------------------
' Type Code Derivation
' ---------------------------------------------------------------------------

Private Function DeriveTypeCode(ByVal srcNum As Long, ByVal refRaw As String) As String
    ' Derive a transaction type code from the R&R SRC value and
    ' the Reference# pattern.
    '
    ' SRC=6  with numeric ref   -> "CHK"     (individual check)
    ' SRC=5  with CA prefix     -> "CASH"    (cash batch)
    ' SRC=5  with CK prefix     -> "CKBATCH" (check batch)
    ' SRC=5  with V prefix      -> "VENDOR"  (vendor/wire)
    ' SRC=5  other              -> "BATCH"   (generic batch)
    ' SRC=11                    -> "FINDEP"  (finance deposit)
    ' Otherwise                 -> SRC as string

    Dim refUpper As String
    refUpper = UCase(Trim(refRaw))

    Select Case srcNum
        Case 6
            ' Individual checks — if reference is numeric (possibly with
            ' a trailing letter suffix), classify as CHK
            If IsNumericReference(refRaw) Then
                DeriveTypeCode = "CHK"
            Else
                DeriveTypeCode = "CHK"
            End If

        Case 5
            ' Batch transactions — classify by reference prefix
            If Left(refUpper, 2) = "CA" Then
                DeriveTypeCode = "CASH"
            ElseIf Left(refUpper, 2) = "CK" Then
                DeriveTypeCode = "CKBATCH"
            ElseIf Left(refUpper, 1) = "V" Then
                DeriveTypeCode = "VENDOR"
            Else
                DeriveTypeCode = "BATCH"
            End If

        Case 11
            DeriveTypeCode = "FINDEP"

        Case 30
            DeriveTypeCode = "OTHER"

        Case Else
            DeriveTypeCode = CStr(srcNum)
    End Select
End Function

' ---------------------------------------------------------------------------
' Reference Number Cleaning
' ---------------------------------------------------------------------------

Private Function CleanReference(ByVal srcNum As Long, ByVal refRaw As String) As String
    ' For SRC=6 with a numeric reference (possibly with trailing letter
    ' suffix like "231557A"), strip the trailing letter(s) to get the
    ' pure check number.
    ' For other SRC values, return the raw reference as-is.

    Dim ref As String
    ref = Trim(refRaw)

    If srcNum = 6 Then
        ' Strip trailing letter suffix from check numbers
        ref = StripTrailingLetters(ref)
    End If

    CleanReference = ref
End Function

Private Function StripTrailingLetters(ByVal s As String) As String
    ' Remove trailing alphabetic characters from a string.
    ' E.g., "231557A" -> "231557", "231543" -> "231543"
    ' Only strips if the core part is numeric.

    Dim i As Long
    Dim result As String
    result = s

    If Len(result) = 0 Then
        StripTrailingLetters = result
        Exit Function
    End If

    ' Walk backward, stripping letters
    Do While Len(result) > 0
        Dim lastChar As String
        lastChar = Right(result, 1)
        If lastChar >= "A" And lastChar <= "Z" Then
            result = Left(result, Len(result) - 1)
        ElseIf lastChar >= "a" And lastChar <= "z" Then
            result = Left(result, Len(result) - 1)
        Else
            Exit Do
        End If
    Loop

    ' Only use the stripped version if the remainder is numeric
    If IsNumeric(result) And Len(result) > 0 Then
        StripTrailingLetters = result
    Else
        StripTrailingLetters = s
    End If
End Function

Private Function IsNumericReference(ByVal ref As String) As Boolean
    ' Check if a reference is numeric, possibly with a trailing letter suffix.
    ' "231543" -> True, "231557A" -> True, "CA051425" -> False

    Dim stripped As String
    stripped = StripTrailingLetters(Trim(ref))

    IsNumericReference = IsNumeric(stripped) And Len(stripped) > 0
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
