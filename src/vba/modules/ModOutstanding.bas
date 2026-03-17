'===============================================================================
' ModOutstanding — Outstanding Items Carry-Forward
'
' Manages unmatched items that carry forward between reconciliation periods.
' Import prior-period outstanding items, export current unmatched for next period.
'===============================================================================

Option Explicit

Private Const OUTSTANDING_SHEET As String = "Outstanding"

' Column positions
Private Const COL_ITEM_ID As Long = 1
Private Const COL_SOURCE As Long = 2
Private Const COL_ORIG_PERIOD As Long = 3
Private Const COL_TXN_DATE As Long = 4
Private Const COL_DESC As Long = 5
Private Const COL_AMOUNT As Long = 6
Private Const COL_CHECK_REF As Long = 7
Private Const COL_TYPE_CODE As Long = 8
Private Const COL_PERIODS_OUT As Long = 9
Private Const COL_NOTES As Long = 10

' ---------------------------------------------------------------------------
' Import Outstanding Items from Prior Period
' ---------------------------------------------------------------------------

Public Function ImportOutstanding(Optional ByVal filePath As String = "") As Long
    ' Import outstanding items from a prior period file.
    ' Returns count of items imported.

    If filePath = "" Then
        filePath = Application.GetOpenFilename( _
            FileFilter:="CSV Files (*.csv),*.csv,All Files (*.*),*.*", _
            Title:="Select Outstanding Items File")
        If filePath = "False" Or filePath = "" Then
            ImportOutstanding = 0
            Exit Function
        End If
    End If

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(OUTSTANDING_SHEET)

    Dim startRow As Long
    startRow = ModHelpers.GetNextRow(ws, COL_ITEM_ID)

    Dim fileNum As Integer
    fileNum = FreeFile
    Open filePath For Input As #fileNum

    ' Skip header
    Dim headerLine As String
    Line Input #fileNum, headerLine

    Dim itemCount As Long
    itemCount = 0

    Dim nextID As Long
    If startRow <= 2 Then
        nextID = 1
    Else
        nextID = CLng(ws.Cells(startRow - 1, COL_ITEM_ID).Value) + 1
    End If

    Do While Not EOF(fileNum)
        Dim dataLine As String
        Line Input #fileNum, dataLine

        If Trim(dataLine) = "" Then GoTo NextLineOS

        Dim fields() As String
        fields = Split(dataLine, ",")

        If UBound(fields) < 5 Then GoTo NextLineOS

        On Error GoTo SkipLineOS

        ws.Cells(startRow, COL_ITEM_ID).Value = nextID
        ws.Cells(startRow, COL_SOURCE).Value = Trim(fields(1))
        ws.Cells(startRow, COL_ORIG_PERIOD).Value = Trim(fields(2))
        ws.Cells(startRow, COL_TXN_DATE).Value = ModHelpers.ParseDateFlexible(Trim(fields(3)))
        ws.Cells(startRow, COL_TXN_DATE).NumberFormat = "MM/DD/YYYY"
        ws.Cells(startRow, COL_DESC).Value = Trim(Replace(fields(4), """", ""))
        ws.Cells(startRow, COL_AMOUNT).Value = ModHelpers.NormalizeCurrency(fields(5))
        ws.Cells(startRow, COL_AMOUNT).NumberFormat = "#,##0.00"

        If UBound(fields) >= 6 Then
            ws.Cells(startRow, COL_CHECK_REF).Value = Trim(fields(6))
        End If
        If UBound(fields) >= 7 Then
            ws.Cells(startRow, COL_TYPE_CODE).Value = Trim(fields(7))
        End If
        If UBound(fields) >= 8 Then
            ' Increment periods outstanding
            Dim priorPeriods As Long
            priorPeriods = 0
            On Error Resume Next
            priorPeriods = CLng(Trim(fields(8)))
            On Error GoTo SkipLineOS
            ws.Cells(startRow, COL_PERIODS_OUT).Value = priorPeriods + 1
        Else
            ws.Cells(startRow, COL_PERIODS_OUT).Value = 1
        End If

        If UBound(fields) >= 9 Then
            ws.Cells(startRow, COL_NOTES).Value = Trim(Replace(fields(9), """", ""))
        End If

        nextID = nextID + 1
        startRow = startRow + 1
        itemCount = itemCount + 1
        GoTo NextLineOS

SkipLineOS:
        On Error GoTo 0

NextLineOS:
    Loop

    Close #fileNum

    ModAuditTrail.LogImport "OUTSTANDING", filePath, itemCount
    ImportOutstanding = itemCount
End Function

' ---------------------------------------------------------------------------
' Export Unmatched Items as Outstanding for Next Period
' ---------------------------------------------------------------------------

Public Sub ExportOutstanding(Optional ByVal outputPath As String = "")
    ' Export current unmatched items for carry-forward to next period.

    If outputPath = "" Then
        outputPath = Application.GetSaveAsFilename( _
            InitialFileName:="Outstanding_" & Format(Date, "YYYY_MM") & ".csv", _
            FileFilter:="CSV Files (*.csv),*.csv", _
            Title:="Save Outstanding Items File")
        If outputPath = "False" Or outputPath = "" Then Exit Sub
    End If

    Dim fileNum As Integer
    fileNum = FreeFile
    Open outputPath For Output As #fileNum

    ' Write header
    Print #fileNum, "Item ID,Source,Original Period,Transaction Date,Description," & _
                    "Amount,Check/Reference,Type Code,Periods Outstanding,Notes"

    Dim currentMonth As String
    currentMonth = ModConfig.GetConfigValue("CurrentMonth")
    If currentMonth = "" Then currentMonth = Format(Date, "YYYY-MM")

    Dim itemCount As Long
    itemCount = 0

    ' Export unmatched bank transactions
    Dim wsBankData As Worksheet
    Set wsBankData = ThisWorkbook.Sheets("BankData")
    Dim lastRow As Long
    lastRow = ModHelpers.GetLastRow(wsBankData, 1)

    Dim i As Long
    For i = 2 To lastRow
        If wsBankData.Cells(i, 10).Value = False Then  ' COL_IS_MATCHED
            itemCount = itemCount + 1
            Print #fileNum, itemCount & "," & _
                "BANK," & _
                currentMonth & "," & _
                Format(wsBankData.Cells(i, 2).Value, "MM/DD/YYYY") & "," & _
                """" & CStr(wsBankData.Cells(i, 4).Value) & """," & _
                Format(wsBankData.Cells(i, 5).Value, "0.00") & "," & _
                CStr(wsBankData.Cells(i, 6).Value) & "," & _
                "," & _
                "1,"
        End If
    Next i

    ' Export unmatched DMS transactions
    Dim wsDMSData As Worksheet
    Set wsDMSData = ThisWorkbook.Sheets("DMSData")
    lastRow = ModHelpers.GetLastRow(wsDMSData, 1)

    For i = 2 To lastRow
        If wsDMSData.Cells(i, 9).Value = False Then  ' COL_IS_MATCHED
            itemCount = itemCount + 1
            Print #fileNum, itemCount & "," & _
                "DMS," & _
                currentMonth & "," & _
                Format(wsDMSData.Cells(i, 2).Value, "MM/DD/YYYY") & "," & _
                """" & CStr(wsDMSData.Cells(i, 3).Value) & """," & _
                Format(wsDMSData.Cells(i, 5).Value, "0.00") & "," & _
                CStr(wsDMSData.Cells(i, 4).Value) & "," & _
                CStr(wsDMSData.Cells(i, 6).Value) & "," & _
                "1,"
        End If
    Next i

    ' Also re-export any prior outstanding items that still aren't matched
    Dim wsOutstanding As Worksheet
    Set wsOutstanding = ThisWorkbook.Sheets(OUTSTANDING_SHEET)
    lastRow = ModHelpers.GetLastRow(wsOutstanding, COL_ITEM_ID)

    For i = 2 To lastRow
        ' Check if this outstanding item got matched
        ' (Outstanding items would have been loaded into bank/DMS data and matched)
        ' For now, re-export all outstanding items with incremented period count
        itemCount = itemCount + 1
        Dim periodsOut As Long
        periodsOut = 1
        If Not IsEmpty(wsOutstanding.Cells(i, COL_PERIODS_OUT).Value) Then
            periodsOut = CLng(wsOutstanding.Cells(i, COL_PERIODS_OUT).Value) + 1
        End If

        Print #fileNum, itemCount & "," & _
            CStr(wsOutstanding.Cells(i, COL_SOURCE).Value) & "," & _
            CStr(wsOutstanding.Cells(i, COL_ORIG_PERIOD).Value) & "," & _
            Format(wsOutstanding.Cells(i, COL_TXN_DATE).Value, "MM/DD/YYYY") & "," & _
            """" & CStr(wsOutstanding.Cells(i, COL_DESC).Value) & """," & _
            Format(wsOutstanding.Cells(i, COL_AMOUNT).Value, "0.00") & "," & _
            CStr(wsOutstanding.Cells(i, COL_CHECK_REF).Value) & "," & _
            CStr(wsOutstanding.Cells(i, COL_TYPE_CODE).Value) & "," & _
            periodsOut & ","
    Next i

    Close #fileNum

    ModAuditTrail.LogExport "OUTSTANDING", outputPath
    MsgBox itemCount & " outstanding items exported to:" & vbCrLf & outputPath, _
           vbInformation, "Export Complete"
End Sub
