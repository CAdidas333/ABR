Attribute VB_Name = "ModAuditTrail"
'===============================================================================
' ModAuditTrail — Audit Trail Logging
'
' Append-only logging to the AuditLog sheet. Every significant action
' is recorded: session starts, match proposals, acceptances, rejections,
' manual matches, CVR groupings, exports, and config changes.
'===============================================================================

Option Explicit

Private Const AUDIT_SHEET As String = "AuditLog"

' Column positions on AuditLog sheet
Private Const COL_LOG_ID As Long = 1
Private Const COL_TIMESTAMP As Long = 2
Private Const COL_USER As Long = 3
Private Const COL_LOCATION As Long = 4
Private Const COL_EVENT_TYPE As Long = 5
Private Const COL_MATCH_ID As Long = 6
Private Const COL_DETAILS As Long = 7
Private Const COL_SESSION_ID As Long = 8

' Module-level session tracking
Private mSessionID As String
Private mNextLogID As Long

' ---------------------------------------------------------------------------
' Session Management
' ---------------------------------------------------------------------------

Public Sub StartSession()
    ' Initialize a new audit session.
    mSessionID = ModHelpers.GenerateSessionID()

    ' Find next log ID
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(AUDIT_SHEET)
    mNextLogID = ModHelpers.GetLastRow(ws, COL_LOG_ID)
    If mNextLogID <= 1 Then
        mNextLogID = 1
    Else
        mNextLogID = ws.Cells(mNextLogID, COL_LOG_ID).Value + 1
    End If

    LogEvent "SESSION_START", "Reconciliation session started"
End Sub

Public Sub EndSession(Optional ByVal summary As String = "")
    Dim details As String
    details = "Session ended"
    If summary <> "" Then
        details = details & ". " & summary
    End If
    LogEvent "SESSION_END", details
End Sub

Public Function GetSessionID() As String
    If mSessionID = "" Then
        mSessionID = ModHelpers.GenerateSessionID()
    End If
    GetSessionID = mSessionID
End Function

' ---------------------------------------------------------------------------
' Core Logging
' ---------------------------------------------------------------------------

Public Sub LogEvent(ByVal eventType As String, ByVal details As String, _
                    Optional ByVal matchID As Long = 0)
    ' Append a single audit log entry.
    On Error GoTo HandleError

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(AUDIT_SHEET)

    Dim nextRow As Long
    nextRow = ModHelpers.GetNextRow(ws, COL_LOG_ID)

    ws.Cells(nextRow, COL_LOG_ID).Value = mNextLogID
    ws.Cells(nextRow, COL_TIMESTAMP).Value = Now
    ws.Cells(nextRow, COL_TIMESTAMP).NumberFormat = "MM/DD/YYYY HH:MM:SS"
    ws.Cells(nextRow, COL_USER).Value = ModHelpers.GetCurrentUserName()
    ws.Cells(nextRow, COL_LOCATION).Value = ModConfig.LocationName()
    ws.Cells(nextRow, COL_EVENT_TYPE).Value = eventType

    If matchID > 0 Then
        ws.Cells(nextRow, COL_MATCH_ID).Value = matchID
    End If

    ws.Cells(nextRow, COL_DETAILS).Value = Left(details, 255)
    ws.Cells(nextRow, COL_SESSION_ID).Value = mSessionID

    mNextLogID = mNextLogID + 1
    Exit Sub

HandleError:
    ' Audit logging should never crash the application
    Debug.Print "AuditTrail error: " & Err.Description
End Sub

' ---------------------------------------------------------------------------
' Convenience Logging Methods
' ---------------------------------------------------------------------------

Public Sub LogImport(ByVal importType As String, ByVal filePath As String, _
                     ByVal recordCount As Long)
    LogEvent "IMPORT_" & UCase(importType), _
             "Imported " & recordCount & " records from: " & filePath
End Sub

Public Sub LogMatchProposed(ByVal matchID As Long, ByVal matchType As String, _
                            ByVal confidence As Double, ByVal bankDesc As String, _
                            ByVal dmsDesc As String)
    LogEvent "MATCH_PROPOSED", _
             matchType & " match proposed (confidence: " & Format(confidence, "0.0") & _
             "%) Bank: " & Left(bankDesc, 50) & " | DMS: " & Left(dmsDesc, 50), _
             matchID
End Sub

Public Sub LogMatchAccepted(ByVal matchID As Long)
    LogEvent "MATCH_ACCEPTED", "Match accepted by controller", matchID
End Sub

Public Sub LogMatchRejected(ByVal matchID As Long, Optional ByVal reason As String = "")
    Dim details As String
    details = "Match rejected by controller"
    If reason <> "" Then
        details = details & ". Reason: " & reason
    End If
    LogEvent "MATCH_REJECTED", details, matchID
End Sub

Public Sub LogManualMatch(ByVal matchID As Long, ByVal bankDesc As String, _
                          ByVal dmsDesc As String)
    LogEvent "MANUAL_MATCH", _
             "Manual match created. Bank: " & Left(bankDesc, 50) & _
             " | DMS: " & Left(dmsDesc, 50), matchID
End Sub

Public Sub LogCVRGroup(ByVal matchID As Long, ByVal fragmentCount As Long, _
                       ByVal totalAmount As Currency)
    LogEvent "CVR_GROUP_CREATED", _
             "CVR group with " & fragmentCount & " fragments, total: " & _
             Format(totalAmount, "$#,##0.00"), matchID
End Sub

Public Sub LogExport(ByVal exportType As String, ByVal filePath As String)
    LogEvent "EXPORT_COMPLETED", _
             exportType & " exported to: " & filePath
End Sub
