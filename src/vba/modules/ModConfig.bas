Attribute VB_Name = "ModConfig"
'===============================================================================
' ModConfig — Configuration Management
'
' Reads and writes values from/to the Config sheet. All configurable
' parameters (thresholds, weights, location info) live on the Config sheet.
'===============================================================================

Option Explicit

Private Const CONFIG_SHEET As String = "Config"
Private Const SETTING_COL As Long = 1   ' Column A
Private Const VALUE_COL As Long = 2     ' Column B

' ---------------------------------------------------------------------------
' Read Configuration
' ---------------------------------------------------------------------------

Public Function GetConfigValue(ByVal settingName As String) As String
    ' Read a config value by setting name.
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(CONFIG_SHEET)

    Dim lastRow As Long
    lastRow = ModHelpers.GetLastRow(ws, SETTING_COL)

    Dim i As Long
    For i = 2 To lastRow
        If UCase(Trim(ws.Cells(i, SETTING_COL).Value)) = UCase(Trim(settingName)) Then
            GetConfigValue = CStr(ws.Cells(i, VALUE_COL).Value)
            Exit Function
        End If
    Next i

    ' Return empty string if not found
    GetConfigValue = ""
End Function

Public Function GetConfigDouble(ByVal settingName As String) As Double
    ' Read a config value as Double.
    Dim val As String
    val = GetConfigValue(settingName)
    If val = "" Then
        GetConfigDouble = 0#
    Else
        GetConfigDouble = CDbl(val)
    End If
End Function

Public Function GetConfigLong(ByVal settingName As String) As Long
    ' Read a config value as Long.
    Dim val As String
    val = GetConfigValue(settingName)
    If val = "" Then
        GetConfigLong = 0
    Else
        GetConfigLong = CLng(val)
    End If
End Function

Public Function GetConfigBool(ByVal settingName As String) As Boolean
    ' Read a config value as Boolean.
    GetConfigBool = (UCase(GetConfigValue(settingName)) = "TRUE")
End Function

' ---------------------------------------------------------------------------
' Write Configuration
' ---------------------------------------------------------------------------

Public Sub SetConfigValue(ByVal settingName As String, ByVal newValue As String)
    ' Write a config value by setting name.
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(CONFIG_SHEET)

    Dim lastRow As Long
    lastRow = ModHelpers.GetLastRow(ws, SETTING_COL)

    Dim i As Long
    For i = 2 To lastRow
        If UCase(Trim(ws.Cells(i, SETTING_COL).Value)) = UCase(Trim(settingName)) Then
            ws.Cells(i, VALUE_COL).Value = newValue

            ' Log the change
            ModAuditTrail.LogEvent "CONFIG_CHANGED", _
                settingName & " changed to: " & newValue
            Exit Sub
        End If
    Next i

    ' Setting not found — add it
    Dim newRow As Long
    newRow = lastRow + 1
    ws.Cells(newRow, SETTING_COL).Value = settingName
    ws.Cells(newRow, VALUE_COL).Value = newValue
End Sub

' ---------------------------------------------------------------------------
' Convenience Accessors (frequently used settings)
' ---------------------------------------------------------------------------

Public Function LocationName() As String
    LocationName = GetConfigValue("LocationName")
End Function

Public Function LocationCode() As String
    LocationCode = GetConfigValue("LocationCode")
End Function

Public Function BankType() As String
    BankType = GetConfigValue("BankType")
End Function

Public Function HighConfidenceThreshold() As Double
    Dim val As Double
    val = GetConfigDouble("HighConfidenceThreshold")
    If val = 0 Then val = 85#  ' Default
    HighConfidenceThreshold = val
End Function

Public Function MediumConfidenceThreshold() As Double
    Dim val As Double
    val = GetConfigDouble("MediumConfidenceThreshold")
    If val = 0 Then val = 60#  ' Default
    MediumConfidenceThreshold = val
End Function

Public Function LowConfidenceThreshold() As Double
    Dim val As Double
    val = GetConfigDouble("LowConfidenceThreshold")
    If val = 0 Then val = 40#  ' Default
    LowConfidenceThreshold = val
End Function

Public Function AmountWeight() As Double
    Dim val As Double
    val = GetConfigDouble("AmountWeight")
    If val = 0 Then val = 0.4  ' Default
    AmountWeight = val
End Function

Public Function CheckNumberWeight() As Double
    Dim val As Double
    val = GetConfigDouble("CheckNumberWeight")
    If val = 0 Then val = 0.25  ' Default
    CheckNumberWeight = val
End Function

Public Function DateProximityWeight() As Double
    Dim val As Double
    val = GetConfigDouble("DateProximityWeight")
    If val = 0 Then val = 0.25  ' Default
    DateProximityWeight = val
End Function

Public Function DescriptionWeight() As Double
    Dim val As Double
    val = GetConfigDouble("DescriptionWeight")
    If val = 0 Then val = 0.1  ' Default
    DescriptionWeight = val
End Function

Public Function DateWindowDays() As Long
    Dim val As Long
    val = GetConfigLong("DateWindowDays")
    If val = 0 Then val = 7  ' Default
    DateWindowDays = val
End Function

Public Function CVRTolerance() As Currency
    Dim val As String
    val = GetConfigValue("CVRTolerance")
    If val = "" Then
        CVRTolerance = 0.01
    Else
        CVRTolerance = CCur(val)
    End If
End Function

Public Function MaxCVRFragments() As Long
    Dim val As Long
    val = GetConfigLong("MaxCVRFragments")
    If val = 0 Then val = 4   ' Cap at 4; beyond this, coincidental sums are likely
    MaxCVRFragments = val
End Function

Public Function MaxCVRCandidates() As Long
    Dim val As Long
    val = GetConfigLong("MaxCVRCandidates")
    If val = 0 Then val = 20
    MaxCVRCandidates = val
End Function

Public Function CVRTimeoutSeconds() As Double
    Dim val As Double
    val = GetConfigDouble("CVRTimeoutSeconds")
    If val = 0 Then val = 2#
    CVRTimeoutSeconds = val
End Function
