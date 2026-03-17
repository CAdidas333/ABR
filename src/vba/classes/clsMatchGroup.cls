'===============================================================================
' clsMatchGroup — CVR / Split Transaction Group
'
' Represents a group of transactions on one side that match a single
' transaction on the other side (many-to-one relationship).
'===============================================================================

Option Explicit

Public GroupID As Long
Public GroupType As String              ' "BANK_TO_DMS" or "DMS_TO_BANK"
Public TargetTransactionID As Long      ' The single-side transaction
Public TargetAmount As Currency
Public GroupedTransactionIDs As String  ' Comma-separated IDs of grouped side
Public GroupedCount As Long
Public SumAmount As Currency
Public Variance As Currency
Public ConfidenceScore As Double
Public Status As String                ' STAGED, ACCEPTED, REJECTED

Private Sub Class_Initialize()
    GroupID = 0
    GroupedCount = 0
    SumAmount = 0
    Variance = 0
    ConfidenceScore = 0
    Status = "STAGED"
End Sub

Public Function IsWithinTolerance() As Boolean
    Dim tolerance As Currency
    tolerance = CCur(ModConfig.GetConfigValue("CVRTolerance"))
    IsWithinTolerance = (Abs(Variance) <= tolerance)
End Function

Public Function GetGroupedIDs() As Variant
    GetGroupedIDs = Split(GroupedTransactionIDs, ",")
End Function
