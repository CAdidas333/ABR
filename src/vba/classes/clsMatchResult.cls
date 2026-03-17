VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMatchResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===============================================================================
' clsMatchResult — Match Proposal Data Object
'
' Represents a single proposed match between bank and DMS transactions.
' Contains confidence score, match type, and status tracking.
'===============================================================================

Option Explicit

Public MatchID As Long
Public BankTransactionIDs As String     ' Comma-separated for many-to-one
Public DMSTransactionIDs As String      ' Comma-separated for many-to-one
Public ConfidenceScore As Double        ' 0.0 to 100.0
Public MatchType As String             ' ONE_TO_ONE, MANY_TO_ONE_BANK, MANY_TO_ONE_DMS
Public Status As String                ' STAGED, ACCEPTED, REJECTED
Public ScoreBreakdown As String        ' Human-readable explanation
Public AmountDifference As Currency
Public DateDifference As Long          ' Days between transaction dates
Public CheckNumberMatch As String      ' "YES", "NO", "N/A"
Public BankAmount As Currency
Public DMSAmount As Currency
Public BankDescription As String
Public DMSDescription As String
Public BankDate As Date
Public DMSDate As Date
Public ProposedTimestamp As Date
Public ResolvedTimestamp As Date
Public ResolvedBy As String
Public RejectReason As String

Private Sub Class_Initialize()
    MatchID = 0
    Status = "STAGED"
    ConfidenceScore = 0
    DateDifference = 0
    CheckNumberMatch = "N/A"
    ProposedTimestamp = Now
End Sub

Public Function GetBankIDs() As Variant
    ' Returns array of bank transaction IDs
    GetBankIDs = Split(BankTransactionIDs, ",")
End Function

Public Function GetDMSIDs() As Variant
    ' Returns array of DMS transaction IDs
    GetDMSIDs = Split(DMSTransactionIDs, ",")
End Function

Public Function GetConfidenceBand() As String
    Dim highThreshold As Double
    Dim medThreshold As Double

    highThreshold = CDbl(ModConfig.GetConfigValue("HighConfidenceThreshold"))
    medThreshold = CDbl(ModConfig.GetConfigValue("MediumConfidenceThreshold"))

    If ConfidenceScore >= highThreshold Then
        GetConfidenceBand = "HIGH"
    ElseIf ConfidenceScore >= medThreshold Then
        GetConfidenceBand = "MEDIUM"
    Else
        GetConfidenceBand = "LOW"
    End If
End Function
