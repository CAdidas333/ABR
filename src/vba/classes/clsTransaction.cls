VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTransaction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===============================================================================
' clsTransaction — Transaction Data Object
'
' Represents a single transaction from either bank statement or DMS GL export.
' Used as the core data structure throughout the matching engine.
'===============================================================================

Option Explicit

' --- Properties ---
Public TransactionID As Long
Public Source As String          ' "BANK" or "DMS"
Public TransactionDate As Date
Public PostDate As Date
Public Description As String
Public Amount As Currency
Public CheckNumber As String
Public ReferenceNumber As String
Public TypeCode As String       ' DMS type code: CHK, DEP, CVR, ACH, WIR, FEE, etc.
Public BankSource As String     ' "BOFA" or "TRUIST"
Public GLAccount As String
Public IsMatched As Boolean
Public MatchID As Long
Public MatchType As String
Public ConfidenceScore As Double
Public SheetRow As Long         ' Row on source sheet for cross-reference

Private Sub Class_Initialize()
    TransactionID = 0
    Source = ""
    IsMatched = False
    MatchID = 0
    SheetRow = 0
    ConfidenceScore = 0
End Sub

Public Function Clone() As clsTransaction
    Dim t As New clsTransaction
    t.TransactionID = Me.TransactionID
    t.Source = Me.Source
    t.TransactionDate = Me.TransactionDate
    t.PostDate = Me.PostDate
    t.Description = Me.Description
    t.Amount = Me.Amount
    t.CheckNumber = Me.CheckNumber
    t.ReferenceNumber = Me.ReferenceNumber
    t.TypeCode = Me.TypeCode
    t.BankSource = Me.BankSource
    t.GLAccount = Me.GLAccount
    t.IsMatched = Me.IsMatched
    t.MatchID = Me.MatchID
    t.MatchType = Me.MatchType
    t.ConfidenceScore = Me.ConfidenceScore
    t.SheetRow = Me.SheetRow
    Set Clone = t
End Function

Public Function ToSummary() As String
    ToSummary = Source & " | " & Format(TransactionDate, "MM/DD/YYYY") & _
                " | " & Format(Amount, "$#,##0.00") & " | " & Description
End Function
