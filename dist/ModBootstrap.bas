Attribute VB_Name = "ModBootstrap"
'===============================================================================
' ModBootstrap — Auto-generated module importer
'
' Run 'RunBootstrap' from Immediate Window to import all VBA modules.
' Run 'ImportAndMatch' to import Honda data and run matching.
'
' AUTO-GENERATED — do not edit manually.
'===============================================================================

Option Explicit

Public Sub RunBootstrap()
    ' Import all VBA modules and class files into this workbook.
    Dim vbProj As Object
    Set vbProj = ThisWorkbook.VBProject
    
    Dim imported As Long
    imported = 0
    
    On Error Resume Next
    
    ' --- Standard Modules ---
    RemoveIfExists vbProj, "ModAuditTrail"
    vbProj.VBComponents.Import "/Users/ChrisJody/Projects/ABR/src/vba/modules/ModAuditTrail.bas"
    imported = imported + 1
    Debug.Print "  Imported: ModAuditTrail.bas"

    RemoveIfExists vbProj, "ModConfig"
    vbProj.VBComponents.Import "/Users/ChrisJody/Projects/ABR/src/vba/modules/ModConfig.bas"
    imported = imported + 1
    Debug.Print "  Imported: ModConfig.bas"

    RemoveIfExists vbProj, "ModExport"
    vbProj.VBComponents.Import "/Users/ChrisJody/Projects/ABR/src/vba/modules/ModExport.bas"
    imported = imported + 1
    Debug.Print "  Imported: ModExport.bas"

    RemoveIfExists vbProj, "ModHelpers"
    vbProj.VBComponents.Import "/Users/ChrisJody/Projects/ABR/src/vba/modules/ModHelpers.bas"
    imported = imported + 1
    Debug.Print "  Imported: ModHelpers.bas"

    RemoveIfExists vbProj, "ModImportBank"
    vbProj.VBComponents.Import "/Users/ChrisJody/Projects/ABR/src/vba/modules/ModImportBank.bas"
    imported = imported + 1
    Debug.Print "  Imported: ModImportBank.bas"

    RemoveIfExists vbProj, "ModImportDMS"
    vbProj.VBComponents.Import "/Users/ChrisJody/Projects/ABR/src/vba/modules/ModImportDMS.bas"
    imported = imported + 1
    Debug.Print "  Imported: ModImportDMS.bas"

    RemoveIfExists vbProj, "ModMain"
    vbProj.VBComponents.Import "/Users/ChrisJody/Projects/ABR/src/vba/modules/ModMain.bas"
    imported = imported + 1
    Debug.Print "  Imported: ModMain.bas"

    RemoveIfExists vbProj, "ModMatchCVR"
    vbProj.VBComponents.Import "/Users/ChrisJody/Projects/ABR/src/vba/modules/ModMatchCVR.bas"
    imported = imported + 1
    Debug.Print "  Imported: ModMatchCVR.bas"

    RemoveIfExists vbProj, "ModMatchEngine"
    vbProj.VBComponents.Import "/Users/ChrisJody/Projects/ABR/src/vba/modules/ModMatchEngine.bas"
    imported = imported + 1
    Debug.Print "  Imported: ModMatchEngine.bas"

    RemoveIfExists vbProj, "ModOutstanding"
    vbProj.VBComponents.Import "/Users/ChrisJody/Projects/ABR/src/vba/modules/ModOutstanding.bas"
    imported = imported + 1
    Debug.Print "  Imported: ModOutstanding.bas"

    RemoveIfExists vbProj, "ModResetAndRerun"
    vbProj.VBComponents.Import "/Users/ChrisJody/Projects/ABR/src/vba/modules/ModResetAndRerun.bas"
    imported = imported + 1
    Debug.Print "  Imported: ModResetAndRerun.bas"

    RemoveIfExists vbProj, "ModStagingManager"
    vbProj.VBComponents.Import "/Users/ChrisJody/Projects/ABR/src/vba/modules/ModStagingManager.bas"
    imported = imported + 1
    Debug.Print "  Imported: ModStagingManager.bas"

    ' --- Class Modules ---
    ' --- Class Modules (create as proper class, inject code) ---
    ImportClassModule vbProj, "clsMatchGroup", "/Users/ChrisJody/Projects/ABR/src/vba/classes/clsMatchGroup.cls"
    imported = imported + 1
    Debug.Print "  Imported class: clsMatchGroup.cls"

    ImportClassModule vbProj, "clsMatchResult", "/Users/ChrisJody/Projects/ABR/src/vba/classes/clsMatchResult.cls"
    imported = imported + 1
    Debug.Print "  Imported class: clsMatchResult.cls"

    ImportClassModule vbProj, "clsTransaction", "/Users/ChrisJody/Projects/ABR/src/vba/classes/clsTransaction.cls"
    imported = imported + 1
    Debug.Print "  Imported class: clsTransaction.cls"

    On Error GoTo 0

    Debug.Print ""
    Debug.Print "Bootstrap complete: " & imported & " modules imported."
    Debug.Print ""
    Debug.Print "Next: run ImportAndMatch to process Honda data."
End Sub

Private Sub RemoveIfExists(ByVal vbProj As Object, ByVal moduleName As String)
    Dim comp As Object
    On Error Resume Next
    Set comp = vbProj.VBComponents(moduleName)
    If Not comp Is Nothing Then
        If comp.Name <> "ModBootstrap" Then
            vbProj.VBComponents.Remove comp
        End If
    End If
    On Error GoTo 0
End Sub

Private Sub ImportClassModule(ByVal vbProj As Object, ByVal className As String, ByVal filePath As String)
    ' Create a proper class module and inject code from a .cls file.
    ' This avoids the VERSION 1.0 CLASS header requirement on macOS.
    RemoveIfExists vbProj, className

    ' Read the .cls file contents
    Dim fileNum As Integer
    fileNum = FreeFile
    Open filePath For Input As #fileNum
    Dim rawCode As String
    rawCode = Input$(LOF(fileNum), fileNum)
    Close #fileNum

    ' Create a new CLASS module (not standard module)
    Dim newClass As Object
    Set newClass = vbProj.VBComponents.Add(2)  ' 2 = vbext_ct_ClassModule
    newClass.Name = className

    ' Inject the code
    newClass.CodeModule.AddFromString rawCode
End Sub

Public Sub ImportAndMatch()
    ' Import Honda data files using direct paths.
    ' macOS sandbox: GrantAccessToMultipleFiles requests permission.
    Dim bankFile As String
    Dim febFile As String
    Dim janFile As String

    bankFile = "/Users/ChrisJody/Projects/ABR/docs/Honda Bkstmt.csv"
    febFile = "/Users/ChrisJody/Projects/ABR/docs/honda Feb GL.xlsx"
    janFile = "/Users/ChrisJody/Projects/ABR/docs/honda Jan GL.xlsx"

    ' Request sandbox access
    Dim filePaths(0 To 2) As String
    filePaths(0) = bankFile
    filePaths(1) = febFile
    filePaths(2) = janFile
    GrantAccessToMultipleFiles filePaths

    Dim bankCount As Long
    Dim febCount As Long
    Dim janCount As Long

    Application.StatusBar = "ABR: Importing bank statement..."
    bankCount = ModImportBank.ImportBankStatement(bankFile)

    Application.StatusBar = "ABR: Importing Feb GL..."
    febCount = ModImportDMS.ImportDMSExport(febFile)

    Application.StatusBar = "ABR: Importing Jan GL..."
    janCount = ModImportDMS.ImportDMSExport(CStr(janFile))
    
    If bankCount = 0 Then
        MsgBox "Bank import returned 0. Check file format.", vbExclamation
        Exit Sub
    End If
    
    Application.StatusBar = "ABR: Running matching engine..."
    ModMatchEngine.RunMatching ModImportBank.LoadBankTransactions(), ModImportDMS.LoadDMSTransactions()
    
    Application.StatusBar = "ABR: Auto-accepting..."
    ModStagingManager.AcceptAllHighConfidence
    
    Application.StatusBar = False
    MsgBox "Done!" & vbCrLf & vbCrLf & _
        "Bank: " & bankCount & " transactions" & vbCrLf & _
        "Feb GL: " & febCount & " transactions" & vbCrLf & _
        "Jan GL: " & janCount & " transactions" & vbCrLf & vbCrLf & _
        "Check StagedMatches and Reconciled sheets.", _
        vbInformation, "ABR Complete"
End Sub