Attribute VB_Name = "ExportImport"
Option Explicit
Const IOEXP_UNIQUE_STRING = "zn8AiLJcRXREAfOSpY"
'Do not move the above line 2 to any other line - it's there to uniquely identify this module
'Procedures prefixed with this string will be called from the corresponding
'event procedure in the ThisWorkbook module
Const EH_PREFIX = "EH_"
Private Const WB_PREFIX = "Workbook_"

'Import tasks to do upon WB open
Public Sub wbOpenImport()
    Call loadHandlersToWB
    'Call to Import Modules must come last or you'll get wbOpenImport duplicated & ambiguous
    Call ImportModules
End Sub

'Loads the handlers defined in the EventHandler module to the ThisWorkbook module
Public Sub loadHandlersToWB()
    Dim eventHandlerModule As VBIDE.CodeModule
    Set eventHandlerModule = getEventHandlerModule
    Dim procNames As Collection: Set procNames = getAllProcNames(eventHandlerModule)
    Dim procName As Variant
    For Each procName In procNames
        If shouldLoadToWb(procName) Then
            Dim targetProcName As String: targetProcName = wbTargetProcName(procName)
            'In case the target WB function isn't there, add it with blank content
            Call maybeAddSpecialSubToThisWB(targetProcName)
            Call maybeAddCallToWbProc(procName, targetProcName)
        End If
    Next
End Sub

'Calls the procedure specified by procName from the procedure targetProcName in the WB module
Private Sub maybeAddCallToWbProc(ByVal procName As String, ByVal targetProcName As String)
    Dim wbModule As VBIDE.CodeModule: Set wbModule = getThisWorkbookModule
    Dim startLineNum As Long, currLineNum As Long, countLines As Long
    startLineNum = wbModule.ProcBodyLine(targetProcName, vbext_pk_Proc)
    countLines = wbModule.ProcCountLines(targetProcName, vbext_pk_Proc)
    Dim currLine As String
    For currLineNum = startLineNum + 1 To startLineNum + countLines - 2
        currLine = wbModule.lines(currLineNum, 1)
        'If the procedure is mentioned at all on any line, be conservative and don't add it
        If InStr(currLine, procName) <> 0 Then
            Debug.Print "Did not add a call to " & procName & ". Found an existing call on line " & _
                currLineNum & ": " & currLine
            Exit Sub
        End If
    Next
    Call wbModule.InsertLines(currLineNum, "    Call " & procName)
    Debug.Print "Inserted a call to " & procName & " on line " & currLineNum
End Sub

Sub testShouldLoadToWb()
    Dim testVal As String: testVal = "Foo"
    MsgBox (testVal & ": " & shouldLoadToWb(testVal))
    testVal = EH_PREFIX & "Bar"
    MsgBox (testVal & ": " & shouldLoadToWb(testVal))
End Sub

'Return True iff the procedure by that name should to to the ThisWorkbook module
Public Function shouldLoadToWb(ByVal procedureName As String) As Boolean
    If InStr(1, procedureName, EH_PREFIX) = 1 Then
        shouldLoadToWb = True
    Else
        shouldLoadToWb = False
    End If
End Function

'Each function in the EventHandler has a target function in the workbook to call it
Public Function wbTargetProcName(ByVal ehFnName As String) As String
    wbTargetProcName = Replace(ehFnName, EH_PREFIX, WB_PREFIX)
End Function

'Collection of Strings of Sub names in that module
Private Function getAllProcNames(module As VBIDE.CodeModule) As Collection
    Dim lineNum As Integer
    Dim procName As String
    Dim coll As New Collection
    Dim ProcKind As VBIDE.vbext_ProcKind
    With module
        lineNum = .CountOfDeclarationLines + 1
        Do Until lineNum >= .CountOfLines
            procName = .ProcOfLine(lineNum, ProcKind)
            lineNum = .ProcStartLine(procName, ProcKind) + _
                    .ProcCountLines(procName, ProcKind) + 1
            coll.Add Item:=procName
        Loop
    End With
    Set getAllProcNames = coll
End Function

Private Sub testMaybeAddSubToThisWorkbook()
    Call maybeAddSubToThisWorkbook("testSub", "", "")
    Call maybeAddSubToThisWorkbook("testSub", "", "")
End Sub

'Adds an empty-bodied sub of the given name to the WB with the right arguments.
Private Sub maybeAddSpecialSubToThisWB(subName As String)
    Dim args As String: args = ""
    Select Case subName
        Case "Workbook_BeforeClose":
            args = "Cancel As Boolean"
        Case "Workbook_BeforeSave":
            args = "ByVal SaveAsUI As Boolean, Cancel As Boolean"
        Case "Workbook_BeforePrint":
            args = "Cancel As Boolean"
        Case "Workbook_AfterSave":
            args = "ByVal Success As Boolean"
        Case "Workbook_SheetActivate":
            args = "ByVal Sh As Object"
        Case "Workbook_SheetBeforeDoubleClick":
            args = "ByVal Sh As Object, ByVal Target As Range, Cancel As Boolean"
        Case "Workbook_SheetBeforeRightClick":
            args = "ByVal Sh As Object, ByVal Target As Range, Cancel As Boolean"
        Case "Workbook_SheetChange":
            args = "ByVal Sh As Object, ByVal Target As Range"
        Case "Workbook_SheetCalculate":
            args = "ByVal Sh As Object"
        Case "Workbook_SheetSelectionChange":
            args = "ByVal Sh As Object, ByVal Target As Range"
        Case "Workbook_NewSheet":
            args = "ByVal Sh As Object"
        Case "Workbook_SheetFollowHyperlink":
            args = "ByVal Sh As Object, ByVal Target As Hyperlink"
        Case Else:
            Dim errorDesc As String: errorDesc = "Unhandled WB sub name: " & subName
            Err.Raise Number:=513, Description:=errorDesc
    End Select
    Call maybeAddSubToThisWorkbook(subName, args, "")
End Sub

Private Sub maybeAddSubToThisWorkbook(subName As String, subParams As String, subCode As String)
    Dim oCodeMod As VBIDE.CodeModule
    Set oCodeMod = getThisWorkbookModule
    Call maybeAddSubToModule(oCodeMod, subName, subParams, subCode)
End Sub

Private Sub maybeAddSubToModule(oCodeMod As VBIDE.CodeModule, subName As String, subParams As String, subCode As String)
    Dim stringToAdd As String: stringToAdd = "Private Sub " & subName & "(" & subParams & ")"
    If subCode <> "" Then _
        stringToAdd = stringToAdd & vbCrLf
    stringToAdd = stringToAdd & subCode & vbCrLf & "End Sub" & vbCrLf
    If Not doesSubExist(subName, oCodeMod) Then
        oCodeMod.AddFromString stringToAdd
        Debug.Print "Added sub: " & subName & " to workbook."
    Else
        Debug.Print "Sub " & subName & " already exists. Didn't add this sub to WB."
    End If
End Sub

Private Sub testDoesSubExist()
    Dim testVal As String: testVal = "Foo"
    MsgBox (testVal & ": " & doesSubExist(testVal, getThisWorkbookModule()))
    testVal = "Workbook_Open"
    MsgBox (testVal & ": " & doesSubExist(testVal, getThisWorkbookModule()))
End Sub

Private Function getEventHandlerModule() As VBIDE.CodeModule
    Set getEventHandlerModule = getModule("EventHandler")
End Function

Private Function getThisWorkbookModule() As VBIDE.CodeModule
    Set getThisWorkbookModule = getModule(ThisWorkbook.CodeName)
End Function

'modName should be the name of a valid code module
Private Function getModule(modName As String) As VBIDE.CodeModule
    Dim VBProj As VBIDE.VBProject
    Dim oComp As VBIDE.VBComponent
    Set VBProj = ThisWorkbook.VBProject
    Set oComp = VBProj.VBComponents(modName)
    Set getModule = oComp.CodeModule
End Function

'Does a sub with this name already exist in the modyule
Private Function doesSubExist(subName As String, oCodeMod As VBIDE.CodeModule)
    Dim startLine As Integer
    'If the sub isn't there, ProcBodyLine will throw an error- which we use as conditional logic
    On Error GoTo sub_does_not_exist
    startLine = oCodeMod.ProcBodyLine(subName, vbext_pk_Proc)
    On Error GoTo 0
    doesSubExist = True
    Exit Function
sub_does_not_exist:
    doesSubExist = False
End Function

Public Sub ExportModules()
Attribute ExportModules.VB_ProcData.VB_Invoke_Func = "p\n14"
    Dim bExport As Boolean
    Dim wkbSource As Excel.Workbook
    Dim szSourceWorkbook As String
    Dim szExportPath As String
    Dim szFileName As String
    Dim cmpComponent As VBIDE.VBComponent

    ''' The code modules will be exported in a folder named.
    ''' VBAProjectFiles in the Documents folder.
    ''' The code below create this folder if it not exist
    ''' or delete all files in the folder if it exist.
    If FolderWithVBAProjectFiles = "Error" Then
        MsgBox "Export Folder not exist"
        Exit Sub
    End If
    
    On Error Resume Next
        Dim whiteListedModule As Variant
        For Each whiteListedModule In whiteListedModules()
            Kill FolderWithVBAProjectFiles & "\" & whiteListedModule & ".*"
            Debug.Print "Deleted module: " & whiteListedModule
        Next
    On Error GoTo 0

    ''' NOTE: This workbook must be open in Excel.
    szSourceWorkbook = ActiveWorkbook.name
    Set wkbSource = Application.Workbooks(szSourceWorkbook)
    
    If wkbSource.VBProject.Protection = 1 Then
    MsgBox "The VBA in this workbook is protected," & _
        " not possible to export the code"
    Exit Sub
    End If
    
    szExportPath = FolderWithVBAProjectFiles & "\"
    
    For Each cmpComponent In wkbSource.VBProject.VBComponents
        
        bExport = True
        szFileName = cmpComponent.name

        If (Not isWhiteListed(szFileName)) Then _
            bExport = False

        ''' Concatenate the correct filename for export.
        Select Case cmpComponent.Type
            Case vbext_ct_ClassModule
                szFileName = szFileName & ".cls"
            Case vbext_ct_MSForm
                szFileName = szFileName & ".frm"
            Case vbext_ct_StdModule
                szFileName = szFileName & ".bas"
            Case vbext_ct_Document
                ''' This is a worksheet or workbook object.
                ''' Don't try to export.
                bExport = False
        End Select
        
        If bExport Then
            ''' Export the component to a text file.
            cmpComponent.Export szExportPath & szFileName
            Debug.Print "Exported " & szFileName
        
        End If
   
    Next cmpComponent

    Debug.Print "**" & Now & "** " & "Files Exported to: " & szExportPath
End Sub

Public Sub ImportModulesWarn()
    Dim answer As Integer
    answer = MsgBox("Import will overwrite the following modules with data from disk: " & _
    vbCrLf & whiteList() & vbCrLf & "Are you sure you want to proceed?", _
    vbYesNo + vbQuestion, "Import and Override?")
    If answer = vbNo Then
        Debug.Print "!!!!!!! No Import done. User cancelled."
    Else
        Call ImportModules
    End If
End Sub

Public Sub ImportModules()
    Dim wkbTarget As Excel.Workbook
    Dim objFSO As Scripting.FileSystemObject
    Dim objFile As Scripting.File
    Dim szTargetWorkbook As String
    Dim szImportPath As String
    Dim szFileName As String
    Dim cmpComponents As VBIDE.VBComponents

    'If ActiveWorkbook.name = ThisWorkbook.name Then
    '    MsgBox "Select another destination workbook" & _
    '    "Not possible to import in this workbook "
    '    Exit Sub
    'End If

    'Get the path to the folder with modules
    If FolderWithVBAProjectFiles = "Error" Then
        MsgBox "Import Folder not exist"
        Exit Sub
    End If

    ''' NOTE: This workbook must be open in Excel.
    szTargetWorkbook = ActiveWorkbook.name
    Set wkbTarget = Application.Workbooks(szTargetWorkbook)
    
    If wkbTarget.VBProject.Protection = 1 Then
    MsgBox "The VBA in this workbook is protected," & _
        "not possible to Import the code"
    Exit Sub
    End If

    ''' NOTE: Path where the code modules are located.
    szImportPath = FolderWithVBAProjectFiles & "\"
    Debug.Print "Ready to import files from: " & szImportPath
            
    Set objFSO = New Scripting.FileSystemObject
    If objFSO.GetFolder(szImportPath).Files.count = 0 Then
       MsgBox "There are no files to import"
       Exit Sub
    End If

    'Potentially rename this module so it works with import
    If (shouldRenameThisModule()) Then
        Call RenameThisModule
        Debug.Print "Renamed this module to " & IOEXP_UNIQUE_STRING & " to avoid name collision with import."
    End If
    
    'Delete all modules/Userforms from the ActiveWorkbook
    Call DeleteVBAModulesAndUserForms

    Set cmpComponents = wkbTarget.VBProject.VBComponents
    
    ''' Import all the code modules in the specified path
    ''' to the ActiveWorkbook.
    For Each objFile In objFSO.GetFolder(szImportPath).Files
        If (objFSO.GetExtensionName(objFile.name) = "cls") Or _
            (objFSO.GetExtensionName(objFile.name) = "frm") Or _
            (objFSO.GetExtensionName(objFile.name) = "bas") Then
                Dim moduleName As String
                moduleName = Split(objFile.name, ".")(0)
                If isWhiteListed(moduleName) Then
                    cmpComponents.Import objFile.path
                    Debug.Print "Imported " & objFile.name
                Else
                    Debug.Print moduleName & " not on white list. Skipped import"
                End If
        End If
    Next
    Debug.Print "************** Import complete"
    Call selectThisModule
End Sub

Function FolderWithVBAProjectFiles() As String
    Dim FSO As Object
    Set FSO = CreateObject("scripting.filesystemobject")
    Dim relativePath As String: relativePath = "VBAProjectFiles"
    Dim prefixPath As String
    
    prefixPath = ActiveWorkbook.path

    If Right(prefixPath, 1) <> "\" Then
        prefixPath = prefixPath & "\"
    End If
    
    Dim totalpath As String: totalpath = prefixPath & relativePath
    
    If FSO.FolderExists(totalpath) = False Then
        On Error Resume Next
        MkDir totalpath
        On Error GoTo 0
    End If
    
    If FSO.FolderExists(totalpath) = True Then
        FolderWithVBAProjectFiles = totalpath
    Else
        FolderWithVBAProjectFiles = "Error"
    End If
    
End Function

Sub DeleteVBAModulesAndUserForms()
    Dim VBProj As VBIDE.VBProject
    Dim VBComp As VBIDE.VBComponent
    
    Set VBProj = ActiveWorkbook.VBProject
    
    For Each VBComp In VBProj.VBComponents
        If VBComp.Type = vbext_ct_Document Then
            'Thisworkbook or worksheet module
            'We do nothing
        ElseIf isWhiteListed(VBComp.name) Then
            VBProj.VBComponents.Remove VBComp
        'We need to delete this module itself it it has been renamed
        ElseIf (VBComp.name = IOEXP_UNIQUE_STRING) Then
            VBProj.VBComponents.Remove VBComp
        End If
    Next VBComp
End Sub

Private Sub RenameThisModule()
    Dim VBComp As VBIDE.VBComponent
    Set VBComp = getThisModule()
    On Error GoTo rename_error
    VBComp.name = IOEXP_UNIQUE_STRING
    Exit Sub
rename_error:
    Dim errorDesc As String: errorDesc = "Failed to rename the current module to " & IOEXP_UNIQUE_STRING & _
        ". Does a module with that name already exist?"
    Err.Raise Number:=513, Description:=errorDesc
End Sub

Private Sub selectThisModule()
    getThisModule().Activate
End Sub

'Only rename the current module if it's on the whitelist, to avoid a collision
Private Function shouldRenameThisModule() As Boolean
    shouldRenameThisModule = isWhiteListed(getThisModule().name)
End Function

Private Function getThisModule() As VBIDE.VBComponent
    Dim VBProj As VBIDE.VBProject
    Set VBProj = ActiveWorkbook.VBProject
    Dim VBComp As VBIDE.VBComponent
    
    'Loop through all modules until you find the one whose second line contains this unique string
    For Each VBComp In VBProj.VBComponents
        If VBComp.Type = vbext_ct_Document Then
            'Thisworkbook or worksheet module
            'We do nothing
        Else
            Dim secondLine As String: secondLine = VBComp.CodeModule.lines(2, 1)
            If InStr(secondLine, IOEXP_UNIQUE_STRING) > 0 And VBComp.name <> IOEXP_UNIQUE_STRING Then
                Set getThisModule = VBComp
                Exit Function
            End If
        End If
    Next VBComp
    Call MsgBox("This module not found. Searched for module with second line equal to " & IOEXP_UNIQUE_STRING, vbExclamation)
End Function

Private Sub testIsWhiteListed()
    Dim moduleName1 As String, moduleName2 As String
    moduleName1 = "Dependencies"
    moduleName2 = "MyModule"
    MsgBox (moduleName1 & " is whitelisted: " & isWhiteListed(moduleName1))
    MsgBox (moduleName2 & " is whitelisted: " & isWhiteListed(moduleName2))
End Sub

Sub testStringInArray()
    Dim strArr() As String
    Dim unsplitArr As String
    unsplitArr = "a, b, c"
    strArr = Split(unsplitArr, ",")
    Dim str1 As String, str2 As String
    str1 = "a"
    str2 = "dog"
    Call MsgBox(str1 & " in " & unsplitArr & ": " & stringInArray(str1, strArr))
    Call MsgBox(str2 & " in " & unsplitArr & ": " & stringInArray(str2, strArr))
End Sub

Public Function stringInArray(str As String, strArr() As String) As Boolean
    Dim strLooper As Variant
    For Each strLooper In strArr
        If str = strLooper Then
            stringInArray = True
            Exit Function
        End If
    Next
    stringInArray = False
End Function

Private Function isWhiteListed(moduleName As String) As Boolean
    isWhiteListed = stringInArray(moduleName, whiteListedModules())
End Function

Private Function whiteListedModules() As String()
    whiteListedModules = Split(whiteList(), ",")
End Function

'Only modules on this list will get imported/exported
'Add your modules to the whiteList variable, separated by commas
Private Function whiteList() As String
    whiteList = "Dependencies,DependencyIndexing,DependencyIndexRun,ExportImport,FormulaChecking,GeneralPurpose" _
    & ",EventHandler"
End Function
