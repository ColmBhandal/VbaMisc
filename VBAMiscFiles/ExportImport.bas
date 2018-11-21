Attribute VB_Name = "ExportImport"
Option Explicit
Const EXPIMP_UNIQUE_STRING = "zn8AiLJcRXREAfOSpY"
'Do not move the above line 2 to any other line - it's there to uniquely identify this module
'Procedures prefixed with this string will be called from the corresponding
'event procedure in the ThisWorkbook module
Const EH_PREFIX = "EH_"
Private Const WB_PREFIX = "Workbook_"
Private Const CONFIG_FILE_NAME = "VbaMisc.config"
Private Const SPECIFIC_WHITELIST_FILE_NAME = "specificWhiteList.config"
Private Const MISC_REL_KEY = "miscRel: "
Private Const MISC_ABS_KEY = "miscAbs: "
Private Const SPECIFIC_REL_KEY = "specificRel: "
Private Const SPECIFIC_ABS_KEY = "specificAbs: "

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
    countLines = countToProcEnd(wbModule, startLineNum)
    Dim currLine As String
    For currLineNum = startLineNum + 1 To startLineNum + countLines - 1
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

Sub testCounttoProcEnd()
    Dim lineNum As Long: lineNum = 1
    Dim wbModule As VBIDE.CodeModule: Set wbModule = getThisWorkbookModule()
    Dim count As Long: count = countToProcEnd(wbModule, lineNum)
    MsgBox ("# lines to end in " & wbModule.Parent.name & " from line " & lineNum & " = " & count)
End Sub

Function countToProcEnd(wbModule As VBIDE.CodeModule, ByVal startLineNum As Long)
    Dim currLineNum As Long: currLineNum = startLineNum
    Dim currLine As String: currLine = wbModule.lines(currLineNum, 1)
    Do While Not isEndProcLine(currLine)
        currLineNum = currLineNum + 1
        currLine = wbModule.lines(currLineNum, 1)
    Loop
    countToProcEnd = currLineNum - startLineNum
End Function

Function isEndProcLine(line As String)
    isEndProcLine = False
    If InStr(line, "End Sub") <> 0 Then isEndProcLine = True
    If InStr(line, "End Function") <> 0 Then isEndProcLine = True
End Function

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
        Case "Workbook_Open":
            args = ""
        Case "Workbook_WindowActivate":
            args = "ByVal Wn As Window"
        Case "Workbook_WindowDeactivate":
            args = "ByVal Wn As Window"
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
    Debug.Print "----" & Now & "---- " & "Exports starting"
    ExportMiscModules
    Debug.Print "----" & Now & "---- " & "All exports complete"
End Sub

Private Sub ExportMiscModules()
    Dim exportFolder As String: exportFolder = createFolderWithVBAMiscFiles
    Dim whiteList() As String: whiteList = miscWhiteList()
    Call ExportModulesTargeted(exportFolder, whiteList)
End Sub

Private Sub ExportModulesTargeted(exportFolder As String, whiteList() As String)
Attribute ExportModulesTargeted.VB_ProcData.VB_Invoke_Func = "p\n14"
    Dim bExport As Boolean
    Dim wkbSource As Excel.Workbook
    Dim szSourceWorkbook As String
    Dim szExportPath As String
    Dim szFileName As String
    Dim cmpComponent As VBIDE.VBComponent

    Debug.Print "**" & Now & "** " & "Ready for export to: " & exportFolder

    ''' NOTE: This workbook must be open in Excel.
    szSourceWorkbook = ActiveWorkbook.name
    Set wkbSource = Application.Workbooks(szSourceWorkbook)
    
    If wkbSource.VBProject.Protection = 1 Then
    MsgBox "The VBA in this workbook is protected," & _
        " not possible to export the code"
    Exit Sub
    End If
    
    ''' The code modules will be exported in a folder named.
    ''' VBAProjectFiles in the Documents folder.
    ''' The code below create this folder if it not exist
    ''' or delete all files in the folder if it exist.
    If exportFolder = "Error" Then
        MsgBox "Export Folder not exist"
        Exit Sub
    End If
    
    szExportPath = exportFolder & "\"
    
    On Error Resume Next
        Dim whiteListedModule As Variant
        For Each whiteListedModule In whiteList()
            Kill szExportPath & whiteListedModule & ".*"
            Debug.Print "Deleted module: " & whiteListedModule
        Next
    On Error GoTo 0
    
    For Each cmpComponent In wkbSource.VBProject.VBComponents
        
        bExport = True
        szFileName = cmpComponent.name

        If (Not isWhiteListed(szFileName, whiteList)) Then _
            bExport = False

        ''' Concatenate the correct filename for export.
        Select Case cmpComponent.Type
            Case vbext_ct_ClassModule
                szFileName = szFileName & ".cls"
            Case vbext_ct_MSForm
                szFileName = szFileName & ".frm"
            Case vbext_ct_StdModule
                szFileName = szFileName & ".bas"
        End Select
        
        If bExport Then
            ''' Export the component to a text file.
            cmpComponent.Export szExportPath & szFileName
            Debug.Print "Exported " & szFileName
        
        End If
   
    Next cmpComponent
    
    Debug.Print "**" & Now & "** " & "Completed export to: " & exportFolder
End Sub

Public Sub ImportModulesWarn()
    Dim answer As Integer
    answer = MsgBox("Import will overwrite the following modules with data from disk: " & _
    vbCrLf & miscRawWhiteList() & vbCrLf & "Are you sure you want to proceed?", _
    vbYesNo + vbQuestion, "Import and Override?")
    If answer = vbNo Then
        Debug.Print "!!!!!!! No Import done. User cancelled."
    Else
        Call ImportModules
    End If
End Sub

Public Sub ImportModules()
    Debug.Print "----" & Now & "---- " & "Imports starting"
    Call ImportMiscModules
    Debug.Print "----" & Now & "---- " & "All Imports complete"
End Sub

Public Sub ImportMiscModules()
    Call ImportModulesTargeted(createFolderWithVBAMiscFiles, miscWhiteList)
End Sub

Public Sub ImportModulesTargeted(importFolder As String, whiteList() As String)
    Dim wkbTarget As Excel.Workbook
    Dim objFSO As Scripting.FileSystemObject
    Dim objFile As Scripting.File
    Dim szTargetWorkbook As String
    Dim szImportPath As String
    Dim szFileName As String
    Dim cmpComponents As VBIDE.VBComponents

    'Get the path to the folder with modules
    If importFolder = "Error" Then
        MsgBox "Problem with import/export folder. Quitting."
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
    szImportPath = importFolder & "\"
    Debug.Print "Ready to import files from: " & szImportPath
            
    Set objFSO = New Scripting.FileSystemObject
    If objFSO.GetFolder(szImportPath).Files.count = 0 Then
       Debug.Print "There are no files to import"
       Exit Sub
    End If
        
    'Delete whitelisted all modules/Userforms from the ActiveWorkbook
    Call DeleteVBAModulesAndUserForms(whiteList)

    'We need to rename some modules that are being used during import to avoid collisions
    Call RenameMetaModules(whiteList)
    
    Set cmpComponents = wkbTarget.VBProject.VBComponents
    
    ''' Import whitelisted code modules in the specified path
    ''' to the ActiveWorkbook.
    For Each objFile In objFSO.GetFolder(szImportPath).Files
        If (objFSO.GetExtensionName(objFile.name) = "cls") Or _
            (objFSO.GetExtensionName(objFile.name) = "frm") Or _
            (objFSO.GetExtensionName(objFile.name) = "bas") Then
                Dim moduleName As String
                moduleName = Split(objFile.name, ".")(0)
                If isWhiteListed(moduleName, whiteList) Then
                    cmpComponents.Import objFile.path
                    Debug.Print "Imported " & objFile.name
                Else
                    Debug.Print moduleName & " not on white list. Skipped import"
                End If
        End If
    Next
    Call selectMetaModule(EXPIMP_UNIQUE_STRING)
    Debug.Print "**" & Now & "** " & "Completed import from: " & importFolder
End Sub

Sub testCreateFolderWithVBAMiscFiles()
    MsgBox (createFolderWithVBAMiscFiles())
    MsgBox (createFolderWithProjectSpecificVBAFiles())
End Sub

Function createFolderWithProjectSpecificVBAFiles() As String
    Dim fso As New FileSystemObject
    Dim totalPath As String: totalPath = getFolderWithProjectSpecificVbaFiles(fso)
    createFolderWithProjectSpecificVBAFiles = createFolderWithVBAFiles(fso, totalPath)
End Function

Function createFolderWithVBAMiscFiles() As String
    Dim fso As New FileSystemObject
    Dim totalPath As String: totalPath = getFolderWithVbaMiscFiles(fso)
    createFolderWithVBAMiscFiles = createFolderWithVBAFiles(fso, totalPath)
End Function

Function createFolderWithVBAFiles(fso As FileSystemObject, totalPath As String) As String
    If (totalPath = "") Then
        totalPath = getWorkingDirPath() & "VBAMiscFiles"
    End If
    
    If Not fso.FolderExists(totalPath) Then
        On Error Resume Next
        Debug.Print "No folder exists. Attempting to make: " & totalPath
        MkDir totalPath
        On Error GoTo 0
    End If
    
    If fso.FolderExists(totalPath) = True Then
        createFolderWithVBAFiles = totalPath
    Else
        createFolderWithVBAFiles = "Error"
    End If
    
End Function

Sub testGetFolderWithVBAFiles()
    MsgBox (getFolderWithVbaMiscFiles(New FileSystemObject))
    MsgBox (getFolderWithProjectSpecificVbaFiles(New FileSystemObject))
End Sub

Function getFolderWithProjectSpecificVbaFiles(fso As FileSystemObject) As String
    getFolderWithProjectSpecificVbaFiles = getFolderWithVbaFiles(fso, SPECIFIC_REL_KEY, SPECIFIC_ABS_KEY)
End Function

Function getFolderWithVbaMiscFiles(fso As FileSystemObject) As String
    getFolderWithVbaMiscFiles = getFolderWithVbaFiles(fso, MISC_REL_KEY, MISC_ABS_KEY)
End Function

Function getFolderWithVbaFiles(fso As FileSystemObject, relKey As String, absKey As String) As String
    If (fso.FileExists(getConfigFileFullPath())) Then
        Dim textStream As textStream: Set textStream = getConfigInputStream(fso)
        Do While (Not textStream.AtEndOfLine)
            Dim currLine As String: currLine = textStream.ReadLine
            Dim val As String
            If InStr(currLine, relKey) = 1 Then
                val = Replace(currLine, relKey, "", 1, 1)
                getFolderWithVbaFiles = getWorkingDirPath() & val
                Exit Function
            ElseIf InStr(currLine, absKey) = 1 Then
                val = Replace(currLine, absKey, "", 1, 1)
                getFolderWithVbaFiles = val
                Exit Function
            End If
        Loop
    End If
    getFolderWithVbaFiles = ""
End Function

Sub testGetInputStream()
    MsgBox ("Config file empty: " & getConfigInputStream(New FileSystemObject).AtEndOfStream)
    MsgBox ("Whitelist file empty: " & getSpecificWhitelistInputStream(New FileSystemObject).AtEndOfStream)
End Sub

Function getConfigInputStream(fso As FileSystemObject) As textStream
    Dim fullPath As String: fullPath = getConfigFileFullPath()
    Set getConfigInputStream = getInputStream(fso, fullPath)
End Function

Function getConfigFileFullPath() As String
    getConfigFileFullPath = getWorkingDirPath() & CONFIG_FILE_NAME
End Function

Sub DeleteVBAModulesAndUserForms(whiteList() As String)
    Dim VBProj As VBIDE.VBProject
    Dim VBComp As VBIDE.VBComponent
    
    Set VBProj = ActiveWorkbook.VBProject
    
    For Each VBComp In VBProj.VBComponents
        Dim name As String: name = VBComp.name
        If isWhiteListed(name, whiteList) Then
            VBProj.VBComponents.Remove VBComp
            Debug.Print "Removed module " & name
        'We need to delete special meta modules
        ElseIf (name = EXPIMP_UNIQUE_STRING) Or (name = EH_UNIQUE_STRING) Then
            VBProj.VBComponents.Remove VBComp
            Debug.Print "Removed special temp module " & name
        End If
    Next VBComp
End Sub

Private Sub RenameMetaModules(whiteList() As String)
    Call RenameMetaModule(EXPIMP_UNIQUE_STRING, whiteList)
    Call RenameMetaModule(EH_UNIQUE_STRING, whiteList)
End Sub

Private Sub RenameMetaModule(uniqueIdentifier As String, whiteList() As String)
    If Not isWhiteListed(getMetaModule(uniqueIdentifier).name, whiteList) Then Exit Sub
    Dim VBComp As VBIDE.VBComponent
    Set VBComp = getMetaModule(uniqueIdentifier)
    Dim oldName As String: oldName = VBComp.name
    On Error GoTo rename_error
    VBComp.name = uniqueIdentifier
    On Error GoTo 0
    Debug.Print "Renamed " & oldName & " to " & uniqueIdentifier
    Exit Sub
rename_error:
    Dim errorDesc As String: errorDesc = "Failed to rename the module " & oldName & "to " & uniqueIdentifier & _
        ". Does a module with that name already exist?"
    Err.Raise Number:=513, Description:=errorDesc
End Sub

Private Sub selectMetaModule(uniqueIdentifier As String)
    getMetaModule(uniqueIdentifier).Activate
End Sub

Private Function getMetaModule(uniqueIdentifier As String) As VBIDE.VBComponent
    Dim VBProj As VBIDE.VBProject
    Set VBProj = ActiveWorkbook.VBProject
    Dim VBComp As VBIDE.VBComponent
    
    'Loop through all modules until you find the one whose second line contains this unique string
    For Each VBComp In VBProj.VBComponents
        Dim secondLine As String: secondLine = VBComp.CodeModule.lines(2, 1)
        If InStr(secondLine, uniqueIdentifier) > 0 And VBComp.name <> uniqueIdentifier Then
            Set getMetaModule = VBComp
            Exit Function
        End If
    Next VBComp
    Call MsgBox("This module not found. Searched for module with second line equal to " & uniqueIdentifier, vbExclamation)
End Function

Private Sub testIsWhiteListed()
    Dim moduleName1 As String, moduleName2 As String
    moduleName1 = "Dependencies"
    moduleName2 = "MyModule"
    MsgBox (moduleName1 & " is whitelisted: " & isWhiteListed(moduleName1, miscWhiteList()))
    MsgBox (moduleName2 & " is whitelisted: " & isWhiteListed(moduleName2, miscWhiteList()))
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

Private Function isWhiteListed(moduleName As String, whiteList() As String) As Boolean
    isWhiteListed = stringInArray(moduleName, whiteList)
End Function

Private Function miscWhiteList() As String()
    miscWhiteList = Split(miscRawWhiteList(), ",")
End Function

Private Function miscRawWhiteList() As String
    miscRawWhiteList = "Dependencies,DependencyIndexing,DependencyIndexRun,ExportImport,FormulaChecking,GeneralPurpose" _
    & ",EventHandler"
End Function

Private Sub testSpecificWhiteList()
    Dim elem As Variant
    Dim all As String: all = ""
    For Each elem In specificWhiteList()
        all = all & vbCrLf & elem
    Next
    MsgBox ("Specific white list: " & all)
End Sub

Private Function specificWhiteList() As String()
    Dim fso As New FileSystemObject
    Dim specTs As textStream: Set specTs = getSpecificWhitelistInputStream(fso)
    If specTs.AtEndOfStream Then
        'Return an empty array here
        Dim arrayToReturn() As String
        specificWhiteList = arrayToReturn
        Exit Function
    End If
    specificWhiteList = Split(specTs.ReadAll, vbNewLine)
End Function

Function getSpecificWhitelistInputStream(fso As FileSystemObject) As textStream
    Dim fullPath As String: fullPath = getSpecificWhitelistFileFullPath()
    Set getSpecificWhitelistInputStream = getInputStream(fso, fullPath)
End Function

Private Function getSpecificWhitelistFileFullPath() As String
    getSpecificWhitelistFileFullPath = getWorkingDirPath() & SPECIFIC_WHITELIST_FILE_NAME
End Function

Function getInputStream(fso As FileSystemObject, fullPath As String) As textStream
    On Error GoTo config_textStream_Error
    Set getInputStream = fso.OpenTextFile(fullPath)
    On Error GoTo 0
    Exit Function
config_textStream_Error:
    Dim errorDesc As String: errorDesc = "Failed to connect text stream to file: " & fullPath
    Err.Raise Number:=513, Description:=errorDesc
End Function


Sub testGetWorkingDirPath()
    MsgBox (getWorkingDirPath())
End Sub

Function getWorkingDirPath()
    Dim prefixPath As String
    prefixPath = ActiveWorkbook.path
    If Right(prefixPath, 1) <> "\" Then
        prefixPath = prefixPath & "\"
    End If
    getWorkingDirPath = prefixPath
End Function
