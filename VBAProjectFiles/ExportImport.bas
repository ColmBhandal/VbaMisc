Attribute VB_Name = "ExportImport"
Option Explicit
Const IOEXP_UNIQUE_STRING = "zn8AiLJcRXREAfOSpY"
'Do not move the above line 2 to any other line - it's there to uniquely identify this module

Public Sub ExportModules()
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

    Debug.Print "************** Files Exported to: " & szExportPath
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

Private Sub ImportModules()
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
