Attribute VB_Name = "Dependencies"
'Module for examining depedencies to/from a sheet from/to other sheets
Option Explicit

Sub showAllPrecendents()
Attribute showAllPrecendents.VB_ProcData.VB_Invoke_Func = "k\n14"
    Dim deps As Collection
    Set deps = findAllPrecedents(ActiveCell)
    Call showDents(deps, True, "All Precedents: ")
End Sub

Sub showAllDependents()
    Dim deps As Collection
    Set deps = findAllDependents(ActiveCell)
    Call showDents(deps, True, "All Dependents: ")
End Sub

Function hasPartialDents(rng As Range, internal As Boolean, precDir As Boolean)
    Dim deps As Collection
    Set deps = findDents(rng, internal, precDir, True)
    If deps.count > 0 Then
        hasPartialDents = True
    Else
        hasPartialDents = False
    End If
End Function

Sub showExternalDependents()
    Dim deps As Collection
    Set deps = findExternalDependents(ActiveCell)
    Call showDents(deps, True, "External Dependents: ")
End Sub

Sub showExternalPrecedents()
Attribute showExternalPrecedents.VB_ProcData.VB_Invoke_Func = "e\n14"
    Dim precs As Collection
    Set precs = findExternalPrecedents(ActiveCell)
    Call showDents(precs, True, "External Precedents: ")
End Sub

'external determines whether or not to print out the absolute address including workbook & worksheet
Sub showDents(dents As Collection, external As Boolean, header As String)
    Dim dent As Variant
    Dim stMsg As String
    stMsg = ""
    For Each dent In dents
        stMsg = stMsg & vbNewLine & stripWorkbookName(dent.Address(external:=external))
    Next dent
    MsgBox header & stMsg
End Sub

Function findAllPrecedents(rng As Range) As Collection
    Set findAllPrecedents = findAllDents(rng, True)
End Function

Function findAllDependents(rng As Range) As Collection
    Set findAllDependents = findAllDents(rng, False)
End Function

Function findExternalPrecedents(rng As Range) As Collection
    Set findExternalPrecedents = findExternalDents(rng, True)
End Function

Function findExternalDependents(rng As Range) As Collection
    Set findExternalDependents = findExternalDents(rng, False)
End Function

'Gives back only the dependencies that are not on the same sheet as rng
Function findInternalDents(rng As Range, precDir As Boolean) As Collection
    Set findInternalDents = findDents(rng, True, precDir, False)
End Function

'Gives back only the dependencies that are not on the same sheet as rng
Function findExternalDents(rng As Range, precDir As Boolean) As Collection
    Set findExternalDents = findDents(rng, False, precDir, False)
End Function

'this procedure finds the cells which are the direct precedents/dependents of the active cell
'If precDir is true, then we look for precedents, else we look for dependents
Function findDents(rng As Range, internal As Boolean, precDir As Boolean, stopAtFirst As Boolean) As Collection
    'Need to unhide sheets for external dependencies or the navigate arrow won't work
    Call mUnhideAll
    Dim ws As Worksheet
    Set ws = rng.Worksheet
    Dim rLast As Range, iLinkNum As Integer, iArrowNum As Integer
    Dim dents As New Collection
    Dim bNewArrow As Boolean
    'Appliciation.ScreenUpdating = False
    If precDir Then
        rng.showPrecedents
    Else
        rng.ShowDependents
    End If
    Set rLast = rng
    iArrowNum = 1
    iLinkNum = 1
    bNewArrow = True
    Do
        Do
            Application.Goto rLast
            On Error Resume Next
            ActiveCell.NavigateArrow TowardPrecedent:=precDir, ArrowNumber:=iArrowNum, LinkNumber:=iLinkNum
            If Err.Number > 0 Then Exit Do
            On Error GoTo 0
            If rLast.Address(external:=True) = ActiveCell.Address(external:=True) Then Exit Do
            bNewArrow = False
            iLinkNum = iLinkNum + 1 ' try another link
            If ((Selection.Worksheet.name <> ws.name) Xor internal) Then
                Dim addrKey As String: addrKey = Selection.Address
                If Not HasKey(dents, addrKey) Then _
                    dents.Add Item:=Selection, key:=addrKey
                If stopAtFirst Then
                    GoTo cleanupAndReturn
                End If
            End If
        Loop
        If bNewArrow Then Exit Do
        iLinkNum = 1
        bNewArrow = True
        iArrowNum = iArrowNum + 1 'try another arrow
    Loop
    
cleanupAndReturn:
    
    rLast.Parent.ClearArrows
    Application.Goto rLast
    Set findDents = dents
End Function

'If precDir is true, then we get all precedents, else we look for all dependents
Function findAllDents(rng As Range, precDir As Boolean) As Collection
    Set findAllDents = findAllDentsUnhide(rng, precDir, True)
End Function

'If precDir is true, then we get all precedents, else we look for all dependents
'If shouldUnhide is true, all sheets are unhidden before the search for dents
Function findAllDentsUnhide(rng As Range, precDir As Boolean, shouldUnhide As Boolean) As Collection
    'Need to unhide sheets for external dependencies or the navigate arrow won't work
    If shouldUnhide Then _
        Call mUnhideAll
    Dim ws As Worksheet
    Set ws = rng.Worksheet
    Dim rLast As Range, iLinkNum As Integer, iArrowNum As Integer
    Dim dents As New Collection
    Dim bNewArrow As Boolean
    'Appliciation.ScreenUpdating = False
    If precDir Then
        rng.showPrecedents
    Else
        rng.ShowDependents
    End If
    Set rLast = rng
    iArrowNum = 1
    iLinkNum = 1
    bNewArrow = True
    Do
        Do
            Application.Goto rLast
            On Error Resume Next
            ActiveCell.NavigateArrow TowardPrecedent:=precDir, ArrowNumber:=iArrowNum, LinkNumber:=iLinkNum
            If Err.Number > 0 Then Exit Do
            On Error GoTo 0
            If rLast.Address(external:=True) = ActiveCell.Address(external:=True) Then Exit Do
            bNewArrow = False
            iLinkNum = iLinkNum + 1 ' try another link
            Dim addrKey As String: addrKey = Selection.Address
            If Not HasKey(dents, addrKey) Then _
                dents.Add Item:=Selection, key:=Selection.Address
        Loop
        If bNewArrow Then Exit Do
        iLinkNum = 1
        bNewArrow = True
        iArrowNum = iArrowNum + 1 'try another arrow
    Loop
    rLast.Parent.ClearArrows
    Application.Goto rLast
    Set findAllDentsUnhide = dents
End Function

Function remo()
