Attribute VB_Name = "GeneralPurpose"
Function HasKey(coll As Collection, strKey As String) As Boolean
    Dim var As Variant
    On Error Resume Next
    var = coll(strKey)
    HasKey = (Err.Number = 0)
    Err.Clear
End Function

'Gets the address of this range including the sheet
Function sheetCellAddress(rng As Range) As String
    sheetCellAddress = stripWorkbookName(rng.Address(external:=True))
End Function

Function stripWorkbookName(rangeAddress As String) As String
    Dim arrayOfStrings: arrayOfStrings = Split(rangeAddress, "]")
    If (ArrayLen(arrayOfStrings) < 2) Then
        stripWorkbookName = rangeAddress
        Exit Function
    End If
    stripWorkbookName = arrayOfStrings(1)
End Function

'The range including the given cell and all contiguous non-blanks below it
Function getContiguousConstsDown(cell As Range) As Range
    Dim directlyBelow As Range: Set directlyBelow = cell.Offset(1, 0)
    If IsEmpty(directlyBelow) Then
        Set getContiguousConstsDown = cell
        Exit Function
    End If
    Set getContiguousConstsDown = Range(cell, cell.End(xlDown))
End Function

'Creates the output sheet or clears contents if it already exists
Sub ResetOutput(sheetName As String)
    If WorksheetExists(sheetName) Then
        ClearContents (sheetName)
    Else
        Call CreateOutputSheet(sheetName)
    End If
End Sub

Sub CreateOutputSheet(sheetName As String)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets.Add(After:= _
             ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
    ws.name = sheetName
End Sub

Function WorksheetExists(sName As String) As Boolean
    WorksheetExists = Evaluate("ISREF('" & sName & "'!A1)")
End Function

Sub ClearContents(sheetName As String)
    Application.Calculation = xlManual
    Application.ScreenUpdating = False
    Sheets(sheetName).Cells.ClearContents
    Sheets(sheetName).Cells.ClearFormats
    Application.ScreenUpdating = True
End Sub

Public Function ArrayLen(arr As Variant) As Integer
    ArrayLen = UBound(arr) - LBound(arr) + 1
End Function
