Attribute VB_Name = "FormulaChecking"
Option Explicit

Const DIFF_COLOUR_INDEX = 6
Const BLANK_DIFF_COLOUR_INDEX = 4
Const COMP_FORM_PREFIX = "COMP_"
Const SUPER_COMP_FORM_PREFIX = "S"
Const CHUNK_SIZE = 10
Const STRIP_CHUNK_SIZE = 8000

'Creates a super-compressed view of the diffs only in a list on a separate shadow sheet
'Must be called with existing shadow sheet as the active sheet
Sub superCompress()
    Dim shadowSheet As Worksheet: Set shadowSheet = ActiveSheet
    Dim sheetName As String: sheetName = shadowSheet.name
    If InStr(sheetName, COMP_FORM_PREFIX) <> 1 Then
        MsgBox ("You must run this from a shadow sheet starting with " & COMP_FORM_PREFIX & vbCrLf _
            & "Sheet name detected: " & sheetName)
        Exit Sub
    End If
    Call resetShadow(sheetName, SUPER_COMP_FORM_PREFIX)
    Dim outputSheet As Worksheet
    Set outputSheet = getOutputSheet(shadowSheet.name, SUPER_COMP_FORM_PREFIX)
    outputSheet.Cells(1, 1) = "Cell"
    outputSheet.Cells(1, 2) = "Formula Above"
    outputSheet.Cells(1, 3) = "Formula in Cell"
    Dim column As Range, cell As Range
    Dim outputRow As Integer: outputRow = 2
    For Each column In shadowSheet.usedRange.Columns
        For Each cell In column.Cells
            If isHighlighted(cell) Then
                outputSheet.Cells(outputRow, 1) = getOriginalCellAddr(cell, shadowSheet)
                outputSheet.Cells(outputRow, 2) = cell.Offset(-1, 0).value
                outputSheet.Cells(outputRow, 3) = cell.value
                outputRow = outputRow + 1
            End If
        Next
    Next
    outputSheet.Select
    outputSheet.usedRange.Rows(1).Font.Bold = True
    outputSheet.usedRange.Columns.AutoFit
    With ActiveWindow
        .SplitRow = 1
        .SplitColumn = 1
        .FreezePanes = True
    End With
End Sub

Private Sub testGetOriginalCellAddr()
    Call MsgBox(getOriginalCellAddr(ActiveCell, ActiveSheet))
End Sub

'Assuming the cell is in our compressed formulas shadow sheet,
'what was the cell's relative address in the original sheet?
Private Function getOriginalCellAddr(cell As Range, shadowSheet As Worksheet) As String
    Dim shadowRow As Integer, shadowCol As Integer
    shadowRow = cell.row: shadowCol = cell.column
    Dim originalRow As String, originalCol As String
    originalRow = shadowSheet.Cells(shadowRow, 1)
    originalCol = shadowSheet.Cells(1, shadowCol)
    getOriginalCellAddr = originalCol & originalRow
End Function

Sub compressToDiffsOnly()
    compressFormulasAux (True)
End Sub

Sub compressFormulas()
    compressFormulasAux (False)
End Sub

'diffsOnly = True means we only want to see formulas that differ from the row above
Sub compressFormulasAux(diffsOnly As Boolean)
    Dim selectedRange As Range
    Set selectedRange = Selection
    Dim startTime As Double
    startTime = Now
    Debug.Print "***** Filtering unused columns *****"
    Call FilterFormulaColumns(selectedRange)
    Debug.Print "***** Unused columns Filtered *****"
    Debug.Print "***** Filtering repeat rows *****"
    Call FilterRepeatRows(selectedRange)
    Debug.Print "***** Repeat rows filtered *****"
    Dim outputSheet As Worksheet: Set outputSheet = getOutputSheet(selectedRange.Worksheet.name, COMP_FORM_PREFIX)
    Call FormatUsedRange(outputSheet)
    If diffsOnly Then
        Debug.Print "Stripping non diffs from " & outputSheet.name
        Call stripNonDiffs(outputSheet)
    End If
    Debug.Print ("Time: " & Minute(Now - startTime) & ":" & Second(Now - startTime))
End Sub

Sub testStripNonDiffs()
    Call stripNonDiffs(ActiveSheet)
End Sub

'Assuming the formulas on the outputSheet have been compressed & colour coded,
'This function deletes the contents of all cells that are not diffs
Sub stripNonDiffs(outputSheet As Worksheet)
    Dim cell As Range
    'Offset is so we don't clear the row/column info
    Dim counter As Integer: counter = 0
    For Each cell In outputSheet.usedRange.Offset(1, 1).Cells
        counter = counter + 1
        counter = counter Mod STRIP_CHUNK_SIZE
        If counter = 0 Then
            Debug.Print Now & " Stripped " & STRIP_CHUNK_SIZE & " cells"
            DoEvents
        End If
        If Not isHighlighted(cell) Then
            cell.Clear
        End If
    Next
    If counter <> 0 Then _
        Debug.Print Now & " Stripped " & counter & " cells"
End Sub

Private Function isHighlighted(cell As Range) As Boolean
    If cell.Interior.ColorIndex = DIFF_COLOUR_INDEX Or _
    cell.Interior.ColorIndex = BLANK_DIFF_COLOUR_INDEX Then
        isHighlighted = True
    Else
        isHighlighted = False
    End If
End Function

'Compresses the selected range by writing only its formula columns to another sheet and skipping non-formula columns
Sub FilterFormulaColumns(selectedRange As Range)
    Dim sheetName As String: sheetName = selectedRange.Worksheet.name
    Call resetShadow(sheetName, COMP_FORM_PREFIX)
    Dim outputSheet As Worksheet: Set outputSheet = getOutputSheet(sheetName, COMP_FORM_PREFIX)
    Dim col As Range
    Dim currOutputCol As Integer
    currOutputCol = 2
    For Each col In selectedRange.Columns
        If hasSomeFormulas(col) Then
            Dim colLetter As String
            colLetter = columnLetter(col.column)
            outputSheet.Cells(1, currOutputCol) = colLetter
            Dim cell As Range
            Dim outputRowIndex As Integer
            outputRowIndex = 2
            Application.ScreenUpdating = False
            For Each cell In col.Cells
                
                If cell.HasFormula Then
                    Dim trimmedFormula As String
                    trimmedFormula = cell.formula
                    trimmedFormula = Right(trimmedFormula, Len(trimmedFormula) - 1)
                    Dim outputCell As Range
                    Set outputCell = outputSheet.Cells(outputRowIndex, currOutputCol)
                    outputCell.value = trimmedFormula
                End If
                outputRowIndex = outputRowIndex + 1
            Next cell
            Application.ScreenUpdating = True
            currOutputCol = currOutputCol + 1
        End If
    Next col
    Dim lastRow As Integer: lastRow = outputSheet.usedRange.Rows.count
    Dim currRow As Integer
    For currRow = 2 To lastRow
        outputSheet.Cells(currRow, 1) = currRow - 2 + selectedRange.row
    Next
    Call FormatUsedRange(outputSheet)
    outputSheet.Activate
End Sub

Private Sub resetShadow(sheetName As String, prefixName As String)
    Dim newSheetName As String
    newSheetName = Left(prefixName & sheetName, 31)
    Application.StatusBar = "Resetting: " & newSheetName
    ResetOutput (newSheetName)
    Debug.Print "Reset Shadow sheet: " & newSheetName
End Sub

'Filters out repeat rows and highlights changes on rows with unique formulas
'A repeat row is a row which does not add any new formulas relative to the row above it
'This subroutine assumes the compressed formula sheet already exists with some formulas in it
Sub FilterRepeatRows(selectedRange As Range)
    Dim outputSheet As Worksheet: Set outputSheet = getOutputSheet(selectedRange.Worksheet.name, COMP_FORM_PREFIX)
    Dim rowNum As Integer
    Dim rowToDelete As Variant
    Dim rowsToDelete As New Collection
    Dim counter As Integer: counter = 0
    'Skip the first row which is a header, and skip the last row which has no rows after it
    For rowNum = 2 To outputSheet.usedRange.Rows.count - 1
        counter = counter + 1
        counter = counter Mod CHUNK_SIZE
        If counter = 0 Then
            Debug.Print Now & " Scanned " & CHUNK_SIZE & " rows"
            DoEvents
        End If
        If doesBelowRepeat(outputSheet, rowNum) Then
            rowsToDelete.Add Item:=outputSheet.usedRange.Rows(rowNum + 1)
        Else
            Call highlightDiffsBelow(outputSheet, rowNum)
            DoEvents
        End If
    Next
    If counter <> 0 Then _
        Debug.Print Now & " Scanned " & counter & " rows"
    Debug.Print "----- Scanning of rows complete -----"
    counter = 0
    For Each rowToDelete In rowsToDelete
        counter = counter + 1
        counter = counter Mod CHUNK_SIZE
        If counter = 0 Then
            Debug.Print Now & " Deleted " & CHUNK_SIZE & " rows"
            DoEvents
        End If
        rowToDelete.Delete
    Next
    If counter <> 0 Then _
        Debug.Print Now & " Deleted " & counter & " rows"
End Sub

Sub testDoesBelowRepeat()
    Dim rowNum As Integer: rowNum = ActiveCell.row
    Call MsgBox("Does row below " & rowNum & " repeat it: " & doesBelowRepeat(ActiveCell.Worksheet, rowNum))
End Sub

'Does the row below repeat all the formulas from the row above?
'Note: we ignore the leftmost column, which just contains indices
Function doesBelowRepeat(ws As Worksheet, rowAbove As Integer) As Boolean
    Dim cellAbove As Range
    Dim row As Range
    Set row = getUsedRow(ws, rowAbove).Offset(0, 1)
    Set row = row.Resize(, row.Columns.count - 1)
    For Each cellAbove In row.Cells
        If Not doesBelowRepeatFormula(cellAbove, cellAbove.Offset(1, 0)) Then
            doesBelowRepeat = False
            Exit Function
        End If
    Next
    doesBelowRepeat = True
End Function

Function getUsedRow(ws As Worksheet, rowNum As Integer) As Range
    Dim usedRange As Range: Set usedRange = ws.usedRange
    Set getUsedRow = usedRange.Rows(rowNum)
End Function

'Highlights the cells in the row below that differ from the row above by more than just autofill diffs
Sub highlightDiffsBelow(ws As Worksheet, rowAbove As Integer)
    Dim cellAbove As Range
    Dim row As Range
    Set row = getUsedRow(ws, rowAbove).Offset(0, 1)
    Set row = row.Resize(, row.Columns.count - 1)
    For Each cellAbove In row.Cells
        Dim cellBelow As Range: Set cellBelow = cellAbove.Offset(1, 0)
        If Not doesBelowRepeatFormula(cellAbove, cellBelow) Then
            If IsEmpty(cellBelow) Then
                cellBelow.Interior.ColorIndex = BLANK_DIFF_COLOUR_INDEX
            Else
                cellBelow.Interior.ColorIndex = DIFF_COLOUR_INDEX
            End If
        End If
    Next
End Sub

Function doesBelowRepeatFormula(cellAbove As Range, cellBelow As Range) As Boolean
    Dim aboveRefs() As String, belowRefs() As String
    aboveRefs = Split(ExtractCellRefs(cellAbove), ",")
    belowRefs = Split(ExtractCellRefs(cellBelow), ",")
    Dim lenAbove As Integer: lenAbove = ArrayLen(aboveRefs)
    If (lenAbove <> ArrayLen(belowRefs)) Then
        doesBelowRepeatFormula = False
        Exit Function
    End If
    Dim transformedBelowValue As String: transformedBelowValue = cellBelow.value
    'If we get to here then the lengths of the arrays match
    Dim i As Integer
    'We need this as a placeholder to remember where to start our next replacement
    Dim start As Integer: start = 1
    For i = 0 To lenAbove - 1
        Dim belowRef As String: belowRef = belowRefs(i)
        Dim aboveRef As String: aboveRef = aboveRefs(i)
        Dim belowRng As Range: Set belowRng = Range(getRelativeRef(belowRef))
        Dim aboveRng As Range: Set aboveRng = Range(getRelativeRef(aboveRef))
        'The subtraction of row indices and comparison > 1 is to account for both absolute & relative
        If (belowRng.row - aboveRng.row > 1) Or _
            (belowRng.row < aboveRng.row) Or _
            (aboveRng.column <> belowRng.column) Then
            doesBelowRepeatFormula = False
            Exit Function
        End If
        'Only replace 1 occurence, starting beyond the previous replacement
        transformedBelowValue = Left(transformedBelowValue, start - 1) & _
            Replace(transformedBelowValue, belowRef, aboveRef, start, 1)
        'Shift our start index beyond the occurence we just replaced
        start = InStr(start, transformedBelowValue, aboveRef) + Len(aboveRef)
    Next
    doesBelowRepeatFormula = cellAbove.value = transformedBelowValue
End Function

'Gives back only the relative part of the ref
Function getRelativeRef(absoluteRef As String) As String
    Dim splitStr() As String: splitStr = Split(absoluteRef, "!")
    If ArrayLen(splitStr) < 2 Then
        getRelativeRef = absoluteRef
    Else
        getRelativeRef = splitStr(1)
    End If
End Function

Sub testExtractCellRefs()
    MsgBox (ExtractCellRefs(ActiveCell))
End Sub

'Yields a comma separated string of references in the formula of the given range
Function ExtractCellRefs(Rg As Range) As String
    Dim xRetList As Object
    Dim xRegEx As Object
    Dim i As Long
    Dim xRet As String
    Application.Volatile
    Set xRegEx = CreateObject("VBSCRIPT.REGEXP")
    With xRegEx
        .Pattern = "('?[a-zA-Z0-9\s\[\]\.]{1,99})?'?!?\$?[A-Z]{1,3}\$?[0-9]{1,7}(:\$?[A-Z]{1,3}\$?[0-9]{1,7})?"
        .Global = True
        .MultiLine = True
        .IgnoreCase = False
    End With
    Set xRetList = xRegEx.Execute(Rg.formula)
    If xRetList.count > 0 Then
        For i = 0 To xRetList.count - 1
            xRet = xRet & xRetList.Item(i) & ","
        Next
        ExtractCellRefs = Left(xRet, Len(xRet) - 1)
    Else
        ExtractCellRefs = ""
    End If
End Function

Private Sub FormatUsedRange(ws As Worksheet)
    Dim outputRange As Range: Set outputRange = ws.usedRange
    outputRange.Borders.LineStyle = xlContinuous
    ws.Cells.RowHeight = 15
    ws.Cells.ColumnWidth = 8.43
    With ActiveWindow
        .SplitColumn = 1
        .SplitRow = 1
        .FreezePanes = True
    End With
End Sub

'True if the rng has at least one formula, False otherwise
Private Function hasSomeFormulas(rng As Range) As Boolean
    If rng.HasFormula = False Then
        hasSomeFormulas = False
    Else
        hasSomeFormulas = True
    End If
End Function

Private Function columnLetter(lngCol As Long) As String
    Dim vArr
    vArr = Split(Cells(1, lngCol).Address(True, False), "$")
    columnLetter = vArr(0)
End Function

Private Function getOutputSheet(sheetName As String, prefix As String) As Worksheet
    On Error GoTo problem_finding_output
    Set getOutputSheet = ActiveWorkbook.Sheets(Left(prefix & sheetName, 31))
    On Error GoTo 0
    Exit Function
problem_finding_output:
    Call MsgBox("There was a problem locating the sheet: " & _
        Left(prefix & sheetName, 31), vbExclamation)
End Function
