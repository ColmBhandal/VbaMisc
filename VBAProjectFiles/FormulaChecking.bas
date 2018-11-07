Attribute VB_Name = "FormulaChecking"
Option Explicit

Const DIFF_COLOUR_INDEX = 6
Const COMPRESSED_FORMULA_SHEET = "CompressedFormulas"
Const CHUNK_SIZE = 10

Sub compressFormulas()
    Dim startTime As Double
    startTime = Now
    Debug.Print "***** Filtering unused columns *****"
    Call FilterFormulaColumns
    Debug.Print "***** Unused columns Filtered *****"
    Debug.Print "***** Filtering repeat rows *****"
    Call FilterRepeatRows
    Debug.Print "***** Repeat rows filtered *****"
    Debug.Print ("Time: " & Minute(Now - startTime) & ":" & Second(Now - startTime))
End Sub

'Compresses the selected range by writing only its formula columns to another sheet and skipping non-formula columns
Sub FilterFormulaColumns()
    Dim selectedRange As Range
    Set selectedRange = Selection
    Call ResetOutput(COMPRESSED_FORMULA_SHEET)
    Dim outputSheet As Worksheet: Set outputSheet = getOutputSheet()
    Dim col As Range
    Dim currOutputCol As Integer
    currOutputCol = 2
    For Each col In selectedRange.Columns
        If hasSomeFormulas(col) Then
            Dim colLetter As String
            colLetter = columnLetter(col.Column)
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
    Dim lastRow As Integer: lastRow = outputSheet.usedRange.rows.count
    Dim currRow As Integer
    For currRow = 2 To lastRow
        outputSheet.Cells(currRow, 1) = currRow - 2 + selectedRange.row
    Next
    Call BordersAroundUsedRange(outputSheet)
    outputSheet.Activate
End Sub

'Filters out repeat rows and highlights changes on rows with unique formulas
'A repeat row is a row which does not add any new formulas relative to the row above it
'This subroutine assumes the compressed formula sheet already exists with some formulas in it
Sub FilterRepeatRows()
    Dim outputSheet As Worksheet: Set outputSheet = getOutputSheet()
    Dim rowNum As Integer
    Dim rowToDelete As Variant
    Dim rowsToDelete As New Collection
    'Skip the first row which is a header, and skip the last row which has no rows after it
    For rowNum = 2 To outputSheet.usedRange.rows.count - 1
        If doesBelowRepeat(outputSheet, rowNum) Then
            rowsToDelete.Add Item:=outputSheet.usedRange.rows(rowNum + 1)
        Else
            Call highlightDiffsBelow(outputSheet, rowNum)
            DoEvents
        End If
    Next
    Dim counter As Integer: counter = 0
    For Each rowToDelete In rowsToDelete
        counter = counter + 1
        counter = counter Mod CHUNK_SIZE
        If counter = 0 Then
            Debug.Print "Deleted " & CHUNK_SIZE & " rows"
            DoEvents
        End If
        rowToDelete.Delete
    Next
    If counter <> 0 Then _
        Debug.Print "Deleted " & counter & " rows"
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
    Set getUsedRow = usedRange.rows(rowNum)
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
            cellBelow.Interior.ColorIndex = DIFF_COLOUR_INDEX
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
            (aboveRng.Column <> belowRng.Column) Then
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

Private Sub BordersAroundUsedRange(ws As Worksheet)
    Dim outputRange As Range
    Set outputRange = ws.usedRange
    outputRange.Borders.LineStyle = xlContinuous
    ws.Cells.RowHeight = 15
    ws.Cells.ColumnWidth = 8.43
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

Private Function getOutputSheet() As Worksheet
    On Error GoTo problem_finding_output
    Set getOutputSheet = ActiveWorkbook.Sheets(COMPRESSED_FORMULA_SHEET)
    On Error GoTo 0
    Exit Function
problem_finding_output:
    Call MsgBox("There was a problem locating the sheet: " & COMPRESSED_FORMULA_SHEET, vbExclamation)
End Function

'TODO: Get original relative address from a range in the compressed sheet- go to first cell in that column/row & take numbers written there
'Test the above with a test sub that displays a message box with original cell's address for each compressed cell
