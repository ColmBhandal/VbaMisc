Attribute VB_Name = "FormulaChecking"
Option Explicit

Const COMPRESSED_FORMULA_SHEET = "CompressedFormulas"

'TODO: Master function = filter columns then rows. Make the time logging better.

'Compresses the selected range by writing only its formula columns to another sheet and skipping non-formula columns
Sub FilterFormulaColumns()
Attribute FilterFormulaColumns.VB_ProcData.VB_Invoke_Func = "k\n14"
    Dim startTime As Double
    startTime = Now
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
        outputSheet.Cells(currRow, 1) = currRow - 1
    Next
    Call BordersAroundUsedRange(outputSheet)
    outputSheet.Activate
    Debug.Print ("Time: " & Minute(Now - startTime) & ":" & Second(Now - startTime))
End Sub

'Filters out repeat rows and highlights changes on rows with unique formulas
'A repeat row is a row which does not add any new formulas relative to the row above it
'This subroutine assumes the compressed formula sheet already exists with some formulas in it
Sub FilterRepeatRows()
    Dim outputSheet As Worksheet: Set outputSheet = getOutputSheet()
    Dim row As Range
    Dim rowToDelete As Variant
    Dim rowsToDelete As New Collection
    For Each row In outputSheet.usedRange.rows
        Dim rowBelow As Range: Set rowBelow = row.Offset(1, 0)
        If isRepeat(row, rowBelow) Then
            Set rowToDelete = rowBelow
            rowsToDelete.Add Item:=rowToDelete
        End If
    Next
    For Each rowToDelete In rowsToDelete
        rowToDelete.Delete
        'TODO: Maybe reduce the number of calls to DoEvents- use chunking: inner & outer loop & only do in outer
        DoEvents
    Next
End Sub

Sub testIsRepeat()
    Dim rowNum As Integer: rowNum = ActiveCell.row
    Call MsgBox("Does row below " & rowNum & " repeat it: " & isRepeat(ActiveCell.Worksheet, rowNum))
End Sub

'Does the row below repeat all the formulas from the row above?
'Note: we ignore the leftmost column, which just contains indices
Function isRepeat(ws As Worksheet, rowAbove As Integer) As Boolean
    Dim cellAbove As Range
    Dim row As Range
    Set row = getUsedRow(ws, rowAbove).Offset(0, 1)
    Set row = row.Resize(, row.Columns.count - 1)
    For Each cellAbove In row.Cells
        If Not isRepeatFormula(cellAbove, cellAbove.Offset(1, 0)) Then
            isRepeat = False
            Exit Function
        End If
    Next
    isRepeat = True
End Function

Function getUsedRow(ws As Worksheet, rowNum As Integer) As Range
    Dim usedRange As Range: Set usedRange = ws.usedRange
    Set getUsedRow = usedRange.rows(rowNum)
End Function

'Highlights the cells in the row below that differ from the row above by more than just autofill diffs
Sub highlightTrueDiffs(rowAbove As Range)
    'TODO
End Sub

Function isRepeatFormula(cellAbove As Range, cellBelow As Range) As Boolean
    Dim aboveRefs() As String, belowRefs() As String
    aboveRefs = Split(ExtractCellRefs(cellAbove), ",")
    belowRefs = Split(ExtractCellRefs(cellBelow), ",")
    Dim lenAbove As Integer: lenAbove = ArrayLen(aboveRefs)
    If (lenAbove <> ArrayLen(belowRefs)) Then
        isRepeatFormula = False
        Exit Function
    End If
    Dim transformedBelowValue As String: transformedBelowValue = cellBelow.value
    'If we get to here then the lengths of the arrays match
    Dim i As Integer
    For i = 0 To lenAbove - 1
        transformedBelowValue = Replace(transformedBelowValue, belowRefs(i), aboveRefs(i))
    Next
    isRepeatFormula = cellAbove.value = transformedBelowValue
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
        ExtractCellRefs = "No Matches"
    End If
End Function

Private Sub BordersAroundUsedRange(ws As Worksheet)
    Dim outputRange As Range
    Set outputRange = ws.usedRange
    outputRange.Borders.LineStyle = xlContinuous
    outputRange.Columns.AutoFit
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
