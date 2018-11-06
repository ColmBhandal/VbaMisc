Attribute VB_Name = "FormulaChecking"
Option Explicit

Const COMPRESSED_FORMULA_SHEET = "CompressedFormulas"

'Compresses the selected range by writing only its formula columns to another sheet and skipping non-formula columns
Sub FilterFormulaColumns()
Attribute FilterFormulaColumns.VB_ProcData.VB_Invoke_Func = "k\n14"
    Dim startTime As Double
    startTime = Now
    Call ResetOutput(COMPRESSED_FORMULA_SHEET)
    Dim outputSheet As Worksheet
    Set outputSheet = ActiveWorkbook.Sheets(COMPRESSED_FORMULA_SHEET)
    Dim selectedRange As Range
    Dim col As Range
    Set selectedRange = Selection
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
    BordersAroundUsedRange (outputSheet)
    outputSheet.Activate
    Debug.Print ("Total time: " & Minute(Now - startTime) & ":" & Second(Now - startTime))
End Sub

'A repeat row is a row which does not add any new formulas relative to the row above it
'This subroutine assumes the compressed formula sheet already exists with some formulas in it
Sub FilterRepeatRows()
    On Error GoTo problem_activating_compressed_sheet
    ActiveWorkbook.Sheets(COMPRESSED_FORMULA_SHEET).Activate
    On Error GoTo 0
    Dim lastRow As Integer: lastRow = ActiveSheet.UsedRange.rows.count
    Dim currRow As Integer
    For currRow = 2 To lastRow
        Debug.Print currRow
    Next
    Call BordersAroundUsedRange(ActiveSheet)
    Exit Sub
problem_activating_compressed_sheet:
    Call MsgBox("There was a problem activating the sheet " & COMPRESSED_FORMULA_SHEET _
    & ". Are you sure it exists?", vbExclamation)
End Sub

Private Sub BordersAroundUsedRange(ws As Worksheet)
    Dim outputRange As Range
    Set outputRange = ws.UsedRange
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

'TODO: Get original relative address from a range in the compressed sheet- go to first cell in that column/row & take numbers written there
'Test the above with a test sub that displays a message box with original cell's address for each compressed cell
