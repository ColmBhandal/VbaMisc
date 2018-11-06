Attribute VB_Name = "FormulaChecking"
Option Explicit

Const COMPRESSED_COL_SHEET = "CompressedColumns"

'Compresses the selected range by writing only its formula columns to another sheet and skipping non-formula columns
Sub CompressColumns()
Attribute CompressColumns.VB_ProcData.VB_Invoke_Func = "k\n14"
    Dim startTime As Double
    startTime = Now
    Call ResetOutput(COMPRESSED_COL_SHEET)
    Dim outputSheet As Worksheet
    Set outputSheet = ActiveWorkbook.Sheets(COMPRESSED_COL_SHEET)
    Dim selectedRange As Range
    Dim col As Range
    Set selectedRange = Selection
    Dim currOutputCol As Integer
    currOutputCol = 1
    For Each col In selectedRange.Columns
        If HasSomeFormulas(col) Then
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
    Dim outputRange As Range
    Set outputRange = outputSheet.UsedRange
    outputRange.Borders.LineStyle = xlContinuous
    outputRange.Columns.AutoFit
    outputSheet.Activate
    Debug.Print ("Total time: " & Minute(Now - startTime) & ":" & Second(Now - startTime))
End Sub

'True if the rng has at least one formula, False otherwise
Private Function HasSomeFormulas(rng As Range) As Boolean
    If rng.HasFormula = False Then
        HasSomeFormulas = False
    Else
        HasSomeFormulas = True
    End If
End Function

Private Function columnLetter(lngCol As Long) As String
    Dim vArr
    vArr = Split(Cells(1, lngCol).Address(True, False), "$")
    columnLetter = vArr(0)
End Function
