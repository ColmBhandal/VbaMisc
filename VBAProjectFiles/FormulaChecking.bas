Attribute VB_Name = "FormulaChecking"
Option Explicit

Const COMPRESSED_COL_SHEET = "CompressedColumns"
Const PRECEDENTS_OUTPUT = "PrecedentsOutput"
Const BOTH_COL = &HFF0000
Const DEP_COL = &HFF&
Const PREC_COL = &HFF00&

Sub CompressColumnsNoColour()
Attribute CompressColumnsNoColour.VB_ProcData.VB_Invoke_Func = " \n14"
    CompressColumns (False)
End Sub

Sub CompressColumnsWithColour()
Attribute CompressColumnsWithColour.VB_ProcData.VB_Invoke_Func = " \n14"
    CompressColumns (True)
End Sub

'Write the precedents for each cell in selected range to an output sheet & logs time taken to do so
Sub WritePrecedents()
Attribute WritePrecedents.VB_ProcData.VB_Invoke_Func = "w\n14"
    Dim startTime As Double
    startTime = Now
    Call ResetOutput(PRECEDENTS_OUTPUT)
    Dim outputSheet As Worksheet
    Set outputSheet = ActiveWorkbook.Sheets(PRECEDENTS_OUTPUT)
    Dim selectedRange As Range
    Dim currentCell As Range
    Set selectedRange = Selection
    Call mUnhideAll
    For Each currentCell In selectedRange
        If currentCell.HasFormula Then
            Dim precedents As Collection
            Set precedents = findAllPrecedents(currentCell)
            Dim dent As Variant
            Dim row As Integer, col As Integer
            row = currentCell.row
            col = currentCell.Column
            For Each dent In precedents
                Dim existing As String
                existing = outputSheet.Cells(row, col).value
                Dim dentRng As Range
                Set dentRng = dent
                outputSheet.Cells(row, col).value = existing & dentRng.Address(external:=True) & vbCrLf
            Next dent
        End If
    Next currentCell
    Dim outputRange As Range
    Set outputRange = outputSheet.UsedRange
    outputRange.Borders.LineStyle = xlContinuous
    outputRange.Columns.AutoFit
    outputSheet.Activate
    Debug.Print ("Total time: " & Minute(Now - startTime) & ":" & Second(Now - startTime))
End Sub

'Compresses the selected range by writing only its formula columns to another sheet and skipping non-formula columns
'colourDents determines whether we want to colour formulas with dependencies
Sub CompressColumns(colourDents As Boolean)
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
                    If colourDents Then
                        outputCell.Interior.Color = getExternalDependencyColour(cell)
                        outputCell.Font.Color = getInternalDependencyColour(cell)
                    End If
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

Sub GetDentColour()
Attribute GetDentColour.VB_ProcData.VB_Invoke_Func = "o\n14"
    MsgBox getExternalDependencyColour(Selection)
End Sub

Private Function getExternalDependencyColour(cell As Range)
Dim precExternal As Boolean
Dim depExternal As Boolean
precExternal = hasPartialDents(cell, False, True)
depExternal = hasPartialDents(cell, False, False)
Select Case True
    Case precExternal And depExternal
        getExternalDependencyColour = BOTH_COL
    Case precExternal And Not depExternal
        getExternalDependencyColour = PREC_COL
    Case Not precExternal And depExternal
        getExternalDependencyColour = DEP_COL
    Case Else
        getExternalDependencyColour = cell.Interior.Color
End Select
End Function

Private Function getInternalDependencyColour(cell As Range)
Dim precInternal As Boolean
Dim depInternal As Boolean
precInternal = hasPartialDents(cell, True, True)
depInternal = hasPartialDents(cell, True, False)
Select Case True
    Case precInternal And depInternal
        getInternalDependencyColour = BOTH_COL
    Case precInternal And Not depInternal
        getInternalDependencyColour = PREC_COL
    Case Not precInternal And depInternal
        getInternalDependencyColour = DEP_COL
    Case Else
        getInternalDependencyColour = cell.Font.Color
End Select
End Function

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
