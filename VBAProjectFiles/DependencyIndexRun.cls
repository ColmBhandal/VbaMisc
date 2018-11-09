VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DependencyIndexRun"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Object for storing data about a single run of dependency indexing

Public m_chunkSize As Long
Public m_runSize As Long
Private m_chunkFormulasCompleted As Long
Private m_runFormulasCompleted As Long
Private m_currentRow As Long
Private m_currentSheet As Worksheet
Private m_externalOnly As Boolean

'If the row goes beyond the last row in the sheet, switches sheet
Public Sub processRow()
    Dim formulasInRow As Long: formulasInRow = 0
    Dim sheetRows As Range: Set sheetRows = m_currentSheet.usedRange.Rows
    If (m_currentRow > sheetRows.count) Then
        Call switchSheet
        Exit Sub
    End If
    Dim row As Range: Set row = sheetRows(m_currentRow)
    Dim cell As Range
    For Each cell In row.Cells
        'Only the first cell of a merge cell should be processed. The rest should be skipped.
        If isFirstMergedCell(cell) Then
            If cell.HasFormula Then
                Dim precs As Collection
                If m_externalOnly Then
                    Set precs = findExternalPrecedents(cell)
                Else
                    Set precs = findAllPrecedents(cell)
                End If
                Dim prec As Range
                'Debug.Print "Source cell: " & cell.Address(external:=True)
                For Each prec In precs
                    Dim precCell As Range
                    'Debug.Print "...Precedent range: " & prec.Address(external:=True)
                    For Each precCell In getShadowRange(prec).Cells
                        'DoEvents
                        'Debug.Print "...--->Precedent single cell: " & precCell.Address(external:=True)
                        If IsEmpty(precCell) Then
                           precCell.value = sheetCellAddress(cell)
                        Else
                           precCell.value = precCell.value & ", " & sheetCellAddress(cell)
                        End If
                    Next precCell
                Next prec
                formulasInRow = formulasInRow + 1
            End If
        End If
    Next cell
    m_currentRow = m_currentRow + 1
    incrementFormulas (formulasInRow)
End Sub

'Get the corresponding range to prec on the shadow sheet
'Restricts shadow range to only used cells in the precedent sheet
'Assumes prec is a range on a regular sheet
Function getShadowRange(prec As Range) As Range
    Dim ws As Worksheet: Set ws = prec.Worksheet
    Dim boundedRange As Range: Set boundedRange = Intersect(prec, ws.usedRange)
    Dim shadowSheet As Worksheet: Set shadowSheet = getBuddySheet(ws)
    shadowSheet.Activate
    Set getShadowRange = Range(boundedRange.Address)
End Function

'A first merged cell is either not merged or it is the first in a range of merged cells
Private Function isFirstMergedCell(cell As Range) As Boolean
    If Not cell.MergeCells Then
        isFirstMergedCell = True
        Exit Function
    End If
    isFirstMergedCell = cell.MergeArea.Cells(1).Address = cell.Address
End Function

Private Sub switchSheet()
    Dim nextSheetCell As Range: Set nextSheetCell = getNextSheetCell()
    nextSheetCell.value = m_currentSheet.name
    Dim currSheetCell As Range: Set currSheetCell = getCellUnderHeader(CURR_SHEET)
    If isLastSheet(m_currentSheet) Then
        Dim name As String: name = m_currentSheet.name
        m_currentSheet = Nothing
        currSheetCell.Clear
        Debug.Print "Finished the last sheet: " & name
        Exit Sub
    End If
    Set m_currentSheet = Sheets(m_currentSheet.index + 1)
    currSheetCell.value = m_currentSheet.name
    m_currentRow = 1
    Dim currRowCell As Range: Set currRowCell = getCellUnderHeader(CURR_ROW)
    currRowCell.value = m_currentRow
End Sub

Private Sub Class_Initialize()
    m_chunkFormulasCompleted = 0
    m_runFormulasCompleted = 0
    m_currentRow = 1
    m_externalOnly = True
End Sub

Public Property Let ExternalOnly(extOnly As Boolean)
    m_externalOnly = extOnly
End Property

Public Property Let CurrentRow(rowNum As Long)
    m_currentRow = rowNum
End Property

Public Property Let CurrentSheet(sheet As Worksheet)
    Set m_currentSheet = sheet
End Property

Public Property Get ContinueChunk() As Boolean
    ContinueChunk = m_chunkFormulasCompleted < m_chunkSize And (Not m_currentSheet Is Nothing)
End Property

Public Property Get ContinueRun() As Boolean
    ContinueRun = m_runFormulasCompleted < m_runSize And (Not m_currentSheet Is Nothing)
End Property

Public Property Get CurrentRow() As Long
    CurrentRow = m_currentRow
End Property

Public Property Get CurrentSheet() As Worksheet
    Set CurrentSheet = m_currentSheet
End Property

Public Sub incrementFormulas(numFormulas As Long)
    m_chunkFormulasCompleted = m_chunkFormulasCompleted + numFormulas
    m_runFormulasCompleted = m_runFormulasCompleted + numFormulas
End Sub

Public Property Get ChunkFormulaCount() As Long
    ChunkFormulaCount = m_chunkFormulasCompleted
End Property

Public Sub resetChunk()
    m_chunkFormulasCompleted = 0
End Sub
