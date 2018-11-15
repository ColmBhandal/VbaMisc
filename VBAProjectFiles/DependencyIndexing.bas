Attribute VB_Name = "DependencyIndexing"
Option Explicit

'''''''''''Constants
Const DEP_PREFIX As String = "DEPS_"
Const GREEN As Long = &HCCFFCC
Const STATUS_NONE As String = "None"
Const BACKUP_SUFFIX As String = "_Backup"
Const DEFAULT_CHUNK_SIZE As Integer = 1 '00
'We expect a rate of about 10 formulas per second so 36000 formulas per hour, and an hour seems like a good default
Const DEFAULT_RUN_SIZE As Long = 36 '000
Const DEFAULT_EXTERNAL_ONLY As Boolean = True

'''''''''''Header constants
Const META_SHEET = DEP_PREFIX & "META"
Const SHEETS_DONE = "Sheets Done"
Public Const CURR_SHEET = "Current Sheet"
Public Const CURR_ROW = "Current Row"
Const TOTAL_WB_FORMULAS = "Total Workbook Formulas"
Const COMPLETED_FORMULAS = "Formulas Completed"
Const PERCENT_COMPLETE = "Percent Completed"
Const STATUS = "Status"
Const CHUNK_SIZE = "Chunk Size"
Const RUN_SIZE = "Run Size"
Const EXTERNAL_ONLY = "External Only"

Public Sub test()
Attribute test.VB_ProcData.VB_Invoke_Func = "q\n14"
    Call doRun
End Sub

Sub doRun()
    Dim startTime As Double
    startTime = Now
    Application.ScreenUpdating = False
    Application.EnableCancelKey = xlInterrupt
    Debug.Print ("************** Running Initial chores **************")
    Dim runObject As DependencyIndexRun: Set runObject = New DependencyIndexRun
    Call initialChores(runObject)
    Debug.Print ("************** Initial chores Complete **************")
    Call doRunMain(runObject)
    Debug.Print ("************** Run Done **************")
    Application.ScreenUpdating = True
    MsgBox ("Run Complete! Total time: " & Minute(Now - startTime) & ":" & Second(Now - startTime))
End Sub

Sub doRunMain(runObject As DependencyIndexRun)
    Do While runObject.ContinueRun
        Call doChunk(runObject)
        DoEvents
        Dim chunkDoneMsg As String: chunkDoneMsg = "************** Chunk Done up to " & _
            runObject.CurrentSheet.name & ": " & runObject.CurrentRow & " **************"
        reportEndOfChunk (chunkDoneMsg)
        Call activateMetaSheet
        Call updateMetaSheet(runObject.ChunkFormulaCount, runObject.CurrentRow)
        runObject.resetChunk
        ActiveSheet.usedRange.Columns.AutoFit
    Loop
End Sub

Sub doChunk(runObject As DependencyIndexRun)
    Do While runObject.ContinueChunk
        runObject.processRow
    Loop
End Sub

Sub reportEndOfChunk(endOfChunkMsg As String)
    Debug.Print Now & endOfChunkMsg
End Sub

Sub initialChores(runObject As DependencyIndexRun)
    'Order and backup
    If verifySheetsOrder() Then
        Debug.Print ("Sheets seem to be in order")
    Else
        Dim errorDesc As String: errorDesc = "Sheets order reported by meta sheet does not match that of actual workbook. Quitting."
        Err.Raise Number:=513, Description:=errorDesc
    End If
    Call backup
    Call readInitialValues(runObject)
    Dim currSheet As Worksheet: Set currSheet = getCurrentSheet()
    If currSheet Is Nothing Then
        Dim firstSheet As Worksheet: Set firstSheet = ThisWorkbook.Sheets(1)
        getCellUnderHeader(CURR_SHEET).Value = firstSheet.name
        getCellUnderHeader(CURR_ROW).Value = 1
        Set currSheet = firstSheet
    End If
    runObject.CurrentSheet = currSheet
End Sub

'Read in values from sheet
Sub readInitialValues(runObject As DependencyIndexRun)
    Dim chunkCell As Range: Set chunkCell = getCellUnderHeader(CHUNK_SIZE)
    Dim runCell As Range: Set runCell = getCellUnderHeader(RUN_SIZE)
    If IsEmpty(chunkCell) Or IsEmpty(runCell) Then
        Dim errorDesc As String: errorDesc = "Chunk Size and Run Size must be specified."
        Err.Raise Number:=513, Description:=errorDesc
    End If
    runObject.m_chunkSize = chunkCell.Value
    runObject.m_runSize = runCell.Value
    'Only read non-blank valus. Else leave them at object default.
    Dim rowCell As Range: Set rowCell = getCellUnderHeader(CURR_ROW)
    If Not IsEmpty(rowCell) Then _
        runObject.CurrentRow = rowCell.Value
    Dim extOnlyCell As Range: Set extOnlyCell = getCellUnderHeader(EXTERNAL_ONLY)
    If Not IsEmpty(extOnlyCell) Then _
        runObject.ExternalOnly = extOnlyCell.Value
End Sub

Sub testIsLastSheet()
    MsgBox (isLastSheet(Sheets("US_States")))
    MsgBox (isLastSheet(Sheets("Canadian_Provinces")))
End Sub

'Is this the last sheet in the workbook, excluding specially prefixed sheets
'Assumes shadow/meta sheets have been set up correctly -
'So there should be 2n+1 sheets if there are n original sheets- n shadows + 1 meta
Public Function isLastSheet(ws As Worksheet) As Boolean
    Dim numRegularSheets As Integer: numRegularSheets = regularSheetsCount
    isLastSheet = (ws.index = numRegularSheets)
End Function

Function getNextSheetCell() As Range
    Dim sheetsRng As Range
    Set sheetsRng = getContiguousConstsDown(getCellUnderHeader(SHEETS_DONE))
    Dim nextEmptyCell As Range
    If IsEmpty(sheetsRng) Then
        Set getNextSheetCell = sheetsRng
    Else
        Set getNextSheetCell = sheetsRng.Offset(sheetsRng.count).Cells(1)
    End If
End Function

Sub updateMetaSheet(ChunkFormulaCount As Long, rowNum As Long)
    Dim formComplCell As Range: Set formComplCell = getCellUnderHeader(COMPLETED_FORMULAS)
    Dim totalFormsComl As Double: totalFormsComl = formComplCell.Value + ChunkFormulaCount
    formComplCell.Value = totalFormsComl
    Dim totWbFormsCell As Range: Set totWbFormsCell = getCellUnderHeader(TOTAL_WB_FORMULAS)
    Dim totalWbForms As Double: totalWbForms = totWbFormsCell.Value
    Dim pcntComplete As Double: pcntComplete = (totalFormsComl / totWbFormsCell)
    Dim pcntComplCell As Range: Set pcntComplCell = getCellUnderHeader(PERCENT_COMPLETE)
    pcntComplCell.Value = pcntComplete
    pcntComplCell.NumberFormat = "0.00%"
    Dim currRowCell As Range: Set currRowCell = getCellUnderHeader(CURR_ROW)
    currRowCell.Value = rowNum
End Sub


'Saves a backup of this workbook if one doesn't already exist
Sub backup()
    Dim fullName As String: fullName = ThisWorkbook.fullName
    Const extension As String = ".xlsb"
    Dim backupFullName As String: backupFullName = Split(fullName, ".xlsb")(0)
    backupFullName = backupFullName & BACKUP_SUFFIX & extension
    If (Len(Dir(backupFullName)) <> 0) Then
        Debug.Print ("Backup file already exists: " & backupFullName)
    Else
        Debug.Print ("No backup file exists. Saving backup as: " & backupFullName)
        Application.EnableCancelKey = xlDisabled
        Application.DisplayAlerts = False
        ThisWorkbook.SaveAs (backupFullName)
        ThisWorkbook.SaveAs (fullName)
        Application.DisplayAlerts = True
        Application.EnableCancelKey = xlInterrupt
    End If
End Sub

'True iff no sheets in the workbook have been skipped so far out of the sheets done
'Also checks that the current sheet we're working on is next in line, if it's not blank
Function verifySheetsOrder() As Boolean
    Dim sheetsSoFar As Collection: Set sheetsSoFar = getSheetsDone()
    Dim CurrentSheet As Worksheet: Set CurrentSheet = getCurrentSheet()
    If Not CurrentSheet Is Nothing Then
        sheetsSoFar.Add Item:=CurrentSheet
    End If
    Dim sheet As Worksheet
    Dim index As Integer: index = 1
    For Each sheet In sheetsSoFar
        If sheet.name <> Sheets(index).name Then
            verifySheetsOrder = False
            Exit Function
        End If
        index = index + 1
    Next sheet
    verifySheetsOrder = True
    Exit Function

End Function

Function getCurrentSheet() As Worksheet
    Dim currentSheetRng As Range: Set currentSheetRng = getCellUnderHeader(CURR_SHEET)
    Dim currentSheetRngName As String
    If Not IsEmpty(currentSheetRng) Then
        currentSheetRngName = currentSheetRng.Value
        On Error GoTo current_sheet_error
        Set getCurrentSheet = Sheets(currentSheetRngName)
        Exit Function
        On Error GoTo 0
    End If
    Set getCurrentSheet = Nothing
    Exit Function
current_sheet_error:
    MsgBox ("Problem getting current sheet: " & currentSheetRngName)
End Function

Sub testGetSheetsDone()
    Dim msg As String
    Dim sheetsDone As Collection: Set sheetsDone = getSheetsDone()
    Dim sheet As Variant
    For Each sheet In sheetsDone
      Dim ws As Worksheet: Set ws = sheet
        msg = msg & ws.name & vbCrLf
    Next sheet
    MsgBox (msg)
End Sub

Function getSheetsDone() As Collection
    Call activateMetaSheet
    Dim sheetsRng As Range
    Set sheetsRng = getContiguousConstsDown(getCellUnderHeader(SHEETS_DONE))
    Dim result As New Collection
    If IsEmpty(sheetsRng) Then
        Set getSheetsDone = result
        Exit Function
    End If
    Dim cell As Range
    Dim sheetName As String
    For Each cell In sheetsRng.Cells
        sheetName = cell.Value
        If (InStr(1, sheetName, DEP_PREFIX)) Then _
            GoTo deps_sheet_error
        On Error GoTo general_error
            Dim ws As Worksheet: Set ws = Sheets(sheetName)
            result.Add Item:=ws
        On Error GoTo 0
    Next cell
    Set getSheetsDone = result
    Exit Function
deps_sheet_error:
    MsgBox ("Error: Detected the sheet " & sheetName & " begins with " & DEP_PREFIX)
    Exit Function
general_error:
    MsgBox ("Error processing sheet: " & sheetName)
End Function

'Deletes all deps sheets then resets them all- use this to start afresh completeley
Sub cleanResetDepsSheets()
Attribute cleanResetDepsSheets.VB_ProcData.VB_Invoke_Func = "s\n14"
    Call removeAllDepsSheets
    Debug.Print "--------- All Deps Sheets Removed ---------"
    Call resetAllDepsSheets
    Debug.Print "--------- All Deps Sheets Reset ---------"
End Sub

Sub resetAllDepsSheets()
    Call resetShadowSheets
    Debug.Print "**** All Shadow Sheets Reset ****"
    Call ResetMetaSheet
    Debug.Print "**** Meta Sheet Reset ****"
End Sub

'For each sheet not prefixed with our prefix, creates another sheet shadowing it
Sub resetShadowSheets()
    Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Sheets
        Dim name As String: name = ws.name
        Dim prefixLength As Integer: prefixLength = Len(DEP_PREFIX)
        Dim prefix As String: prefix = Left(name, prefixLength)
        If prefix <> DEP_PREFIX Then
            'We need to chop to 31 characters as that's Excel's limit
            Dim newName As String: newName = Left(DEP_PREFIX & name, 31)
            Application.StatusBar = "Resetting: " & newName
            ResetOutput (newName)
            Debug.Print "Reset Shadow sheet: " & newName
        End If
    Next ws
End Sub

'Removes the shadow & meta sheets from the workbook
Sub removeAllDepsSheets()
    Dim ws As Worksheet
    Dim toDelete As New Collection
    For Each ws In ActiveWorkbook.Sheets
        Dim name As String: name = ws.name
        Dim prefixLength As Integer: prefixLength = Len(DEP_PREFIX)
        Dim prefix As String: prefix = Left(name, prefixLength)
        If prefix = DEP_PREFIX Then
            toDelete.Add Item:=ws
        End If
    Next ws
    Application.DisplayAlerts = False
    For Each ws In toDelete
        name = ws.name
        ws.Delete
        Debug.Print ("Deleted sheet: " & name)
    Next ws
    Application.DisplayAlerts = True
End Sub

'Toggles between shadow sheet and regular sheet
Sub toggleBuddy()
Attribute toggleBuddy.VB_ProcData.VB_Invoke_Func = "w\n14"
    If isMetaSheet(ActiveSheet) Then
        Call MsgBox("Meta sheet has no buddies.", vbExclamation)
        Exit Sub
    End If
    getBuddySheet(ActiveSheet).Activate
End Sub

'A shadow sheet and its original sheet are each other's buddies. Assumes ws is not meta- that has no buddies.
Function getBuddySheet(ws As Worksheet) As Worksheet
    Dim newIndex As Integer
    If isDepsSheet(ws) Then
        newIndex = ws.index - regularSheetsCount()
    Else
        newIndex = ws.index + regularSheetsCount()
    End If
    On Error GoTo noSheetFound
    Set getBuddySheet = Sheets(newIndex)
    On Error GoTo 0
    Exit Function
noSheetFound:
    MsgBox ("No corresponding sheet found for sheet: " & ws.name)
End Function

Function isDepsSheet(ws As Worksheet) As Boolean
    isDepsSheet = InStr(1, ws.name, DEP_PREFIX) = 1
End Function

Function isMetaSheet(ws As Worksheet) As Boolean
    isMetaSheet = ws.name = META_SHEET
End Function

Function regularSheetsCount()
    Dim totalNumberOfSheets As Integer: totalNumberOfSheets = ThisWorkbook.Sheets.count
    Dim sheetsPlusShadows As Integer: sheetsPlusShadows = totalNumberOfSheets - 1
    If sheetsPlusShadows Mod 2 = 1 Then
        Dim errorDesc As String: errorDesc = "Uneven number of sheets detected. Have you generated all shadow sheets?"
        Err.Raise Number:=513, Description:=errorDesc
    End If
    regularSheetsCount = sheetsPlusShadows / 2
End Function

Sub ResetMetaSheet()
Attribute ResetMetaSheet.VB_ProcData.VB_Invoke_Func = "s\n14"
    ResetOutput (META_SHEET)
    Dim currentCell As Range: Set currentCell = Range("A1")
    Call writeStepRight(currentCell, SHEETS_DONE)
    Call writeStepRight(currentCell, CURR_SHEET)
    Call writeStepRight(currentCell, CURR_ROW)
    Call writeStepRight(currentCell, COMPLETED_FORMULAS)
    Call writeStepRight(currentCell, TOTAL_WB_FORMULAS)
    Call writeStepRight(currentCell, PERCENT_COMPLETE)
    Call writeStepRight(currentCell, STATUS)
    currentCell.Interior.Color = GREEN
    Call writeStepRight(currentCell, CHUNK_SIZE)
    currentCell.Interior.Color = GREEN
    Call writeStepRight(currentCell, RUN_SIZE)
    currentCell.Interior.Color = GREEN
    Call writeStepRight(currentCell, EXTERNAL_ONLY)
    Call resetMetaValues
    ActiveSheet.usedRange.Columns.AutoFit
End Sub

'Writes the value to the cell and then steps the reference right to the next cell
Private Sub writeStepRight(ByRef currentCell As Range, header As String)
    currentCell.Value = header: Set currentCell = currentCell.Offset(0, 1)
End Sub

'Gets the value in row 2 directly below said header in row 1
Public Function getCellUnderHeader(ByVal header As String) As Range
    Set getCellUnderHeader = Cells(2, getColumnByHeader(header))
End Function

'Returns -1 if no column is found
Private Function getColumnByHeader(header As String) As Integer
    Call activateMetaSheet
    Dim cell As Range
    For Each cell In ActiveSheet.usedRange.Rows(1).Cells
        If cell.Value = header Then
            getColumnByHeader = cell.column
            Exit Function
        End If
    Next cell
    getColumnByHeader = -1
End Function

Private Sub activateMetaSheet()
    ActiveWorkbook.Sheets(META_SHEET).Activate
End Sub

Sub resetFormulaCount()
    Dim count As Long: count = formulaCount()
    Call activateMetaSheet
    Cells(2, getColumnByHeader(TOTAL_WB_FORMULAS)).Value = count
End Sub

Sub resetMetaValues()
Attribute resetMetaValues.VB_ProcData.VB_Invoke_Func = " \n14"
    Call activateMetaSheet
    Call resetFormulaCount
    Cells(2, getColumnByHeader(STATUS)).Value = STATUS_NONE
    Cells(2, getColumnByHeader(CHUNK_SIZE)).Value = DEFAULT_CHUNK_SIZE
    Cells(2, getColumnByHeader(RUN_SIZE)).Value = DEFAULT_RUN_SIZE
    Cells(2, getColumnByHeader(EXTERNAL_ONLY)).Value = DEFAULT_EXTERNAL_ONLY
End Sub

Function formulaCount()
Dim ws As Worksheet
Dim rCheck As Range
Dim lCount As Long
    On Error Resume Next
        For Each ws In Worksheets
            Set rCheck = Nothing
            Set rCheck = ws.usedRange.SpecialCells(xlCellTypeFormulas)
                If Not rCheck Is Nothing Then lCount = _
                    ws.usedRange.SpecialCells(xlCellTypeFormulas).Cells.count + lCount
        Next ws
    On Error GoTo 0
    formulaCount = lCount
End Function


