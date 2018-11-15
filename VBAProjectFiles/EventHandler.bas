Attribute VB_Name = "EventHandler"
Option Explicit

'Procedures prefixed with this string will be called from the corresponding
'event procedure in the ThisWorkbook module
Const EH_PREFIX = "EH_"
Private Const WB_PREFIX = "Workbook_"

Sub testShouldLoadToWb()
    Dim testVal As String: testVal = "Foo"
    MsgBox (testVal & ": " & shouldLoadToWb(testVal))
    testVal = EH_PREFIX & "Bar"
    MsgBox (testVal & ": " & shouldLoadToWb(testVal))
End Sub

'Return True iff the procedure by that name should to to the ThisWorkbook module
Public Function shouldLoadToWb(ByVal procedureName As String) As Boolean
    If InStr(1, procedureName, EH_PREFIX) = 1 Then
        shouldLoadToWb = True
    Else
        shouldLoadToWb = False
    End If
End Function

'Each function in the EventHandler has a target function in the workbook to call it
Public Function wbTargetFunction(ByVal ehFnName As String) As String
    wbTargetFunction = Replace(ehFnName, EH_PREFIX, WB_PREFIX)
End Function

Private Sub EH_BeforeSave()
    MsgBox ("Stub Before Save: Please implement me!")
End Sub
