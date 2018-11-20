Attribute VB_Name = "EventHandler"
Option Explicit
Public Const EH_UNIQUE_STRING = "j9gDJlWQZSUFmzOFWK"
'Do not move the above line 2 to any other line - it's there to uniquely identify this module

Public Sub EH_BeforeClose()
    Call deleteShortcuts
    If MsgBox("Export Before closing?", vbYesNo + vbQuestion) = vbNo Then Exit Sub
    Call ExportModules
End Sub

Public Sub EH_BeforeSave()
    Call ExportModules
End Sub

Public Sub EH_Open()
    Call setShortcuts
End Sub

Private Sub setShortcuts()
    Application.OnKey "+^{X}", "ExportModules"
    Application.OnKey "+^{I}", "ImportModulesWarn"
End Sub

Private Sub deleteShortcuts()
    Application.OnKey "+^{X}"
    Application.OnKey "+^{I}"
End Sub
