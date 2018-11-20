Attribute VB_Name = "EventHandler"
Option Explicit

Public Sub EH_BeforeClose()
    If MsgBox("Export Before closing?", vbYesNo + vbQuestion) = vbNo Then Exit Sub
    Call ExportModules
End Sub

Public Sub EH_BeforeSave()
    Call ExportModules
End Sub
