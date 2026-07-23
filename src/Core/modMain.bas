Attribute VB_Name = "modMain"
Option Explicit

Public Sub Main()
    On Error GoTo ErrorHandler

    LogInfo "MTR_Control startup", "modMain.Main"

    Exit Sub

ErrorHandler:
    LogError Err.Number, Err.Description, "modMain.Main"
End Sub
