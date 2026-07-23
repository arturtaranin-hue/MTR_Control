Attribute VB_Name = "modLogger"
Option Explicit

Public Enum LogLevel
    LogLevelInfo = 1
    LogLevelWarning = 2
    LogLevelError = 3
End Enum

Public Sub LogInfo(ByVal message As String, Optional ByVal context As String = vbNullString)
    WriteLog LogLevelInfo, message, context
End Sub

Public Sub LogWarning(ByVal message As String, Optional ByVal context As String = vbNullString)
    WriteLog LogLevelWarning, message, context
End Sub

Public Sub LogError(ByVal errorNumber As Long, ByVal message As String, Optional ByVal context As String = vbNullString)
    WriteLog LogLevelError, "Error " & CStr(errorNumber) & ": " & message, context
End Sub

Private Sub WriteLog(ByVal level As LogLevel, ByVal message As String, ByVal context As String)
    Debug.Print Format$(Now, "yyyy-mm-dd hh:nn:ss"), LogLevelName(level), context, message
End Sub

Private Function LogLevelName(ByVal level As LogLevel) As String
    Select Case level
        Case LogLevelInfo
            LogLevelName = "INFO"
        Case LogLevelWarning
            LogLevelName = "WARN"
        Case LogLevelError
            LogLevelName = "ERROR"
        Case Else
            LogLevelName = "UNKNOWN"
    End Select
End Function
