Attribute VB_Name = "modImportLogger"
Option Explicit

Public Sub LogImportStep(ByVal sourceName As String, ByVal message As String)
    Debug.Print Format$(Now, "yyyy-mm-dd hh:nn:ss") & " [Import] [" & sourceName & "] " & message
End Sub
