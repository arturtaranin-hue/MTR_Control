VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsWorkbookRepository"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mWorkbook As Workbook

Public Sub Initialize(ByVal targetWorkbook As Workbook)
    Set mWorkbook = targetWorkbook
End Sub

Public Function ReadRangeToArray(ByVal worksheetName As String, ByVal address As String) As Variant
    On Error GoTo ErrorHandler

    EnsureInitialized
    ReadRangeToArray = mWorkbook.Worksheets(worksheetName).Range(address).Value2

    Exit Function

ErrorHandler:
    LogError Err.Number, Err.Description, "clsWorkbookRepository.ReadRangeToArray"
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Private Sub EnsureInitialized()
    If mWorkbook Is Nothing Then
        Err.Raise vbObjectError + 1000, "clsWorkbookRepository", "Repository has not been initialized."
    End If
End Sub
