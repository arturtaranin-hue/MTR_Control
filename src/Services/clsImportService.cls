VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsImportService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mRepository As clsWorkbookRepository

Public Sub Initialize(ByVal repository As clsWorkbookRepository)
    Set mRepository = repository
End Sub

Public Sub ImportMaterials(ByVal worksheetName As String, ByVal address As String)
    On Error GoTo ErrorHandler

    EnsureInitialized

    Dim data As Variant
    data = mRepository.ReadRangeToArray(worksheetName, address)

    LogInfo "Material import skeleton read source array.", "clsImportService.ImportMaterials"

    Exit Sub

ErrorHandler:
    LogError Err.Number, Err.Description, "clsImportService.ImportMaterials"
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Sub EnsureInitialized()
    If mRepository Is Nothing Then
        Err.Raise vbObjectError + 1100, "clsImportService", "Import service has not been initialized."
    End If
End Sub
