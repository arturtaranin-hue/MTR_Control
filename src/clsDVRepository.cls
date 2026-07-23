VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDVRepository"
Option Explicit

Public Function Load(ByVal path As String) As Collection
    Dim helper As clsWorkbookHelper
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim data As Variant
    Dim headers As Object
    Dim result As Collection
    Dim row As Long
    Dim item As clsMaterial
    Dim errorNumber As Long
    Dim errorSource As String
    Dim errorDescription As String

    On Error GoTo CleanFail
    LogImportStep "DVRepository", "Starting DV import"
    Set helper = New clsWorkbookHelper
    Set wb = helper.OpenWorkbook(path)
    Set ws = helper.FindWorksheet(wb, Array("Code", "Name", "Quantity"), Array("DV", "Materials", "Material"))
    data = helper.ReadRange(ws)
    Set headers = helper.BuildHeaderMap(data)
    Set result = New Collection

    If UBound(data, 1) <= LBound(data, 1) Then Err.Raise vbObjectError + 5200, "DVRepository.Load", "Empty file"
    For row = LBound(data, 1) + 1 To UBound(data, 1)
        If Len(Trim$(CStr(data(row, headers("Code"))))) > 0 Then
            Set item = New clsMaterial
            item.Code = Trim$(CStr(data(row, headers("Code"))))
            item.Name = Trim$(CStr(data(row, headers("Name"))))
            If Not IsNumeric(data(row, headers("Quantity"))) Then Err.Raise vbObjectError + 5201, "DVRepository.Load", "Wrong format: Quantity must be numeric at row " & row
            item.Quantity = CDbl(data(row, headers("Quantity")))
            item.SourceRow = row
            result.Add item
        End If
    Next row
    If result.Count = 0 Then Err.Raise vbObjectError + 5202, "DVRepository.Load", "Empty file"

    LogImportStep "DVRepository", "Imported materials: " & result.Count
    Set Load = result
CleanExit:
    helper.CloseWorkbook wb
    LogImportStep "DVRepository", "Finished DV import"
    Exit Function
CleanFail:
    errorNumber = Err.Number
    errorSource = Err.Source
    errorDescription = Err.Description
    LogImportStep "DVRepository", "Failed DV import: " & errorDescription
    helper.CloseWorkbook wb
    Err.Raise errorNumber, errorSource, errorDescription
End Function
