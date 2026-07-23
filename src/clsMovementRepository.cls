VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMovementRepository"
Option Explicit

Public Function Load(ByVal path As String) As Collection
    Set Load = LoadMovementWorkbook(path, "MovementRepository", Array("Movement", "Movements"))
End Function

Public Function LoadMovementWorkbook(ByVal path As String, ByVal sourceName As String, ByVal sheetNames As Variant) As Collection
    Dim helper As clsWorkbookHelper
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim data As Variant
    Dim headers As Object
    Dim result As Collection
    Dim row As Long
    Dim item As clsMovement
    Dim errorNumber As Long
    Dim errorSource As String
    Dim errorDescription As String

    On Error GoTo CleanFail
    LogImportStep sourceName, "Starting movement import"
    Set helper = New clsWorkbookHelper
    Set wb = helper.OpenWorkbook(path)
    Set ws = helper.FindWorksheet(wb, Array("MaterialCode", "MovementDate", "Quantity", "MovementType"), sheetNames)
    data = helper.ReadRange(ws)
    Set headers = helper.BuildHeaderMap(data)
    Set result = New Collection

    If UBound(data, 1) <= LBound(data, 1) Then Err.Raise vbObjectError + 5300, sourceName & ".Load", "Empty file"
    For row = LBound(data, 1) + 1 To UBound(data, 1)
        If Len(Trim$(CStr(data(row, headers("MaterialCode"))))) > 0 Then
            Set item = New clsMovement
            item.MaterialCode = Trim$(CStr(data(row, headers("MaterialCode"))))
            If Not IsDate(data(row, headers("MovementDate"))) Then Err.Raise vbObjectError + 5301, sourceName & ".Load", "Wrong format: MovementDate must be a date at row " & row
            If Not IsNumeric(data(row, headers("Quantity"))) Then Err.Raise vbObjectError + 5302, sourceName & ".Load", "Wrong format: Quantity must be numeric at row " & row
            item.MovementDate = CDate(data(row, headers("MovementDate")))
            item.Quantity = CDbl(data(row, headers("Quantity")))
            item.MovementType = Trim$(CStr(data(row, headers("MovementType"))))
            item.SourceRow = row
            result.Add item
        End If
    Next row
    If result.Count = 0 Then Err.Raise vbObjectError + 5303, sourceName & ".Load", "Empty file"

    LogImportStep sourceName, "Imported movements: " & result.Count
    Set LoadMovementWorkbook = result
CleanExit:
    helper.CloseWorkbook wb
    LogImportStep sourceName, "Finished movement import"
    Exit Function
CleanFail:
    errorNumber = Err.Number
    errorSource = Err.Source
    errorDescription = Err.Description
    LogImportStep sourceName, "Failed movement import: " & errorDescription
    helper.CloseWorkbook wb
    Err.Raise errorNumber, errorSource, errorDescription
End Function
