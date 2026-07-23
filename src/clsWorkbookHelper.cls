VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsWorkbookHelper"
Option Explicit

Public Function OpenWorkbook(ByVal path As String) As Workbook
    LogImportStep "WorkbookHelper", "Opening workbook read-only: " & path
    If Len(Dir$(path)) = 0 Then Err.Raise vbObjectError + 5100, "WorkbookHelper.OpenWorkbook", "File not found: " & path
    If IsWorkbookAlreadyOpen(path) Then Err.Raise vbObjectError + 5101, "WorkbookHelper.OpenWorkbook", "Workbook already opened: " & path
    Set OpenWorkbook = Application.Workbooks.Open(Filename:=path, ReadOnly:=True, UpdateLinks:=False, AddToMru:=False)
End Function

Public Sub CloseWorkbook(ByRef workbookToClose As Workbook)
    On Error Resume Next
    If Not workbookToClose Is Nothing Then
        LogImportStep "WorkbookHelper", "Closing workbook: " & workbookToClose.Name
        workbookToClose.Close SaveChanges:=False
    End If
    Set workbookToClose = Nothing
    On Error GoTo 0
End Sub

Public Function FindWorksheet(ByVal workbookToSearch As Workbook, ByVal expectedHeaders As Variant, Optional ByVal preferredNames As Variant) As Worksheet
    Dim sheet As Worksheet
    Dim headerMap As Object

    LogImportStep "WorkbookHelper", "Detecting worksheet"
    If workbookToSearch.Worksheets.Count = 0 Then Err.Raise vbObjectError + 5102, "WorkbookHelper.FindWorksheet", "Missing worksheet"

    If Not IsMissing(preferredNames) Then
        For Each sheet In workbookToSearch.Worksheets
            If NameMatches(sheet.Name, preferredNames) Then
                Set headerMap = BuildHeaderMap(sheet.UsedRange.Value2)
                If ContainsHeaders(headerMap, expectedHeaders) Then
                    Set FindWorksheet = sheet
                    Exit Function
                End If
            End If
        Next sheet
    End If

    For Each sheet In workbookToSearch.Worksheets
        Set headerMap = BuildHeaderMap(sheet.UsedRange.Value2)
        If ContainsHeaders(headerMap, expectedHeaders) Then
            Set FindWorksheet = sheet
            Exit Function
        End If
    Next sheet

    Err.Raise vbObjectError + 5103, "WorkbookHelper.FindWorksheet", "Missing worksheet with required headers"
End Function

Public Function ReadRange(ByVal worksheetToRead As Worksheet) As Variant
    LogImportStep "WorkbookHelper", "Reading UsedRange from worksheet: " & worksheetToRead.Name
    If WorksheetFunction.CountA(worksheetToRead.UsedRange) = 0 Then Err.Raise vbObjectError + 5104, "WorkbookHelper.ReadRange", "Empty file"
    ReadRange = worksheetToRead.UsedRange.Value2
End Function

Private Function IsWorkbookAlreadyOpen(ByVal path As String) As Boolean
    Dim candidate As Workbook
    For Each candidate In Application.Workbooks
        If StrComp(candidate.FullName, path, vbTextCompare) = 0 Then
            IsWorkbookAlreadyOpen = True
            Exit Function
        End If
    Next candidate
End Function

Private Function NameMatches(ByVal sheetName As String, ByVal names As Variant) As Boolean
    Dim i As Long
    For i = LBound(names) To UBound(names)
        If StrComp(sheetName, CStr(names(i)), vbTextCompare) = 0 Then
            NameMatches = True
            Exit Function
        End If
    Next i
End Function

Public Function BuildHeaderMap(ByVal data As Variant) As Object
    Dim result As Object
    Dim col As Long
    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = vbTextCompare

    If Not IsArray(data) Then
        Set BuildHeaderMap = result
        Exit Function
    End If

    For col = LBound(data, 2) To UBound(data, 2)
        If Len(Trim$(CStr(data(LBound(data, 1), col)))) > 0 Then result(Trim$(CStr(data(LBound(data, 1), col)))) = col
    Next col
    Set BuildHeaderMap = result
End Function

Public Function ContainsHeaders(ByVal headerMap As Object, ByVal expectedHeaders As Variant) As Boolean
    Dim i As Long
    For i = LBound(expectedHeaders) To UBound(expectedHeaders)
        If Not headerMap.Exists(CStr(expectedHeaders(i))) Then Exit Function
    Next i
    ContainsHeaders = True
End Function
