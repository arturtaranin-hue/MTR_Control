VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsExcelExporter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const REPORT_SHEET As String = "Report"
Private Const DASHBOARD_SHEET As String = "Dashboard"

Public Sub ExportFinalReport(ByVal finalReport As Collection, Optional ByVal targetWorkbook As Workbook)
    Dim workbookToUse As Workbook
    If targetWorkbook Is Nothing Then Set workbookToUse = ActiveWorkbook Else Set workbookToUse = targetWorkbook
    If workbookToUse Is Nothing Then Err.Raise vbObjectError + 5300, "ExcelExporter.ExportFinalReport", "No active workbook is available"

    LogExportStep "Starting export for " & CStr(finalReport.Count) & " report rows"
    Dim reportSheet As Worksheet
    Set reportSheet = PrepareWorksheet(workbookToUse, REPORT_SHEET)
    WriteReport reportSheet, finalReport

    Dim reportService As New clsReportService
    Dim dashboardSheet As Worksheet
    Set dashboardSheet = PrepareWorksheet(workbookToUse, DASHBOARD_SHEET)
    WriteDashboard dashboardSheet, reportService.BuildDashboard(finalReport)
    LogExportStep "Completed export to workbook: " & workbookToUse.Name
End Sub

Private Sub WriteReport(ByVal sheet As Worksheet, ByVal finalReport As Collection)
    Dim headers As Variant
    headers = Array("Material", "Planned Qty", "Arrived Qty", "Remaining Qty", "Delivery %", "Status", "Last Arrival Date", "Confidence", "Match Rule")

    Dim col As Long
    For col = LBound(headers) To UBound(headers)
        sheet.Cells(1, col + 1).Value = headers(col)
        sheet.Cells(1, col + 1).Font.Bold = True
    Next col

    Dim rowIndex As Long: rowIndex = 2
    Dim row As Variant
    For Each row In finalReport
        sheet.Cells(rowIndex, 1).Value = row("Material")
        sheet.Cells(rowIndex, 2).Value = row("PlannedQty")
        sheet.Cells(rowIndex, 3).Value = row("ArrivedQty")
        sheet.Cells(rowIndex, 4).Value = row("RemainingQty")
        sheet.Cells(rowIndex, 5).Value = row("DeliveryPercent")
        sheet.Cells(rowIndex, 6).Value = row("Status")
        sheet.Cells(rowIndex, 7).Value = row("LastArrivalDate")
        sheet.Cells(rowIndex, 8).Value = row("Confidence")
        sheet.Cells(rowIndex, 9).Value = row("MatchRule")
        rowIndex = rowIndex + 1
    Next row

    AddTotalsRow sheet, rowIndex
    FormatReport sheet, rowIndex
End Sub

Private Sub AddTotalsRow(ByVal sheet As Worksheet, ByVal totalRow As Long)
    sheet.Cells(totalRow, 1).Value = "Totals"
    sheet.Cells(totalRow, 1).Font.Bold = True
    sheet.Cells(totalRow, 2).Formula = "=SUM(B2:B" & totalRow - 1 & ")"
    sheet.Cells(totalRow, 3).Formula = "=SUM(C2:C" & totalRow - 1 & ")"
    sheet.Cells(totalRow, 4).Formula = "=SUM(D2:D" & totalRow - 1 & ")"
    sheet.Cells(totalRow, 5).Formula = "=IF(B" & totalRow & ">0,C" & totalRow & "/B" & totalRow & ",0)"
    sheet.Rows(totalRow).Font.Bold = True
End Sub

Private Sub FormatReport(ByVal sheet As Worksheet, ByVal lastRow As Long)
    sheet.Range("A1:I" & lastRow).AutoFilter
    sheet.Range("A1:I" & lastRow).Columns.AutoFit
    sheet.Range("E:E,H:H").NumberFormat = "0.00%"
    sheet.Activate
    sheet.Range("A2").Select
    ActiveWindow.FreezePanes = True
    ApplyStatusFormatting sheet, lastRow
End Sub

Private Sub ApplyStatusFormatting(ByVal sheet As Worksheet, ByVal lastRow As Long)
    Dim statusRange As Range
    Set statusRange = sheet.Range("F2:F" & lastRow)
    statusRange.FormatConditions.Delete
    AddStatusCondition statusRange, "Completed", RGB(198, 239, 206)
    AddStatusCondition statusRange, "Partial", RGB(255, 235, 156)
    AddStatusCondition statusRange, "Missing", RGB(255, 199, 206)
End Sub

Private Sub AddStatusCondition(ByVal statusRange As Range, ByVal statusText As String, ByVal fillColor As Long)
    With statusRange.FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="=\"" & statusText & "\"")
        .Interior.Color = fillColor
    End With
End Sub

Private Sub WriteDashboard(ByVal sheet As Worksheet, ByVal dashboard As Object)
    sheet.Range("A1:B1").Value = Array("Metric", "Value")
    sheet.Range("A1:B1").Font.Bold = True
    sheet.Range("A2").Value = "Total planned": sheet.Range("B2").Value = dashboard("TotalPlanned")
    sheet.Range("A3").Value = "Total arrived": sheet.Range("B3").Value = dashboard("TotalArrived")
    sheet.Range("A4").Value = "Total remaining": sheet.Range("B4").Value = dashboard("TotalRemaining")
    sheet.Range("A5").Value = "Completion %": sheet.Range("B5").Value = dashboard("CompletionPercent")
    sheet.Range("A6").Value = "Number of matched materials": sheet.Range("B6").Value = dashboard("MatchedMaterials")
    sheet.Range("A7").Value = "Number requiring manual review": sheet.Range("B7").Value = dashboard("ManualReviewCount")
    sheet.Range("B5").NumberFormat = "0.00%"
    sheet.Range("A:B").Columns.AutoFit
End Sub

Private Function PrepareWorksheet(ByVal workbookToUse As Workbook, ByVal sheetName As String) As Worksheet
    Dim sheet As Worksheet
    On Error Resume Next
    Set sheet = workbookToUse.Worksheets(sheetName)
    On Error GoTo 0
    If sheet Is Nothing Then
        Set sheet = workbookToUse.Worksheets.Add(After:=workbookToUse.Worksheets(workbookToUse.Worksheets.Count))
        sheet.Name = sheetName
    Else
        sheet.Cells.Clear
    End If
    Set PrepareWorksheet = sheet
End Function

Private Sub LogExportStep(ByVal message As String)
    Debug.Print Format$(Now, "yyyy-mm-dd hh:nn:ss") & " [Export] " & message
End Sub
