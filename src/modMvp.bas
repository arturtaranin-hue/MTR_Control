Attribute VB_Name = "modMvp"
Option Explicit

Private Const APP_WORKBOOK_NAME As String = "MTR_Control.xlsm"
Private Const DASHBOARD_SHEET As String = "Dashboard"
Private Const LOG_SHEET As String = "Log"
Private Const REPORT_SHEET As String = "Report"

Public Sub Run()
    Dim oldScreenUpdating As Boolean
    Dim oldEnableEvents As Boolean
    Dim oldDisplayAlerts As Boolean
    Dim oldCalculation As XlCalculation
    Dim dvPath As String
    Dim movementPath As String
    Dim angaraPath As String
    Dim dvItems As Collection
    Dim movementItems As Collection
    Dim angaraItems As Collection
    Dim dvRows As Variant
    Dim movementRows As Variant
    Dim angaraRows As Variant
    Dim normalizedDvRows As Variant
    Dim normalizedMovementRows As Variant
    Dim normalizedAngaraRows As Variant
    Dim matches As Collection
    Dim arrivals As Collection
    Dim states As Collection
    Dim finalReport As Collection
    Dim targetWorkbook As Workbook
    Dim reportPath As String

    On Error GoTo Fail
    Set targetWorkbook = EnsureMvpWorkbook()
    CaptureApplicationState oldScreenUpdating, oldEnableEvents, oldDisplayAlerts, oldCalculation
    ConfigureApplicationForRun
    LogMvpStep "Run", "Starting complete MVP pipeline"

    dvPath = PickSourceFile("Select DV Angara file")
    movementPath = PickSourceFile("Select Movement Summary file")
    angaraPath = PickSourceFile("Select GPN Angara file")

    Dim importer As New clsImportService
    Set dvItems = importer.LoadDV(dvPath)
    Set movementItems = importer.LoadMovement(movementPath)
    Set angaraItems = importer.LoadAngara(angaraPath)

    dvRows = MaterialsToRows(dvItems)
    movementRows = MovementsToRows(movementItems)
    angaraRows = MovementsToRows(angaraItems)

    normalizedDvRows = NormalizeRows(dvRows)
    normalizedMovementRows = NormalizeRows(movementRows)
    normalizedAngaraRows = NormalizeRows(angaraRows)

    Dim matcher As New clsMatchingService
    matcher.Initialize
    Set matches = matcher.MatchMaterials(normalizedDvRows, normalizedMovementRows)
    LogMvpStep "Match", "Matched rows: " & CStr(matches.Count)

    Dim arrivalEngine As New clsArrivalEngine
    arrivalEngine.Initialize
    Set arrivals = arrivalEngine.CalculateArrivals(normalizedDvRows, normalizedMovementRows)
    LogMvpStep "Arrival", "Calculated arrival rows: " & CStr(arrivals.Count)

    Dim reconciliation As New clsReconciliationService
    reconciliation.Initialize
    Set states = reconciliation.Reconcile(normalizedDvRows, normalizedMovementRows, normalizedAngaraRows)
    LogMvpStep "Reconciliation", "Reconciled states: " & CStr(states.Count)

    Dim reportService As New clsReportService
    Set finalReport = reportService.BuildFinalReport(states)

    Dim exporter As New clsExcelExporter
    exporter.ExportFinalReport finalReport, targetWorkbook
    reportPath = SaveReport(targetWorkbook)

    LogMvpStep "Run", "Pipeline completed successfully. Report saved: " & reportPath
    MsgBox "MTR Control report generated successfully." & vbCrLf & reportPath, vbInformation, "MTR Control"

CleanExit:
    RestoreApplicationState oldScreenUpdating, oldEnableEvents, oldDisplayAlerts, oldCalculation
    Exit Sub
Fail:
    LogMvpStep "Error", "Runtime error " & Err.Number & " from " & Err.Source & ": " & Err.Description
    MsgBox "MTR Control failed: " & Err.Description, vbCritical, "MTR Control"
    Resume CleanExit
End Sub

Public Sub ImportFiles()
    Run
End Sub

Public Sub GenerateReport()
    Run
End Sub

Public Sub OpenLog()
    On Error GoTo Fail
    EnsureLogSheet(EnsureMvpWorkbook()).Activate
    Exit Sub
Fail:
    MsgBox "Unable to open log: " & Err.Description, vbExclamation, "MTR Control"
End Sub

Public Function CreateWorkbook() As Workbook
    Dim wb As Workbook
    Set wb = Workbooks.Add(xlWBATWorksheet)
    wb.Worksheets(1).Name = DASHBOARD_SHEET
    SetupDashboard wb
    wb.SaveAs Filename:=ThisWorkbook.Path & Application.PathSeparator & APP_WORKBOOK_NAME, FileFormat:=xlOpenXMLWorkbookMacroEnabled
    Set CreateWorkbook = wb
End Function

Public Function EnsureMvpWorkbook() As Workbook
    If Not ActiveWorkbook Is Nothing Then
        If StrComp(ActiveWorkbook.Name, APP_WORKBOOK_NAME, vbTextCompare) = 0 Then
            Set EnsureMvpWorkbook = ActiveWorkbook
            SetupDashboard EnsureMvpWorkbook
            Exit Function
        End If
    End If
    Set EnsureMvpWorkbook = ThisWorkbook
    SetupDashboard EnsureMvpWorkbook
End Function

Public Sub SetupDashboard(ByVal workbookToUse As Workbook)
    Dim dashboard As Worksheet
    Set dashboard = EnsureWorksheet(workbookToUse, DASHBOARD_SHEET)
    dashboard.Range("A1").Value = "MTR Control"
    dashboard.Range("A2").Value = "Use the buttons below to import files, generate the MVP report, or open the runtime log."
    dashboard.Range("A1").Font.Bold = True
    dashboard.Range("A1").Font.Size = 18
    dashboard.Columns("A:D").ColumnWidth = 24
    EnsureButton dashboard, "Import Files", "ImportFiles", 30, 70
    EnsureButton dashboard, "Generate Report", "GenerateReport", 30, 115
    EnsureButton dashboard, "Open Log", "OpenLog", 30, 160
    EnsureLogSheet workbookToUse
End Sub

Private Function PickSourceFile(ByVal title As String) As String
    With Application.FileDialog(msoFileDialogFilePicker)
        .Title = title
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Excel files", "*.xlsx;*.xlsm;*.xlsb;*.xls"
        If .Show <> -1 Then Err.Raise vbObjectError + 5400, "modMvp.PickSourceFile", title & " was cancelled"
        PickSourceFile = .SelectedItems(1)
    End With
    LogMvpStep "Import", title & ": " & PickSourceFile
End Function

Private Function NormalizeRows(ByVal rows As Variant) As Variant
    Dim normalizer As New clsNormalizeService
    normalizer.Initialize
    Dim result As Variant
    result = rows
    Dim rowIndex As Long
    For rowIndex = LBound(result, 1) To UBound(result, 1)
        result(rowIndex, 3) = normalizer.NormalizeText(result(rowIndex, 3))
    Next rowIndex
    NormalizeRows = result
End Function

Private Function MaterialsToRows(ByVal materials As Collection) As Variant
    Dim rows() As Variant
    ReDim rows(1 To materials.Count, 1 To 5)
    Dim index As Long
    Dim material As clsMaterial
    For index = 1 To materials.Count
        Set material = materials(index)
        rows(index, 1) = vbNullString
        rows(index, 2) = material.Code
        rows(index, 3) = material.Name
        rows(index, 4) = material.Quantity
        rows(index, 5) = Empty
    Next index
    MaterialsToRows = rows
End Function

Private Function MovementsToRows(ByVal movements As Collection) As Variant
    Dim rows() As Variant
    ReDim rows(1 To movements.Count, 1 To 5)
    Dim index As Long
    Dim movement As clsMovement
    For index = 1 To movements.Count
        Set movement = movements(index)
        rows(index, 1) = vbNullString
        rows(index, 2) = movement.MaterialCode
        rows(index, 3) = movement.MaterialCode
        rows(index, 4) = movement.Quantity
        rows(index, 5) = movement.MovementDate
    Next index
    MovementsToRows = rows
End Function

Private Function SaveReport(ByVal workbookToUse As Workbook) As String
    Dim outputPath As String
    outputPath = workbookToUse.Path
    If Len(outputPath) = 0 Then outputPath = ThisWorkbook.Path
    If Len(outputPath) = 0 Then outputPath = CurDir$
    SaveReport = outputPath & Application.PathSeparator & "MTR_Control_Report_" & Format$(Now, "yyyymmdd_hhnnss") & ".xlsm"
    workbookToUse.SaveCopyAs SaveReport
End Function

Private Sub EnsureButton(ByVal sheet As Worksheet, ByVal caption As String, ByVal macroName As String, ByVal leftPos As Double, ByVal topPos As Double)
    Dim shape As Shape
    For Each shape In sheet.Shapes
        If shape.Name = "btn" & Replace(caption, " ", vbNullString) Then shape.Delete: Exit For
    Next shape
    Set shape = sheet.Shapes.AddFormControl(xlButtonControl, leftPos, topPos, 150, 30)
    shape.Name = "btn" & Replace(caption, " ", vbNullString)
    shape.TextFrame.Characters.Text = caption
    shape.OnAction = macroName
End Sub

Private Function EnsureWorksheet(ByVal workbookToUse As Workbook, ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set EnsureWorksheet = workbookToUse.Worksheets(sheetName)
    On Error GoTo 0
    If EnsureWorksheet Is Nothing Then
        Set EnsureWorksheet = workbookToUse.Worksheets.Add(After:=workbookToUse.Worksheets(workbookToUse.Worksheets.Count))
        EnsureWorksheet.Name = sheetName
    End If
End Function

Private Function EnsureLogSheet(ByVal workbookToUse As Workbook) As Worksheet
    Set EnsureLogSheet = EnsureWorksheet(workbookToUse, LOG_SHEET)
    If Len(CStr(EnsureLogSheet.Range("A1").Value)) = 0 Then EnsureLogSheet.Range("A1:C1").Value = Array("Timestamp", "Step", "Message")
End Function

Private Sub LogMvpStep(ByVal stepName As String, ByVal message As String)
    Debug.Print Format$(Now, "yyyy-mm-dd hh:nn:ss") & " [MVP] [" & stepName & "] " & message
    On Error Resume Next
    Dim logSheet As Worksheet
    Set logSheet = EnsureLogSheet(EnsureMvpWorkbook())
    Dim nextRow As Long
    nextRow = logSheet.Cells(logSheet.Rows.Count, 1).End(xlUp).Row + 1
    logSheet.Cells(nextRow, 1).Value = Now
    logSheet.Cells(nextRow, 2).Value = stepName
    logSheet.Cells(nextRow, 3).Value = message
    On Error GoTo 0
End Sub

Private Sub CaptureApplicationState(ByRef oldScreenUpdating As Boolean, ByRef oldEnableEvents As Boolean, ByRef oldDisplayAlerts As Boolean, ByRef oldCalculation As XlCalculation)
    oldScreenUpdating = Application.ScreenUpdating
    oldEnableEvents = Application.EnableEvents
    oldDisplayAlerts = Application.DisplayAlerts
    oldCalculation = Application.Calculation
End Sub

Private Sub ConfigureApplicationForRun()
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
End Sub

Private Sub RestoreApplicationState(ByVal oldScreenUpdating As Boolean, ByVal oldEnableEvents As Boolean, ByVal oldDisplayAlerts As Boolean, ByVal oldCalculation As XlCalculation)
    Application.ScreenUpdating = oldScreenUpdating
    Application.EnableEvents = oldEnableEvents
    Application.DisplayAlerts = oldDisplayAlerts
    Application.Calculation = oldCalculation
End Sub
