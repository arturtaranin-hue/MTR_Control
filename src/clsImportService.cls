VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsImportService"
Option Explicit

Public Function LoadDV(ByVal path As String) As Collection ' Collection(Of clsMaterial)
    Dim oldCalculation As XlCalculation
    Dim repository As clsDVRepository
    Dim errorNumber As Long
    Dim errorSource As String
    Dim errorDescription As String
    On Error GoTo CleanFail
    BeginFastImport oldCalculation
    LogImportStep "ImportService", "LoadDV requested"
    Set repository = New clsDVRepository
    Set LoadDV = repository.Load(path)
CleanExit:
    EndFastImport oldCalculation
    LogImportStep "ImportService", "LoadDV completed"
    Exit Function
CleanFail:
    errorNumber = Err.Number
    errorSource = Err.Source
    errorDescription = Err.Description
    LogImportStep "ImportService", "LoadDV failed: " & errorDescription
    EndFastImport oldCalculation
    Err.Raise errorNumber, errorSource, errorDescription
End Function

Public Function LoadMovement(ByVal path As String) As Collection ' Collection(Of clsMovement)
    Dim oldCalculation As XlCalculation
    Dim repository As clsMovementRepository
    Dim errorNumber As Long
    Dim errorSource As String
    Dim errorDescription As String
    On Error GoTo CleanFail
    BeginFastImport oldCalculation
    LogImportStep "ImportService", "LoadMovement requested"
    Set repository = New clsMovementRepository
    Set LoadMovement = repository.Load(path)
CleanExit:
    EndFastImport oldCalculation
    LogImportStep "ImportService", "LoadMovement completed"
    Exit Function
CleanFail:
    errorNumber = Err.Number
    errorSource = Err.Source
    errorDescription = Err.Description
    LogImportStep "ImportService", "LoadMovement failed: " & errorDescription
    EndFastImport oldCalculation
    Err.Raise errorNumber, errorSource, errorDescription
End Function

Public Function LoadAngara(ByVal path As String) As Collection ' Collection(Of clsMovement)
    Dim oldCalculation As XlCalculation
    Dim repository As clsAngaraRepository
    Dim errorNumber As Long
    Dim errorSource As String
    Dim errorDescription As String
    On Error GoTo CleanFail
    BeginFastImport oldCalculation
    LogImportStep "ImportService", "LoadAngara requested"
    Set repository = New clsAngaraRepository
    Set LoadAngara = repository.Load(path)
CleanExit:
    EndFastImport oldCalculation
    LogImportStep "ImportService", "LoadAngara completed"
    Exit Function
CleanFail:
    errorNumber = Err.Number
    errorSource = Err.Source
    errorDescription = Err.Description
    LogImportStep "ImportService", "LoadAngara failed: " & errorDescription
    EndFastImport oldCalculation
    Err.Raise errorNumber, errorSource, errorDescription
End Function

Private Sub BeginFastImport(ByRef oldCalculation As XlCalculation)
    oldCalculation = Application.Calculation
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
End Sub

Private Sub EndFastImport(ByVal oldCalculation As XlCalculation)
    Application.EnableEvents = True
    Application.Calculation = oldCalculation
    Application.ScreenUpdating = True
End Sub
