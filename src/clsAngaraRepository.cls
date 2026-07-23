VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsAngaraRepository"
Option Explicit

Public Function Load(ByVal path As String) As Collection
    Dim movementRepository As clsMovementRepository
    LogImportStep "AngaraRepository", "Delegating Angara import to movement parser"
    Set movementRepository = New clsMovementRepository
    Set Load = movementRepository.LoadMovementWorkbook(path, "AngaraRepository", Array("Angara", "Movement", "Movements"))
End Function
