VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMovement"
Option Explicit

Private mMaterialCode As String
Private mMovementDate As Date
Private mQuantity As Double
Private mMovementType As String
Private mSourceRow As Long

Public Property Get MaterialCode() As String
    MaterialCode = mMaterialCode
End Property

Public Property Let MaterialCode(ByVal value As String)
    mMaterialCode = value
End Property

Public Property Get MovementDate() As Date
    MovementDate = mMovementDate
End Property

Public Property Let MovementDate(ByVal value As Date)
    mMovementDate = value
End Property

Public Property Get Quantity() As Double
    Quantity = mQuantity
End Property

Public Property Let Quantity(ByVal value As Double)
    mQuantity = value
End Property

Public Property Get MovementType() As String
    MovementType = mMovementType
End Property

Public Property Let MovementType(ByVal value As String)
    mMovementType = value
End Property

Public Property Get SourceRow() As Long
    SourceRow = mSourceRow
End Property

Public Property Let SourceRow(ByVal value As Long)
    mSourceRow = value
End Property
