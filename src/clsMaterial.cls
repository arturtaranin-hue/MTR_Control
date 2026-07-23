VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMaterial"
Option Explicit

Private mCode As String
Private mName As String
Private mQuantity As Double
Private mSourceRow As Long

Public Property Get Code() As String
    Code = mCode
End Property

Public Property Let Code(ByVal value As String)
    mCode = value
End Property

Public Property Get Name() As String
    Name = mName
End Property

Public Property Let Name(ByVal value As String)
    mName = value
End Property

Public Property Get Quantity() As Double
    Quantity = mQuantity
End Property

Public Property Let Quantity(ByVal value As Double)
    mQuantity = value
End Property

Public Property Get SourceRow() As Long
    SourceRow = mSourceRow
End Property

Public Property Let SourceRow(ByVal value As Long)
    mSourceRow = value
End Property
