VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsArrival"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mArrivalId As String
Private mMaterialId As String
Private mArrivalDate As Date
Private mQuantity As Double
Private mReferenceNumber As String

Public Property Get ArrivalId() As String
    ArrivalId = mArrivalId
End Property

Public Property Let ArrivalId(ByVal value As String)
    mArrivalId = value
End Property

Public Property Get MaterialId() As String
    MaterialId = mMaterialId
End Property

Public Property Let MaterialId(ByVal value As String)
    mMaterialId = value
End Property

Public Property Get ArrivalDate() As Date
    ArrivalDate = mArrivalDate
End Property

Public Property Let ArrivalDate(ByVal value As Date)
    mArrivalDate = value
End Property

Public Property Get Quantity() As Double
    Quantity = mQuantity
End Property

Public Property Let Quantity(ByVal value As Double)
    mQuantity = value
End Property

Public Property Get ReferenceNumber() As String
    ReferenceNumber = mReferenceNumber
End Property

Public Property Let ReferenceNumber(ByVal value As String)
    mReferenceNumber = value
End Property
