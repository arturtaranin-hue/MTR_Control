VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsImportMapping"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mSapCodeColumn As Long
Private mMaterialCodeColumn As Long
Private mMaterialNameColumn As Long
Private mQuantityColumn As Long
Private mArrivalDateColumn As Long

Private Sub Class_Initialize()
    mSapCodeColumn = 1
    mMaterialCodeColumn = 2
    mMaterialNameColumn = 3
    mQuantityColumn = 4
    mArrivalDateColumn = 5
End Sub

Public Property Get SapCodeColumn() As Long
    SapCodeColumn = mSapCodeColumn
End Property

Public Property Let SapCodeColumn(ByVal value As Long)
    mSapCodeColumn = value
End Property

Public Property Get MaterialCodeColumn() As Long
    MaterialCodeColumn = mMaterialCodeColumn
End Property

Public Property Let MaterialCodeColumn(ByVal value As Long)
    mMaterialCodeColumn = value
End Property

Public Property Get MaterialNameColumn() As Long
    MaterialNameColumn = mMaterialNameColumn
End Property

Public Property Let MaterialNameColumn(ByVal value As Long)
    mMaterialNameColumn = value
End Property

Public Property Get QuantityColumn() As Long
    QuantityColumn = mQuantityColumn
End Property

Public Property Let QuantityColumn(ByVal value As Long)
    mQuantityColumn = value
End Property

Public Property Get ArrivalDateColumn() As Long
    ArrivalDateColumn = mArrivalDateColumn
End Property

Public Property Let ArrivalDateColumn(ByVal value As Long)
    mArrivalDateColumn = value
End Property
