VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMaterial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mMaterialId As String
Private mMaterialCode As String
Private mDescription As String
Private mUnitOfMeasure As String

Public Property Get MaterialId() As String
    MaterialId = mMaterialId
End Property

Public Property Let MaterialId(ByVal value As String)
    mMaterialId = value
End Property

Public Property Get MaterialCode() As String
    MaterialCode = mMaterialCode
End Property

Public Property Let MaterialCode(ByVal value As String)
    mMaterialCode = value
End Property

Public Property Get Description() As String
    Description = mDescription
End Property

Public Property Let Description(ByVal value As String)
    mDescription = value
End Property

Public Property Get UnitOfMeasure() As String
    UnitOfMeasure = mUnitOfMeasure
End Property

Public Property Let UnitOfMeasure(ByVal value As String)
    mUnitOfMeasure = value
End Property
