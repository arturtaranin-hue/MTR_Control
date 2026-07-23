VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMovement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mMovementId As String
Private mMaterialId As String
Private mMovementDate As Date
Private mQuantity As Double
Private mSourceLocation As String
Private mDestinationLocation As String

Public Property Get MovementId() As String
    MovementId = mMovementId
End Property

Public Property Let MovementId(ByVal value As String)
    mMovementId = value
End Property

Public Property Get MaterialId() As String
    MaterialId = mMaterialId
End Property

Public Property Let MaterialId(ByVal value As String)
    mMaterialId = value
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

Public Property Get SourceLocation() As String
    SourceLocation = mSourceLocation
End Property

Public Property Let SourceLocation(ByVal value As String)
    mSourceLocation = value
End Property

Public Property Get DestinationLocation() As String
    DestinationLocation = mDestinationLocation
End Property

Public Property Let DestinationLocation(ByVal value As String)
    mDestinationLocation = value
End Property
