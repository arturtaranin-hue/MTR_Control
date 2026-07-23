VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsArrivalResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum ArrivalStatus
    ARRIVED = 1
    PARTIALLY_ARRIVED = 2
    NOT_ARRIVED = 3
    OVER_DELIVERED = 4
    MANUAL_REVIEW = 5
End Enum

Private mStatus As ArrivalStatus
Private mMaterial As String
Private mSapCode As String
Private mMaterialCode As String
Private mPlannedQty As Double
Private mArrivedQty As Double
Private mRemainingQty As Double
Private mCompletionPercent As Double
Private mFirstArrivalDate As Variant
Private mLastArrivalDate As Variant
Private mDeliveryCount As Long
Private mSupplier As String
Private mRequiredDate As Variant
Private mMatchKey As String
Private mReviewReason As String
Private mUnexpectedArrival As Boolean

Public Property Get Status() As ArrivalStatus
    Status = mStatus
End Property

Public Property Let Status(ByVal value As ArrivalStatus)
    mStatus = value
End Property

Public Property Get Material() As String
    Material = mMaterial
End Property

Public Property Let Material(ByVal value As String)
    mMaterial = value
End Property

Public Property Get SapCode() As String
    SapCode = mSapCode
End Property

Public Property Let SapCode(ByVal value As String)
    mSapCode = value
End Property

Public Property Get MaterialCode() As String
    MaterialCode = mMaterialCode
End Property

Public Property Let MaterialCode(ByVal value As String)
    mMaterialCode = value
End Property

Public Property Get PlannedQty() As Double
    PlannedQty = mPlannedQty
End Property

Public Property Let PlannedQty(ByVal value As Double)
    mPlannedQty = value
End Property

Public Property Get ArrivedQty() As Double
    ArrivedQty = mArrivedQty
End Property

Public Property Let ArrivedQty(ByVal value As Double)
    mArrivedQty = value
End Property

Public Property Get RemainingQty() As Double
    RemainingQty = mRemainingQty
End Property

Public Property Let RemainingQty(ByVal value As Double)
    mRemainingQty = value
End Property

Public Property Get CompletionPercent() As Double
    CompletionPercent = mCompletionPercent
End Property

Public Property Let CompletionPercent(ByVal value As Double)
    mCompletionPercent = value
End Property

Public Property Get FirstArrivalDate() As Variant
    FirstArrivalDate = mFirstArrivalDate
End Property

Public Property Let FirstArrivalDate(ByVal value As Variant)
    mFirstArrivalDate = value
End Property

Public Property Get LastArrivalDate() As Variant
    LastArrivalDate = mLastArrivalDate
End Property

Public Property Let LastArrivalDate(ByVal value As Variant)
    mLastArrivalDate = value
End Property

Public Property Get DeliveryCount() As Long
    DeliveryCount = mDeliveryCount
End Property

Public Property Let DeliveryCount(ByVal value As Long)
    mDeliveryCount = value
End Property

Public Property Get Supplier() As String
    Supplier = mSupplier
End Property

Public Property Let Supplier(ByVal value As String)
    mSupplier = value
End Property

Public Property Get RequiredDate() As Variant
    RequiredDate = mRequiredDate
End Property

Public Property Let RequiredDate(ByVal value As Variant)
    mRequiredDate = value
End Property

Public Property Get MatchKey() As String
    MatchKey = mMatchKey
End Property

Public Property Let MatchKey(ByVal value As String)
    mMatchKey = value
End Property

Public Property Get ReviewReason() As String
    ReviewReason = mReviewReason
End Property

Public Property Let ReviewReason(ByVal value As String)
    mReviewReason = value
End Property

Public Property Get UnexpectedArrival() As Boolean
    UnexpectedArrival = mUnexpectedArrival
End Property

Public Property Let UnexpectedArrival(ByVal value As Boolean)
    mUnexpectedArrival = value
End Property
