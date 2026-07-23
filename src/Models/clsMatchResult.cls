VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMatchResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum MatchResultStatus
    MatchResultMatched = 1
    MatchResultPartiallyMatched = 2
    MatchResultNotMatched = 3
    MatchResultNeedManualReview = 4
End Enum

Private mStatus As MatchResultStatus
Private mMaterial As String
Private mPlannedQty As Double
Private mArrivedQty As Double
Private mRemainingQty As Double
Private mArrivalDate As Variant
Private mConfidence As Double
Private mMatchKey As String
Private mMatchRule As String

Public Property Get Status() As MatchResultStatus
    Status = mStatus
End Property

Public Property Let Status(ByVal value As MatchResultStatus)
    mStatus = value
End Property

Public Property Get Material() As String
    Material = mMaterial
End Property

Public Property Let Material(ByVal value As String)
    mMaterial = value
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

Public Property Get ArrivalDate() As Variant
    ArrivalDate = mArrivalDate
End Property

Public Property Let ArrivalDate(ByVal value As Variant)
    mArrivalDate = value
End Property

Public Property Get Confidence() As Double
    Confidence = mConfidence
End Property

Public Property Let Confidence(ByVal value As Double)
    mConfidence = value
End Property

Public Property Get MatchKey() As String
    MatchKey = mMatchKey
End Property

Public Property Let MatchKey(ByVal value As String)
    mMatchKey = value
End Property

Public Property Get MatchRule() As String
    MatchRule = mMatchRule
End Property

Public Property Let MatchRule(ByVal value As String)
    mMatchRule = value
End Property
