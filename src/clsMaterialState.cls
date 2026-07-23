VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMaterialState"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mStatus As MaterialReconciliationStatus
Private mSapCode As String
Private mMaterialCode As String
Private mMaterial As String
Private mPlannedQty As Double
Private mArrivedQty As Double
Private mRemainingQty As Double
Private mDeliveryPercent As Double
Private mMovementQty As Double
Private mAngaraQty As Double
Private mArrivalCount As Long
Private mDuplicateArrival As Boolean
Private mNegativeBalance As Boolean
Private mOverDelivery As Boolean
Private mUnknownMaterial As Boolean
Private mManualReview As Boolean
Private mReviewReason As String
Private mMatchKey As String
Private mSourceRow As Long

Public Property Get Status() As MaterialReconciliationStatus: Status = mStatus: End Property
Public Property Let Status(ByVal value As MaterialReconciliationStatus): mStatus = value: End Property
Public Property Get SapCode() As String: SapCode = mSapCode: End Property
Public Property Let SapCode(ByVal value As String): mSapCode = value: End Property
Public Property Get MaterialCode() As String: MaterialCode = mMaterialCode: End Property
Public Property Let MaterialCode(ByVal value As String): mMaterialCode = value: End Property
Public Property Get Material() As String: Material = mMaterial: End Property
Public Property Let Material(ByVal value As String): mMaterial = value: End Property
Public Property Get PlannedQty() As Double: PlannedQty = mPlannedQty: End Property
Public Property Let PlannedQty(ByVal value As Double): mPlannedQty = value: End Property
Public Property Get ArrivedQty() As Double: ArrivedQty = mArrivedQty: End Property
Public Property Let ArrivedQty(ByVal value As Double): mArrivedQty = value: End Property
Public Property Get RemainingQty() As Double: RemainingQty = mRemainingQty: End Property
Public Property Let RemainingQty(ByVal value As Double): mRemainingQty = value: End Property
Public Property Get DeliveryPercent() As Double: DeliveryPercent = mDeliveryPercent: End Property
Public Property Let DeliveryPercent(ByVal value As Double): mDeliveryPercent = value: End Property
Public Property Get MovementQty() As Double: MovementQty = mMovementQty: End Property
Public Property Let MovementQty(ByVal value As Double): mMovementQty = value: End Property
Public Property Get AngaraQty() As Double: AngaraQty = mAngaraQty: End Property
Public Property Let AngaraQty(ByVal value As Double): mAngaraQty = value: End Property
Public Property Get ArrivalCount() As Long: ArrivalCount = mArrivalCount: End Property
Public Property Let ArrivalCount(ByVal value As Long): mArrivalCount = value: End Property
Public Property Get DuplicateArrival() As Boolean: DuplicateArrival = mDuplicateArrival: End Property
Public Property Let DuplicateArrival(ByVal value As Boolean): mDuplicateArrival = value: End Property
Public Property Get NegativeBalance() As Boolean: NegativeBalance = mNegativeBalance: End Property
Public Property Let NegativeBalance(ByVal value As Boolean): mNegativeBalance = value: End Property
Public Property Get OverDelivery() As Boolean: OverDelivery = mOverDelivery: End Property
Public Property Let OverDelivery(ByVal value As Boolean): mOverDelivery = value: End Property
Public Property Get UnknownMaterial() As Boolean: UnknownMaterial = mUnknownMaterial: End Property
Public Property Let UnknownMaterial(ByVal value As Boolean): mUnknownMaterial = value: End Property
Public Property Get ManualReview() As Boolean: ManualReview = mManualReview: End Property
Public Property Let ManualReview(ByVal value As Boolean): mManualReview = value: End Property
Public Property Get ReviewReason() As String: ReviewReason = mReviewReason: End Property
Public Property Let ReviewReason(ByVal value As String): mReviewReason = value: End Property
Public Property Get MatchKey() As String: MatchKey = mMatchKey: End Property
Public Property Let MatchKey(ByVal value As String): mMatchKey = value: End Property
Public Property Get SourceRow() As Long: SourceRow = mSourceRow: End Property
Public Property Let SourceRow(ByVal value As Long): mSourceRow = value: End Property
