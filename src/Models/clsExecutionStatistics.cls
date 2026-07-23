VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsExecutionStatistics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mTotalPlanned As Double
Private mTotalArrived As Double
Private mTotalRemaining As Double
Private mCompletionPercent As Double
Private mDailyArrivals As Object
Private mOpenDeliveries As Collection
Private mLateDeliveries As Collection
Private mSupplierStatistics As Object

Private Sub Class_Initialize()
    Set mDailyArrivals = CreateObject("Scripting.Dictionary")
    Set mOpenDeliveries = New Collection
    Set mLateDeliveries = New Collection
    Set mSupplierStatistics = CreateObject("Scripting.Dictionary")
End Sub

Public Property Get TotalPlanned() As Double
    TotalPlanned = mTotalPlanned
End Property

Public Property Let TotalPlanned(ByVal value As Double)
    mTotalPlanned = value
End Property

Public Property Get TotalArrived() As Double
    TotalArrived = mTotalArrived
End Property

Public Property Let TotalArrived(ByVal value As Double)
    mTotalArrived = value
End Property

Public Property Get TotalRemaining() As Double
    TotalRemaining = mTotalRemaining
End Property

Public Property Let TotalRemaining(ByVal value As Double)
    mTotalRemaining = value
End Property

Public Property Get CompletionPercent() As Double
    CompletionPercent = mCompletionPercent
End Property

Public Property Let CompletionPercent(ByVal value As Double)
    mCompletionPercent = value
End Property

Public Property Get DailyArrivals() As Object
    Set DailyArrivals = mDailyArrivals
End Property

Public Property Get OpenDeliveries() As Collection
    Set OpenDeliveries = mOpenDeliveries
End Property

Public Property Get LateDeliveries() As Collection
    Set LateDeliveries = mLateDeliveries
End Property

Public Property Get SupplierStatistics() As Object
    Set SupplierStatistics = mSupplierStatistics
End Property
