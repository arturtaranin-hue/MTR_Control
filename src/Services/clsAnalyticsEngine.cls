VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsAnalyticsEngine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function CalculateStatistics(ByVal arrivalResults As Collection, Optional ByVal asOfDate As Variant) As clsExecutionStatistics
    Dim statistics As New clsExecutionStatistics
    Dim reportDate As Date
    If IsDate(asOfDate) Then reportDate = CDate(asOfDate) Else reportDate = Date

    Dim result As clsArrivalResult
    For Each result In arrivalResults
        statistics.TotalPlanned = statistics.TotalPlanned + result.PlannedQty
        statistics.TotalArrived = statistics.TotalArrived + result.ArrivedQty
        statistics.TotalRemaining = statistics.TotalRemaining + result.RemainingQty

        AddDailyArrival statistics.DailyArrivals, result
        AddSupplierStatistic statistics.SupplierStatistics, result

        If result.RemainingQty > 0 Or result.Status = NOT_ARRIVED Or result.Status = PARTIALLY_ARRIVED Then
            statistics.OpenDeliveries.Add result
        End If

        If IsDate(result.RequiredDate) And result.RemainingQty > 0 Then
            If CDate(result.RequiredDate) < reportDate Then statistics.LateDeliveries.Add result
        End If
    Next result

    If statistics.TotalPlanned > 0 Then
        statistics.CompletionPercent = statistics.TotalArrived / statistics.TotalPlanned
    Else
        statistics.CompletionPercent = 0#
    End If

    Set CalculateStatistics = statistics
End Function

Private Sub AddDailyArrival(ByVal dailyArrivals As Object, ByVal result As clsArrivalResult)
    If Not IsDate(result.LastArrivalDate) Or result.ArrivedQty = 0 Then Exit Sub

    Dim dayKey As String
    dayKey = Format$(CDate(result.LastArrivalDate), "yyyy-mm-dd")
    If Not dailyArrivals.Exists(dayKey) Then dailyArrivals.Add dayKey, 0#
    dailyArrivals(dayKey) = CDbl(dailyArrivals(dayKey)) + result.ArrivedQty
End Sub

Private Sub AddSupplierStatistic(ByVal supplierStatistics As Object, ByVal result As clsArrivalResult)
    Dim supplierKey As String
    supplierKey = Trim$(result.Supplier)
    If Len(supplierKey) = 0 Then supplierKey = "UNKNOWN"

    Dim bucket As Object
    If supplierStatistics.Exists(supplierKey) Then
        Set bucket = supplierStatistics(supplierKey)
    Else
        Set bucket = CreateObject("Scripting.Dictionary")
        bucket("planned") = 0#
        bucket("arrived") = 0#
        bucket("remaining") = 0#
        bucket("openCount") = 0&
        bucket("lateCount") = 0&
        supplierStatistics.Add supplierKey, bucket
    End If

    bucket("planned") = CDbl(bucket("planned")) + result.PlannedQty
    bucket("arrived") = CDbl(bucket("arrived")) + result.ArrivedQty
    bucket("remaining") = CDbl(bucket("remaining")) + result.RemainingQty
    If result.RemainingQty > 0 Then bucket("openCount") = CLng(bucket("openCount")) + 1
    If IsDate(result.RequiredDate) And result.RemainingQty > 0 Then
        If CDate(result.RequiredDate) < Date Then bucket("lateCount") = CLng(bucket("lateCount")) + 1
    End If
End Sub
