VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsReportBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function BuildSummary(ByVal states As Collection) As Object
    Dim summary As Object
    Set summary = CreateObject("Scripting.Dictionary")
    summary("TotalMaterials") = 0&
    summary("TotalPlannedQty") = 0#
    summary("TotalArrivedQty") = 0#
    summary("TotalRemainingQty") = 0#
    summary("DeliveryPercent") = 0#
    summary("PlannedCount") = 0&
    summary("PartialCount") = 0&
    summary("DeliveredCount") = 0&
    summary("OverdeliveredCount") = 0&
    summary("NotFoundCount") = 0&
    summary("ManualReviewCount") = 0&
    summary("DuplicateArrivalCount") = 0&
    summary("NegativeBalanceCount") = 0&
    summary("UnknownMaterialCount") = 0&

    Dim state As clsMaterialState
    For Each state In states
        summary("TotalMaterials") = CLng(summary("TotalMaterials")) + 1
        summary("TotalPlannedQty") = CDbl(summary("TotalPlannedQty")) + state.PlannedQty
        summary("TotalArrivedQty") = CDbl(summary("TotalArrivedQty")) + state.ArrivedQty
        summary("TotalRemainingQty") = CDbl(summary("TotalRemainingQty")) + state.RemainingQty
        CountStatus summary, state.Status
        If state.DuplicateArrival Then summary("DuplicateArrivalCount") = CLng(summary("DuplicateArrivalCount")) + 1
        If state.NegativeBalance Then summary("NegativeBalanceCount") = CLng(summary("NegativeBalanceCount")) + 1
        If state.UnknownMaterial Then summary("UnknownMaterialCount") = CLng(summary("UnknownMaterialCount")) + 1
    Next state

    If CDbl(summary("TotalPlannedQty")) > 0# Then summary("DeliveryPercent") = CDbl(summary("TotalArrivedQty")) / CDbl(summary("TotalPlannedQty"))
    Set BuildSummary = summary
End Function

Public Function BuildRows(ByVal states As Collection) As Variant
    Dim rows() As Variant
    ReDim rows(1 To states.Count + 1, 1 To 12)
    rows(1, 1) = "SAP Code": rows(1, 2) = "Material Code": rows(1, 3) = "Material"
    rows(1, 4) = "Status": rows(1, 5) = "PlannedQty": rows(1, 6) = "ArrivedQty"
    rows(1, 7) = "RemainingQty": rows(1, 8) = "DeliveryPercent": rows(1, 9) = "MovementQty"
    rows(1, 10) = "AngaraQty": rows(1, 11) = "ManualReview": rows(1, 12) = "ReviewReason"

    Dim helper As New clsMaterialStatus
    Dim index As Long: index = 2
    Dim state As clsMaterialState
    For Each state In states
        rows(index, 1) = state.SapCode
        rows(index, 2) = state.MaterialCode
        rows(index, 3) = state.Material
        rows(index, 4) = helper.StatusName(state.Status)
        rows(index, 5) = state.PlannedQty
        rows(index, 6) = state.ArrivedQty
        rows(index, 7) = state.RemainingQty
        rows(index, 8) = state.DeliveryPercent
        rows(index, 9) = state.MovementQty
        rows(index, 10) = state.AngaraQty
        rows(index, 11) = state.ManualReview
        rows(index, 12) = state.ReviewReason
        index = index + 1
    Next state

    BuildRows = rows
End Function

Private Sub CountStatus(ByVal summary As Object, ByVal status As MaterialReconciliationStatus)
    Select Case status
        Case PLANNED: summary("PlannedCount") = CLng(summary("PlannedCount")) + 1
        Case PARTIAL: summary("PartialCount") = CLng(summary("PartialCount")) + 1
        Case DELIVERED: summary("DeliveredCount") = CLng(summary("DeliveredCount")) + 1
        Case OVERDELIVERED: summary("OverdeliveredCount") = CLng(summary("OverdeliveredCount")) + 1
        Case NOT_FOUND: summary("NotFoundCount") = CLng(summary("NotFoundCount")) + 1
        Case MANUAL_REVIEW: summary("ManualReviewCount") = CLng(summary("ManualReviewCount")) + 1
    End Select
End Sub
