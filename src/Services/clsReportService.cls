VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsReportService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function BuildFinalReport(ByVal results As Collection) As Collection
    Dim grouped As Object
    Set grouped = CreateObject("Scripting.Dictionary")
    grouped.CompareMode = vbTextCompare

    Dim item As Variant
    For Each item In results
        AddResult grouped, item
    Next item

    Dim finalReport As New Collection
    Dim key As Variant
    For Each key In grouped.Keys
        FinalizeRow grouped(key)
        finalReport.Add grouped(key)
    Next key

    Set BuildFinalReport = finalReport
End Function

Public Function BuildDashboard(ByVal finalReport As Collection) As Object
    Dim dashboard As Object
    Set dashboard = CreateObject("Scripting.Dictionary")
    dashboard("TotalPlanned") = 0#
    dashboard("TotalArrived") = 0#
    dashboard("TotalRemaining") = 0#
    dashboard("CompletionPercent") = 0#
    dashboard("MatchedMaterials") = 0&
    dashboard("ManualReviewCount") = 0&

    Dim row As Variant
    For Each row In finalReport
        dashboard("TotalPlanned") = CDbl(dashboard("TotalPlanned")) + CDbl(row("PlannedQty"))
        dashboard("TotalArrived") = CDbl(dashboard("TotalArrived")) + CDbl(row("ArrivedQty"))
        dashboard("TotalRemaining") = CDbl(dashboard("TotalRemaining")) + CDbl(row("RemainingQty"))
        If CDbl(row("ArrivedQty")) > 0# Then dashboard("MatchedMaterials") = CLng(dashboard("MatchedMaterials")) + 1
        If RequiresManualReview(row) Then dashboard("ManualReviewCount") = CLng(dashboard("ManualReviewCount")) + 1
    Next row

    If CDbl(dashboard("TotalPlanned")) > 0# Then dashboard("CompletionPercent") = CDbl(dashboard("TotalArrived")) / CDbl(dashboard("TotalPlanned"))
    Set BuildDashboard = dashboard
End Function

Private Sub AddResult(ByVal grouped As Object, ByVal source As Object)
    Dim material As String
    material = Trim$(CStr(source.Material))
    If Len(material) = 0 Then material = "(Unspecified material)"

    Dim row As Object
    If grouped.Exists(material) Then
        Set row = grouped(material)
    Else
        Set row = CreateObject("Scripting.Dictionary")
        row("Material") = material
        row("PlannedQty") = 0#
        row("ArrivedQty") = 0#
        row("RemainingQty") = 0#
        row("DeliveryPercent") = 0#
        row("Status") = "Missing"
        row("LastArrivalDate") = Empty
        row("Confidence") = 0#
        row("MatchRule") = ""
        row("ManualReview") = False
        grouped.Add material, row
    End If

    row("PlannedQty") = CDbl(row("PlannedQty")) + ReadDouble(source, "PlannedQty")
    row("ArrivedQty") = CDbl(row("ArrivedQty")) + ReadDouble(source, "ArrivedQty")
    row("LastArrivalDate") = LatestDate(row("LastArrivalDate"), ReadVariant(source, "LastArrivalDate", ReadVariant(source, "ArrivalDate", Empty)))
    If ReadDouble(source, "Confidence") > CDbl(row("Confidence")) Then row("Confidence") = ReadDouble(source, "Confidence")
    AppendMatchRule row, CStr(ReadVariant(source, "MatchRule", ""))
    If IsManualReviewStatus(ReadVariant(source, "Status", "")) Or Len(CStr(ReadVariant(source, "ReviewReason", ""))) > 0 Then row("ManualReview") = True
End Sub

Private Sub FinalizeRow(ByVal row As Object)
    row("RemainingQty") = CDbl(row("PlannedQty")) - CDbl(row("ArrivedQty"))
    If CDbl(row("RemainingQty")) < 0# Then row("RemainingQty") = 0#
    If CDbl(row("PlannedQty")) > 0# Then row("DeliveryPercent") = CDbl(row("ArrivedQty")) / CDbl(row("PlannedQty"))

    If CBool(row("ManualReview")) Then
        row("Status") = "Manual Review"
    ElseIf CDbl(row("ArrivedQty")) = 0# Then
        row("Status") = "Missing"
    ElseIf CDbl(row("ArrivedQty")) >= CDbl(row("PlannedQty")) Then
        row("Status") = "Completed"
    Else
        row("Status") = "Partial"
    End If
End Sub

Private Sub AppendMatchRule(ByVal row As Object, ByVal matchRule As String)
    If Len(matchRule) = 0 Then Exit Sub
    If Len(CStr(row("MatchRule"))) = 0 Then
        row("MatchRule") = matchRule
    ElseIf InStr(1, CStr(row("MatchRule")), matchRule, vbTextCompare) = 0 Then
        row("MatchRule") = CStr(row("MatchRule")) & "; " & matchRule
    End If
End Sub

Private Function IsManualReviewStatus(ByVal statusValue As Variant) As Boolean
    IsManualReviewStatus = (CLng(Val(CStr(statusValue))) = 4 Or CLng(Val(CStr(statusValue))) = 5)
End Function

Private Function RequiresManualReview(ByVal row As Object) As Boolean
    RequiresManualReview = CBool(row("ManualReview")) Or CStr(row("Status")) = "Manual Review" Or CDbl(row("Confidence")) < 0.8
End Function

Private Function ReadDouble(ByVal source As Object, ByVal propertyName As String) As Double
    On Error Resume Next
    ReadDouble = CDbl(CallByName(source, propertyName, VbGet))
    On Error GoTo 0
End Function

Private Function ReadVariant(ByVal source As Object, ByVal propertyName As String, ByVal defaultValue As Variant) As Variant
    On Error GoTo UseDefault
    ReadVariant = CallByName(source, propertyName, VbGet)
    Exit Function
UseDefault:
    ReadVariant = defaultValue
End Function

Private Function LatestDate(ByVal currentValue As Variant, ByVal candidate As Variant) As Variant
    If IsEmpty(currentValue) Or Len(CStr(currentValue)) = 0 Then
        LatestDate = candidate
    ElseIf IsDate(candidate) And IsDate(currentValue) Then
        If CDate(candidate) > CDate(currentValue) Then LatestDate = candidate Else LatestDate = currentValue
    Else
        LatestDate = currentValue
    End If
End Function
