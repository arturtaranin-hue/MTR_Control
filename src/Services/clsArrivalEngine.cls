VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsArrivalEngine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mPlannedMapping As clsImportMapping
Private mArrivalMapping As clsImportMapping
Private mRules As clsMatchingRules
Private mNormalizer As clsNormalizeService

Public Sub Initialize(Optional ByVal rules As clsMatchingRules, Optional ByVal plannedMapping As clsImportMapping, Optional ByVal arrivalMapping As clsImportMapping)
    If rules Is Nothing Then Set mRules = New clsMatchingRules Else Set mRules = rules
    If plannedMapping Is Nothing Then Set mPlannedMapping = New clsImportMapping Else Set mPlannedMapping = plannedMapping
    If arrivalMapping Is Nothing Then Set mArrivalMapping = New clsImportMapping Else Set mArrivalMapping = arrivalMapping

    Set mNormalizer = New clsNormalizeService
    mNormalizer.Initialize mRules
End Sub

Public Function CalculateArrivals(ByVal plannedRows As Variant, ByVal movementRows As Variant) As Collection
    EnsureInitialized

    Dim indexes As Object
    Set indexes = BuildMovementIndexes(movementRows)

    Dim consumedKeys As Object
    Set consumedKeys = CreateObject("Scripting.Dictionary")

    Dim results As New Collection
    Dim rowIndex As Long
    For rowIndex = LBound(plannedRows, 1) To UBound(plannedRows, 1)
        results.Add CalculatePlannedArrival(plannedRows, rowIndex, indexes, consumedKeys)
    Next rowIndex

    AddUnexpectedArrivals results, indexes, consumedKeys
    Set CalculateArrivals = results
End Function

Private Function CalculatePlannedArrival(ByVal plannedRows As Variant, ByVal rowIndex As Long, ByVal indexes As Object, ByVal consumedKeys As Object) As clsArrivalResult
    Dim result As New clsArrivalResult
    result.SapCode = CleanKey(plannedRows(rowIndex, mPlannedMapping.SapCodeColumn))
    result.MaterialCode = CleanKey(plannedRows(rowIndex, mPlannedMapping.MaterialCodeColumn))
    result.Material = CStr(plannedRows(rowIndex, mPlannedMapping.MaterialNameColumn))
    result.PlannedQty = ToNumber(plannedRows(rowIndex, mPlannedMapping.QuantityColumn))
    result.RequiredDate = plannedRows(rowIndex, mPlannedMapping.ArrivalDateColumn)

    Dim bucket As Object
    Set bucket = ResolveBucket(indexes, result.SapCode, result.MaterialCode, mNormalizer.NormalizeText(result.Material))
    If Not bucket Is Nothing Then
        consumedKeys(CStr(bucket("uniqueKey"))) = True
        ApplyBucket result, bucket
    End If

    FinalizeStatus result
    Set CalculatePlannedArrival = result
End Function

Private Function ResolveBucket(ByVal indexes As Object, ByVal sapKey As String, ByVal materialKey As String, ByVal nameKey As String) As Object
    If Len(sapKey) > 0 And indexes("sap").Exists(sapKey) Then Set ResolveBucket = indexes("sap")(sapKey): Exit Function
    If Len(materialKey) > 0 And indexes("material").Exists(materialKey) Then Set ResolveBucket = indexes("material")(materialKey): Exit Function
    If Len(nameKey) > 0 And indexes("name").Exists(nameKey) Then Set ResolveBucket = indexes("name")(nameKey): Exit Function
End Function

Private Sub ApplyBucket(ByVal result As clsArrivalResult, ByVal bucket As Object)
    result.ArrivedQty = CDbl(bucket("quantity"))
    result.FirstArrivalDate = bucket("firstDate")
    result.LastArrivalDate = bucket("lastDate")
    result.DeliveryCount = CLng(bucket("deliveryCount"))
    result.MatchKey = CStr(bucket("uniqueKey"))
End Sub

Private Sub FinalizeStatus(ByVal result As clsArrivalResult)
    If result.PlannedQty <= 0 Then
        result.CompletionPercent = 0#
        result.Status = MANUAL_REVIEW
        result.ReviewReason = "Missing planned material or non-positive planned quantity"
        Exit Sub
    End If

    result.RemainingQty = result.PlannedQty - result.ArrivedQty
    result.CompletionPercent = result.ArrivedQty / result.PlannedQty

    If result.ArrivedQty = 0 Then
        result.Status = NOT_ARRIVED
    ElseIf result.ArrivedQty > result.PlannedQty Then
        result.RemainingQty = 0#
        result.Status = OVER_DELIVERED
    ElseIf result.ArrivedQty = result.PlannedQty Then
        result.Status = ARRIVED
    Else
        result.Status = PARTIALLY_ARRIVED
    End If
End Sub

Private Function BuildMovementIndexes(ByVal movementRows As Variant) As Object
    Dim indexes As Object
    Set indexes = CreateObject("Scripting.Dictionary")
    Set indexes("sap") = CreateObject("Scripting.Dictionary")
    Set indexes("material") = CreateObject("Scripting.Dictionary")
    Set indexes("name") = CreateObject("Scripting.Dictionary")
    Set indexes("all") = CreateObject("Scripting.Dictionary")

    Dim seenRecords As Object
    Set seenRecords = CreateObject("Scripting.Dictionary")

    Dim rowIndex As Long
    For rowIndex = LBound(movementRows, 1) To UBound(movementRows, 1)
        AddMovement indexes, seenRecords, movementRows, rowIndex
    Next rowIndex

    Set BuildMovementIndexes = indexes
End Function

Private Sub AddMovement(ByVal indexes As Object, ByVal seenRecords As Object, ByVal movementRows As Variant, ByVal rowIndex As Long)
    Dim sapKey As String: sapKey = CleanKey(movementRows(rowIndex, mArrivalMapping.SapCodeColumn))
    Dim materialKey As String: materialKey = CleanKey(movementRows(rowIndex, mArrivalMapping.MaterialCodeColumn))
    Dim nameKey As String: nameKey = mNormalizer.NormalizeText(movementRows(rowIndex, mArrivalMapping.MaterialNameColumn))
    Dim qty As Double: qty = ToNumber(movementRows(rowIndex, mArrivalMapping.QuantityColumn))
    Dim movementDate As Variant: movementDate = movementRows(rowIndex, mArrivalMapping.ArrivalDateColumn)
    Dim recordKey As String: recordKey = sapKey & "|" & materialKey & "|" & nameKey & "|" & CStr(qty) & "|" & CStr(movementDate)
    If seenRecords.Exists(recordKey) Then Exit Sub
    seenRecords.Add recordKey, True

    Dim uniqueKey As String
    If Len(sapKey) > 0 Then
        uniqueKey = "SAP:" & sapKey
    ElseIf Len(materialKey) > 0 Then
        uniqueKey = "MAT:" & materialKey
    Else
        uniqueKey = "NAME:" & nameKey
    End If

    Dim bucket As Object
    If indexes("all").Exists(uniqueKey) Then
        Set bucket = indexes("all")(uniqueKey)
    Else
        Set bucket = CreateObject("Scripting.Dictionary")
        bucket("quantity") = 0#
        bucket("firstDate") = Empty
        bucket("lastDate") = Empty
        bucket("deliveryCount") = 0&
        bucket("uniqueKey") = uniqueKey
        bucket("material") = CStr(movementRows(rowIndex, mArrivalMapping.MaterialNameColumn))
        indexes("all").Add uniqueKey, bucket
    End If

    bucket("quantity") = CDbl(bucket("quantity")) + qty
    bucket("deliveryCount") = CLng(bucket("deliveryCount")) + 1
    bucket("firstDate") = EarliestDate(bucket("firstDate"), movementDate)
    bucket("lastDate") = LatestDate(bucket("lastDate"), movementDate)
    If Len(sapKey) > 0 Then Set indexes("sap")(sapKey) = bucket
    If Len(materialKey) > 0 Then Set indexes("material")(materialKey) = bucket
    If Len(nameKey) > 0 Then Set indexes("name")(nameKey) = bucket
End Sub

Private Sub AddUnexpectedArrivals(ByVal results As Collection, ByVal indexes As Object, ByVal consumedKeys As Object)
    Dim key As Variant
    For Each key In indexes("all").Keys
        If Not consumedKeys.Exists(CStr(key)) Then
            Dim result As New clsArrivalResult
            ApplyBucket result, indexes("all")(key)
            result.Material = CStr(indexes("all")(key)("material"))
            result.Status = MANUAL_REVIEW
            result.UnexpectedArrival = True
            result.ReviewReason = "Unexpected arrival without matching planned material"
            results.Add result
        End If
    Next key
End Sub

Private Function ToNumber(ByVal value As Variant) As Double
    ToNumber = CDbl(Val(CStr(value)))
End Function

Private Function CleanKey(ByVal value As Variant) As String
    CleanKey = UCase$(Trim$(CStr(value)))
End Function

Private Function EarliestDate(ByVal currentValue As Variant, ByVal candidate As Variant) As Variant
    If IsEmpty(currentValue) Or Len(CStr(currentValue)) = 0 Then
        EarliestDate = candidate
    ElseIf IsDate(candidate) And IsDate(currentValue) Then
        If CDate(candidate) < CDate(currentValue) Then EarliestDate = candidate Else EarliestDate = currentValue
    Else
        EarliestDate = currentValue
    End If
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

Private Sub EnsureInitialized()
    If mPlannedMapping Is Nothing Then Initialize
End Sub
