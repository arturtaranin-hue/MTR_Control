VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsReconciliationService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mPlanMapping As clsImportMapping
Private mMovementMapping As clsImportMapping
Private mAngaraMapping As clsImportMapping
Private mRules As clsMatchingRules
Private mNormalizer As clsNormalizeService

Public Sub Initialize(Optional ByVal rules As clsMatchingRules, Optional ByVal planMapping As clsImportMapping, Optional ByVal movementMapping As clsImportMapping, Optional ByVal angaraMapping As clsImportMapping)
    If rules Is Nothing Then Set mRules = New clsMatchingRules Else Set mRules = rules
    If planMapping Is Nothing Then Set mPlanMapping = New clsImportMapping Else Set mPlanMapping = planMapping
    If movementMapping Is Nothing Then Set mMovementMapping = New clsImportMapping Else Set mMovementMapping = movementMapping
    If angaraMapping Is Nothing Then Set mAngaraMapping = New clsImportMapping Else Set mAngaraMapping = angaraMapping
    Set mNormalizer = New clsNormalizeService
    mNormalizer.Initialize mRules
End Sub

Public Function Reconcile(ByVal dvPlanRows As Variant, ByVal movementSummaryRows As Variant, ByVal gpnAngaraRows As Variant) As Collection
    EnsureInitialized

    Dim arrivals As Object
    Set arrivals = BuildArrivalIndexes(movementSummaryRows, gpnAngaraRows)

    Dim consumedKeys As Object
    Set consumedKeys = CreateObject("Scripting.Dictionary")

    Dim states As New Collection
    Dim rowIndex As Long
    For rowIndex = LBound(dvPlanRows, 1) To UBound(dvPlanRows, 1)
        states.Add ReconcilePlannedMaterial(dvPlanRows, rowIndex, arrivals, consumedKeys)
    Next rowIndex

    AddUnknownMaterials states, arrivals, consumedKeys
    Set Reconcile = states
End Function

Private Function ReconcilePlannedMaterial(ByVal planRows As Variant, ByVal rowIndex As Long, ByVal arrivals As Object, ByVal consumedKeys As Object) As clsMaterialState
    Dim state As New clsMaterialState
    state.SapCode = CleanKey(planRows(rowIndex, mPlanMapping.SapCodeColumn))
    state.MaterialCode = CleanKey(planRows(rowIndex, mPlanMapping.MaterialCodeColumn))
    state.Material = CStr(planRows(rowIndex, mPlanMapping.MaterialNameColumn))
    state.PlannedQty = ToNumber(planRows(rowIndex, mPlanMapping.QuantityColumn))
    state.SourceRow = rowIndex

    Dim bucket As Object
    Set bucket = ResolveBucket(arrivals, state.SapCode, state.MaterialCode, mNormalizer.NormalizeText(state.Material))
    If Not bucket Is Nothing Then
        consumedKeys(CStr(bucket("uniqueKey"))) = True
        ApplyArrivalBucket state, bucket
    End If

    FinalizeState state
    Set ReconcilePlannedMaterial = state
End Function

Private Function BuildArrivalIndexes(ByVal movementRows As Variant, ByVal angaraRows As Variant) As Object
    Dim indexes As Object
    Set indexes = CreateObject("Scripting.Dictionary")
    Set indexes("sap") = CreateObject("Scripting.Dictionary")
    Set indexes("material") = CreateObject("Scripting.Dictionary")
    Set indexes("name") = CreateObject("Scripting.Dictionary")
    Set indexes("all") = CreateObject("Scripting.Dictionary")
    Set indexes("records") = CreateObject("Scripting.Dictionary")

    AddArrivalRows indexes, movementRows, mMovementMapping, "Movement Summary"
    AddArrivalRows indexes, angaraRows, mAngaraMapping, "GPN Angara"

    Set BuildArrivalIndexes = indexes
End Function

Private Sub AddArrivalRows(ByVal indexes As Object, ByVal rows As Variant, ByVal mapping As clsImportMapping, ByVal sourceName As String)
    If IsEmpty(rows) Then Exit Sub

    Dim rowIndex As Long
    For rowIndex = LBound(rows, 1) To UBound(rows, 1)
        AddArrival indexes, rows, mapping, rowIndex, sourceName
    Next rowIndex
End Sub

Private Sub AddArrival(ByVal indexes As Object, ByVal rows As Variant, ByVal mapping As clsImportMapping, ByVal rowIndex As Long, ByVal sourceName As String)
    Dim sapKey As String: sapKey = CleanKey(rows(rowIndex, mapping.SapCodeColumn))
    Dim materialKey As String: materialKey = CleanKey(rows(rowIndex, mapping.MaterialCodeColumn))
    Dim materialName As String: materialName = CStr(rows(rowIndex, mapping.MaterialNameColumn))
    Dim nameKey As String: nameKey = mNormalizer.NormalizeText(materialName)
    Dim qty As Double: qty = ToNumber(rows(rowIndex, mapping.QuantityColumn))
    Dim arrivalDate As Variant: arrivalDate = rows(rowIndex, mapping.ArrivalDateColumn)
    Dim recordKey As String: recordKey = sapKey & "|" & materialKey & "|" & nameKey & "|" & CStr(qty) & "|" & CStr(arrivalDate)

    Dim duplicateArrival As Boolean
    duplicateArrival = indexes("records").Exists(recordKey)
    If Not duplicateArrival Then indexes("records").Add recordKey, sourceName

    Dim uniqueKey As String: uniqueKey = UniqueMaterialKey(sapKey, materialKey, nameKey)
    Dim bucket As Object
    If indexes("all").Exists(uniqueKey) Then
        Set bucket = indexes("all")(uniqueKey)
    Else
        Set bucket = CreateObject("Scripting.Dictionary")
        bucket("uniqueKey") = uniqueKey
        bucket("quantity") = 0#
        bucket("movementQty") = 0#
        bucket("angaraQty") = 0#
        bucket("arrivalCount") = 0&
        bucket("duplicateArrival") = False
        bucket("material") = materialName
        indexes("all").Add uniqueKey, bucket
    End If

    bucket("quantity") = CDbl(bucket("quantity")) + qty
    If sourceName = "GPN Angara" Then bucket("angaraQty") = CDbl(bucket("angaraQty")) + qty Else bucket("movementQty") = CDbl(bucket("movementQty")) + qty
    bucket("arrivalCount") = CLng(bucket("arrivalCount")) + 1
    If duplicateArrival Then bucket("duplicateArrival") = True

    If Len(sapKey) > 0 Then Set indexes("sap")(sapKey) = bucket
    If Len(materialKey) > 0 Then Set indexes("material")(materialKey) = bucket
    If Len(nameKey) > 0 Then Set indexes("name")(nameKey) = bucket
End Sub

Private Function ResolveBucket(ByVal indexes As Object, ByVal sapKey As String, ByVal materialKey As String, ByVal nameKey As String) As Object
    If Len(sapKey) > 0 And indexes("sap").Exists(sapKey) Then Set ResolveBucket = indexes("sap")(sapKey): Exit Function
    If Len(materialKey) > 0 And indexes("material").Exists(materialKey) Then Set ResolveBucket = indexes("material")(materialKey): Exit Function
    If Len(nameKey) > 0 And indexes("name").Exists(nameKey) Then Set ResolveBucket = indexes("name")(nameKey): Exit Function
End Function

Private Sub ApplyArrivalBucket(ByVal state As clsMaterialState, ByVal bucket As Object)
    state.ArrivedQty = CDbl(bucket("quantity"))
    state.MovementQty = CDbl(bucket("movementQty"))
    state.AngaraQty = CDbl(bucket("angaraQty"))
    state.ArrivalCount = CLng(bucket("arrivalCount"))
    state.DuplicateArrival = CBool(bucket("duplicateArrival"))
    state.MatchKey = CStr(bucket("uniqueKey"))
End Sub

Private Sub FinalizeState(ByVal state As clsMaterialState)
    state.RemainingQty = state.PlannedQty - state.ArrivedQty
    If state.PlannedQty > 0 Then state.DeliveryPercent = state.ArrivedQty / state.PlannedQty Else state.DeliveryPercent = 0#
    state.NegativeBalance = (state.RemainingQty < 0#)
    state.OverDelivery = (state.ArrivedQty > state.PlannedQty And state.PlannedQty > 0#)

    If state.PlannedQty <= 0# Then
        MarkManualReview state, "Missing or non-positive planned quantity"
    ElseIf state.DuplicateArrival Then
        MarkManualReview state, "Duplicate arrival detected"
    ElseIf state.OverDelivery Then
        state.Status = OVERDELIVERED
        state.RemainingQty = 0#
    ElseIf state.ArrivedQty = state.PlannedQty Then
        state.Status = DELIVERED
    ElseIf state.ArrivedQty > 0# Then
        state.Status = PARTIAL
    Else
        state.Status = NOT_FOUND
    End If
End Sub

Private Sub AddUnknownMaterials(ByVal states As Collection, ByVal arrivals As Object, ByVal consumedKeys As Object)
    Dim key As Variant
    For Each key In arrivals("all").Keys
        If Not consumedKeys.Exists(CStr(key)) Then
            Dim state As New clsMaterialState
            ApplyArrivalBucket state, arrivals("all")(key)
            state.Material = CStr(arrivals("all")(key)("material"))
            state.UnknownMaterial = True
            MarkManualReview state, "Unknown material arrival without DV Plan match"
            states.Add state
        End If
    Next key
End Sub

Private Sub MarkManualReview(ByVal state As clsMaterialState, ByVal reason As String)
    state.Status = MANUAL_REVIEW
    state.ManualReview = True
    state.ReviewReason = reason
End Sub

Private Function UniqueMaterialKey(ByVal sapKey As String, ByVal materialKey As String, ByVal nameKey As String) As String
    If Len(sapKey) > 0 Then
        UniqueMaterialKey = "SAP:" & sapKey
    ElseIf Len(materialKey) > 0 Then
        UniqueMaterialKey = "MAT:" & materialKey
    Else
        UniqueMaterialKey = "NAME:" & nameKey
    End If
End Function

Private Function ToNumber(ByVal value As Variant) As Double
    ToNumber = CDbl(Val(CStr(value)))
End Function

Private Function CleanKey(ByVal value As Variant) As String
    CleanKey = UCase$(Trim$(CStr(value)))
End Function

Private Sub EnsureInitialized()
    If mPlanMapping Is Nothing Then Initialize
End Sub
