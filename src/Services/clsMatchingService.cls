VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMatchingService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type ArrivalBucket
    Quantity As Double
    ArrivalDate As Variant
    MatchKey As String
End Type

Private mRules As clsMatchingRules
Private mNormalizer As clsNormalizeService
Private mPlannedMapping As clsImportMapping
Private mArrivalMapping As clsImportMapping

Public Sub Initialize(Optional ByVal rules As clsMatchingRules, Optional ByVal plannedMapping As clsImportMapping, Optional ByVal arrivalMapping As clsImportMapping)
    If rules Is Nothing Then
        Set mRules = New clsMatchingRules
    Else
        Set mRules = rules
    End If

    Set mNormalizer = New clsNormalizeService
    mNormalizer.Initialize mRules

    If plannedMapping Is Nothing Then
        Set mPlannedMapping = New clsImportMapping
    Else
        Set mPlannedMapping = plannedMapping
    End If

    If arrivalMapping Is Nothing Then
        Set mArrivalMapping = New clsImportMapping
    Else
        Set mArrivalMapping = arrivalMapping
    End If
End Sub

Public Function MatchMaterials(ByVal plannedRows As Variant, ByVal arrivalRows As Variant) As Collection
    EnsureInitialized

    Dim indexes As Object
    Set indexes = BuildArrivalIndexes(arrivalRows)

    Dim results As New Collection
    Dim rowIndex As Long
    For rowIndex = LBound(plannedRows, 1) To UBound(plannedRows, 1)
        results.Add MatchPlannedRow(plannedRows, rowIndex, indexes)
    Next rowIndex

    Set MatchMaterials = results
End Function

Private Function MatchPlannedRow(ByVal plannedRows As Variant, ByVal rowIndex As Long, ByVal indexes As Object) As clsMatchResult
    Dim result As New clsMatchResult
    Dim materialName As String
    Dim plannedQty As Double

    materialName = CStr(plannedRows(rowIndex, mPlannedMapping.MaterialNameColumn))
    plannedQty = CDbl(Val(CStr(plannedRows(rowIndex, mPlannedMapping.QuantityColumn))))

    result.Material = materialName
    result.PlannedQty = plannedQty

    If ApplyExactMatch(result, indexes("sap"), CleanKey(plannedRows(rowIndex, mPlannedMapping.SapCodeColumn)), plannedQty, "SAP Code", 1#) Then
        Set MatchPlannedRow = result
        Exit Function
    End If

    If ApplyExactMatch(result, indexes("material"), CleanKey(plannedRows(rowIndex, mPlannedMapping.MaterialCodeColumn)), plannedQty, "Material Code", 1#) Then
        Set MatchPlannedRow = result
        Exit Function
    End If

    If ApplyExactMatch(result, indexes("name"), mNormalizer.NormalizeText(materialName), plannedQty, "Normalized Name", 0.98) Then
        Set MatchPlannedRow = result
        Exit Function
    End If

    If ApplyFuzzyMatch(result, indexes("name"), mNormalizer.NormalizeText(materialName), plannedQty) Then
        Set MatchPlannedRow = result
        Exit Function
    End If

    result.Status = MatchResultNotMatched
    result.RemainingQty = plannedQty
    result.Confidence = 0#
    Set MatchPlannedRow = result
End Function

Private Function ApplyExactMatch(ByVal result As clsMatchResult, ByVal index As Object, ByVal key As String, ByVal plannedQty As Double, ByVal ruleName As String, ByVal confidence As Double) As Boolean
    If Len(key) = 0 Or Not index.Exists(key) Then Exit Function

    Dim bucket As Object
    Set bucket = index(key)
    FillResult result, bucket, plannedQty, ruleName, confidence
    ApplyExactMatch = True
End Function

Private Function ApplyFuzzyMatch(ByVal result As clsMatchResult, ByVal nameIndex As Object, ByVal plannedName As String, ByVal plannedQty As Double) As Boolean
    Dim bestKey As Variant
    Dim bestScore As Double
    Dim score As Double

    For Each bestKey In nameIndex.Keys
        score = Similarity(plannedName, CStr(bestKey))
        If score > bestScore Then
            bestScore = score
            result.MatchKey = CStr(bestKey)
        End If
    Next bestKey

    If bestScore < mRules.FuzzyThreshold Then Exit Function

    FillResult result, nameIndex(result.MatchKey), plannedQty, "Fuzzy Name", bestScore
    ApplyFuzzyMatch = True
End Function

Private Sub FillResult(ByVal result As clsMatchResult, ByVal bucket As Object, ByVal plannedQty As Double, ByVal ruleName As String, ByVal confidence As Double)
    result.ArrivedQty = CDbl(bucket("quantity"))
    result.ArrivalDate = bucket("arrivalDate")
    result.RemainingQty = plannedQty - result.ArrivedQty
    If result.RemainingQty <= 0 Then
        result.RemainingQty = 0
        result.Status = MatchResultMatched
    Else
        result.Status = MatchResultPartiallyMatched
    End If
    result.Confidence = confidence
    result.MatchKey = CStr(bucket("key"))
    result.MatchRule = ruleName
End Sub

Private Function BuildArrivalIndexes(ByVal arrivalRows As Variant) As Object
    Dim indexes As Object
    Set indexes = CreateObject("Scripting.Dictionary")
    Set indexes("sap") = CreateObject("Scripting.Dictionary")
    Set indexes("material") = CreateObject("Scripting.Dictionary")
    Set indexes("name") = CreateObject("Scripting.Dictionary")

    Dim rowIndex As Long
    For rowIndex = LBound(arrivalRows, 1) To UBound(arrivalRows, 1)
        AddArrival indexes("sap"), CleanKey(arrivalRows(rowIndex, mArrivalMapping.SapCodeColumn)), arrivalRows, rowIndex
        AddArrival indexes("material"), CleanKey(arrivalRows(rowIndex, mArrivalMapping.MaterialCodeColumn)), arrivalRows, rowIndex
        AddArrival indexes("name"), mNormalizer.NormalizeText(arrivalRows(rowIndex, mArrivalMapping.MaterialNameColumn)), arrivalRows, rowIndex
    Next rowIndex

    Set BuildArrivalIndexes = indexes
End Function

Private Sub AddArrival(ByVal index As Object, ByVal key As String, ByVal arrivalRows As Variant, ByVal rowIndex As Long)
    If Len(key) = 0 Then Exit Sub

    Dim bucket As Object
    If index.Exists(key) Then
        Set bucket = index(key)
    Else
        Set bucket = CreateObject("Scripting.Dictionary")
        bucket("quantity") = 0#
        bucket("arrivalDate") = Empty
        bucket("key") = key
        index.Add key, bucket
    End If

    bucket("quantity") = CDbl(bucket("quantity")) + CDbl(Val(CStr(arrivalRows(rowIndex, mArrivalMapping.QuantityColumn))))
    bucket("arrivalDate") = LatestDate(bucket("arrivalDate"), arrivalRows(rowIndex, mArrivalMapping.ArrivalDateColumn))
End Sub

Private Function LatestDate(ByVal currentValue As Variant, ByVal candidate As Variant) As Variant
    If IsEmpty(currentValue) Or Len(CStr(currentValue)) = 0 Then
        LatestDate = candidate
    ElseIf IsDate(candidate) And IsDate(currentValue) Then
        If CDate(candidate) > CDate(currentValue) Then LatestDate = candidate Else LatestDate = currentValue
    Else
        LatestDate = currentValue
    End If
End Function

Private Function CleanKey(ByVal value As Variant) As String
    CleanKey = UCase$(Trim$(CStr(value)))
End Function

Public Function Similarity(ByVal leftText As String, ByVal rightText As String) As Double
    Dim distance As Long
    Dim maxLength As Long
    maxLength = Len(leftText)
    If Len(rightText) > maxLength Then maxLength = Len(rightText)
    If maxLength = 0 Then
        Similarity = 1#
        Exit Function
    End If

    distance = LevenshteinDistance(leftText, rightText)
    Similarity = 1# - (distance / maxLength)
End Function

Private Function LevenshteinDistance(ByVal source As String, ByVal target As String) As Long
    Dim sourceLength As Long: sourceLength = Len(source)
    Dim targetLength As Long: targetLength = Len(target)
    Dim distances() As Long
    ReDim distances(0 To sourceLength, 0 To targetLength)

    Dim sourceIndex As Long
    Dim targetIndex As Long
    For sourceIndex = 0 To sourceLength
        distances(sourceIndex, 0) = sourceIndex
    Next sourceIndex
    For targetIndex = 0 To targetLength
        distances(0, targetIndex) = targetIndex
    Next targetIndex

    Dim cost As Long
    For sourceIndex = 1 To sourceLength
        For targetIndex = 1 To targetLength
            If Mid$(source, sourceIndex, 1) = Mid$(target, targetIndex, 1) Then cost = 0 Else cost = 1
            distances(sourceIndex, targetIndex) = Min3(distances(sourceIndex - 1, targetIndex) + 1, distances(sourceIndex, targetIndex - 1) + 1, distances(sourceIndex - 1, targetIndex - 1) + cost)
        Next targetIndex
    Next sourceIndex

    LevenshteinDistance = distances(sourceLength, targetLength)
End Function

Private Function Min3(ByVal firstValue As Long, ByVal secondValue As Long, ByVal thirdValue As Long) As Long
    Min3 = firstValue
    If secondValue < Min3 Then Min3 = secondValue
    If thirdValue < Min3 Then Min3 = thirdValue
End Function

Private Sub EnsureInitialized()
    If mRules Is Nothing Then Initialize
End Sub
