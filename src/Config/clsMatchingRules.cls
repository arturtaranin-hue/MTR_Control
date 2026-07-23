VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMatchingRules"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mFuzzyThreshold As Double
Private mUnitDictionary As Object
Private mAbbreviationDictionary As Object

Private Sub Class_Initialize()
    mFuzzyThreshold = 0.95
    Set mUnitDictionary = CreateObject("Scripting.Dictionary")
    Set mAbbreviationDictionary = CreateObject("Scripting.Dictionary")
    LoadDefaultUnits
    LoadDefaultAbbreviations
End Sub

Public Property Get FuzzyThreshold() As Double
    FuzzyThreshold = mFuzzyThreshold
End Property

Public Property Let FuzzyThreshold(ByVal value As Double)
    mFuzzyThreshold = value
End Property

Public Property Get UnitDictionary() As Object
    Set UnitDictionary = mUnitDictionary
End Property

Public Property Get AbbreviationDictionary() As Object
    Set AbbreviationDictionary = mAbbreviationDictionary
End Property

Public Sub AddUnitAlias(ByVal aliasValue As String, ByVal normalizedUnit As String)
    mUnitDictionary(UCase$(Trim$(aliasValue))) = UCase$(Trim$(normalizedUnit))
End Sub

Public Sub AddAbbreviation(ByVal abbreviation As String, ByVal expandedValue As String)
    mAbbreviationDictionary(UCase$(Trim$(abbreviation))) = UCase$(Trim$(expandedValue))
End Sub

Private Sub LoadDefaultUnits()
    AddUnitAlias "KG", "KG"
    AddUnitAlias "KGS", "KG"
    AddUnitAlias "КГ", "KG"
    AddUnitAlias "T", "T"
    AddUnitAlias "TN", "T"
    AddUnitAlias "ТОНН", "T"
    AddUnitAlias "TON", "T"
    AddUnitAlias "M", "M"
    AddUnitAlias "М", "M"
    AddUnitAlias "M2", "M2"
    AddUnitAlias "М2", "M2"
    AddUnitAlias "M3", "M3"
    AddUnitAlias "М3", "M3"
    AddUnitAlias "PCS", "PCS"
    AddUnitAlias "PC", "PCS"
    AddUnitAlias "ШТ", "PCS"
End Sub

Private Sub LoadDefaultAbbreviations()
    AddAbbreviation "ТР", "ТРУБА"
    AddAbbreviation "КР", "КРАН"
    AddAbbreviation "ЗДВ", "ЗАДВИЖКА"
    AddAbbreviation "ФЛ", "ФЛАНЕЦ"
    AddAbbreviation "ЭЛ", "ЭЛЕМЕНТ"
End Sub
