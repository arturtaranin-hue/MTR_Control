VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsNormalizeService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mRules As clsMatchingRules

Public Sub Initialize(Optional ByVal rules As clsMatchingRules)
    If rules Is Nothing Then
        Set mRules = New clsMatchingRules
    Else
        Set mRules = rules
    End If
End Sub

Public Function NormalizeText(ByVal value As Variant) As String
    EnsureInitialized

    Dim textValue As String
    textValue = UCase$(Trim$(CStr(value)))
    textValue = RemovePunctuation(textValue)
    textValue = CollapseSpaces(textValue)
    textValue = NormalizeTokens(textValue)
    NormalizeText = CollapseSpaces(textValue)
End Function

Private Function NormalizeTokens(ByVal textValue As String) As String
    Dim tokens() As String
    tokens = Split(textValue, " ")

    Dim index As Long
    For index = LBound(tokens) To UBound(tokens)
        If mRules.UnitDictionary.Exists(tokens(index)) Then
            tokens(index) = mRules.UnitDictionary(tokens(index))
        ElseIf mRules.AbbreviationDictionary.Exists(tokens(index)) Then
            tokens(index) = mRules.AbbreviationDictionary(tokens(index))
        End If
    Next index

    NormalizeTokens = Join(tokens, " ")
End Function

Private Function RemovePunctuation(ByVal textValue As String) As String
    Dim punctuation As Variant
    punctuation = Array(".", ",", ";", ":", "!", "?", "'", """", "(", ")", "[", "]", "{", "}", "/", "\", "|", "-", "_", "+", "=", Chr$(9), Chr$(10), Chr$(13))

    Dim item As Variant
    For Each item In punctuation
        textValue = Replace(textValue, CStr(item), " ")
    Next item

    RemovePunctuation = textValue
End Function

Private Function CollapseSpaces(ByVal textValue As String) As String
    textValue = Trim$(textValue)
    Do While InStr(1, textValue, "  ", vbBinaryCompare) > 0
        textValue = Replace(textValue, "  ", " ")
    Loop
    CollapseSpaces = textValue
End Function

Private Sub EnsureInitialized()
    If mRules Is Nothing Then Initialize
End Sub
