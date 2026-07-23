Attribute VB_Name = "modUtils"
Option Explicit

Public Function NzString(ByVal value As Variant, Optional ByVal fallback As String = vbNullString) As String
    If IsError(value) Or IsNull(value) Or IsEmpty(value) Then
        NzString = fallback
    Else
        NzString = CStr(value)
    End If
End Function

Public Function IsBlankString(ByVal value As String) As Boolean
    IsBlankString = (Len(Trim$(value)) = 0)
End Function
