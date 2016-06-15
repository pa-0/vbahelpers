Attribute VB_Name = "wOlScrollBars"
Function OlScrollBarsFromString(value As String) As OlScrollBars
    If IsNumeric(value) Then
        OlScrollBarsFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "olScrollBarsNone": OlScrollBarsFromString = olScrollBarsNone
        Case "olScrollBarsHorizontal": OlScrollBarsFromString = olScrollBarsHorizontal
        Case "olScrollBarsVertical": OlScrollBarsFromString = olScrollBarsVertical
        Case "olScrollBarsBoth": OlScrollBarsFromString = olScrollBarsBoth
    End Select
End Function

Function OlScrollBarsToString(value As OlScrollBars) As String
    Select Case value
        Case olScrollBarsNone: OlScrollBarsToString = "olScrollBarsNone"
        Case olScrollBarsHorizontal: OlScrollBarsToString = "olScrollBarsHorizontal"
        Case olScrollBarsVertical: OlScrollBarsToString = "olScrollBarsVertical"
        Case olScrollBarsBoth: OlScrollBarsToString = "olScrollBarsBoth"
    End Select
End Function
