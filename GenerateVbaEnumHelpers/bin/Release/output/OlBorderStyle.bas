Attribute VB_Name = "wOlBorderStyle"
Function OlBorderStyleFromString(value As String) As OlBorderStyle
    If IsNumeric(value) Then
        OlBorderStyleFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "olBorderStyleNone": OlBorderStyleFromString = olBorderStyleNone
        Case "olBorderStyleSingle": OlBorderStyleFromString = olBorderStyleSingle
    End Select
End Function

Function OlBorderStyleToString(value As OlBorderStyle) As String
    Select Case value
        Case olBorderStyleNone: OlBorderStyleToString = "olBorderStyleNone"
        Case olBorderStyleSingle: OlBorderStyleToString = "olBorderStyleSingle"
    End Select
End Function
