Attribute VB_Name = "wOlTimeStyle"
Function OlTimeStyleFromString(value As String) As OlTimeStyle
    If IsNumeric(value) Then
        OlTimeStyleFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "olTimeStyleTimeOnly": OlTimeStyleFromString = olTimeStyleTimeOnly
        Case "olTimeStyleTimeDuration": OlTimeStyleFromString = olTimeStyleTimeDuration
        Case "olTimeStyleShortDuration": OlTimeStyleFromString = olTimeStyleShortDuration
    End Select
End Function

Function OlTimeStyleToString(value As OlTimeStyle) As String
    Select Case value
        Case olTimeStyleTimeOnly: OlTimeStyleToString = "olTimeStyleTimeOnly"
        Case olTimeStyleTimeDuration: OlTimeStyleToString = "olTimeStyleTimeDuration"
        Case olTimeStyleShortDuration: OlTimeStyleToString = "olTimeStyleShortDuration"
    End Select
End Function
