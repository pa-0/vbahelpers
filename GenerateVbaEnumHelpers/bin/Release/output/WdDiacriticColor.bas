Attribute VB_Name = "wWdDiacriticColor"
Function WdDiacriticColorFromString(value As String) As WdDiacriticColor
    If IsNumeric(value) Then
        WdDiacriticColorFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdDiacriticColorBidi": WdDiacriticColorFromString = wdDiacriticColorBidi
        Case "wdDiacriticColorLatin": WdDiacriticColorFromString = wdDiacriticColorLatin
    End Select
End Function

Function WdDiacriticColorToString(value As WdDiacriticColor) As String
    Select Case value
        Case wdDiacriticColorBidi: WdDiacriticColorToString = "wdDiacriticColorBidi"
        Case wdDiacriticColorLatin: WdDiacriticColorToString = "wdDiacriticColorLatin"
    End Select
End Function
