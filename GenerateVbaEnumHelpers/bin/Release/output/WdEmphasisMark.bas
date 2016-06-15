Attribute VB_Name = "wWdEmphasisMark"
Function WdEmphasisMarkFromString(value As String) As WdEmphasisMark
    If IsNumeric(value) Then
        WdEmphasisMarkFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdEmphasisMarkNone": WdEmphasisMarkFromString = wdEmphasisMarkNone
        Case "wdEmphasisMarkOverSolidCircle": WdEmphasisMarkFromString = wdEmphasisMarkOverSolidCircle
        Case "wdEmphasisMarkOverComma": WdEmphasisMarkFromString = wdEmphasisMarkOverComma
        Case "wdEmphasisMarkOverWhiteCircle": WdEmphasisMarkFromString = wdEmphasisMarkOverWhiteCircle
        Case "wdEmphasisMarkUnderSolidCircle": WdEmphasisMarkFromString = wdEmphasisMarkUnderSolidCircle
    End Select
End Function

Function WdEmphasisMarkToString(value As WdEmphasisMark) As String
    Select Case value
        Case wdEmphasisMarkNone: WdEmphasisMarkToString = "wdEmphasisMarkNone"
        Case wdEmphasisMarkOverSolidCircle: WdEmphasisMarkToString = "wdEmphasisMarkOverSolidCircle"
        Case wdEmphasisMarkOverComma: WdEmphasisMarkToString = "wdEmphasisMarkOverComma"
        Case wdEmphasisMarkOverWhiteCircle: WdEmphasisMarkToString = "wdEmphasisMarkOverWhiteCircle"
        Case wdEmphasisMarkUnderSolidCircle: WdEmphasisMarkToString = "wdEmphasisMarkUnderSolidCircle"
    End Select
End Function
