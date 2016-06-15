Attribute VB_Name = "wPpArrangeStyle"
Function PpArrangeStyleFromString(value As String) As PpArrangeStyle
    If IsNumeric(value) Then
        PpArrangeStyleFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "ppArrangeTiled": PpArrangeStyleFromString = ppArrangeTiled
        Case "ppArrangeCascade": PpArrangeStyleFromString = ppArrangeCascade
    End Select
End Function

Function PpArrangeStyleToString(value As PpArrangeStyle) As String
    Select Case value
        Case ppArrangeTiled: PpArrangeStyleToString = "ppArrangeTiled"
        Case ppArrangeCascade: PpArrangeStyleToString = "ppArrangeCascade"
    End Select
End Function
