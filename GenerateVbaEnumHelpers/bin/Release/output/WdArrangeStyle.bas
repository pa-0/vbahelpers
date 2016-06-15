Attribute VB_Name = "wWdArrangeStyle"
Function WdArrangeStyleFromString(value As String) As WdArrangeStyle
    If IsNumeric(value) Then
        WdArrangeStyleFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdTiled": WdArrangeStyleFromString = wdTiled
        Case "wdIcons": WdArrangeStyleFromString = wdIcons
    End Select
End Function

Function WdArrangeStyleToString(value As WdArrangeStyle) As String
    Select Case value
        Case wdTiled: WdArrangeStyleToString = "wdTiled"
        Case wdIcons: WdArrangeStyleToString = "wdIcons"
    End Select
End Function
