Attribute VB_Name = "wWdMoveToTextMark"
Function WdMoveToTextMarkFromString(value As String) As WdMoveToTextMark
    If IsNumeric(value) Then
        WdMoveToTextMarkFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdMoveToTextMarkNone": WdMoveToTextMarkFromString = wdMoveToTextMarkNone
        Case "wdMoveToTextMarkBold": WdMoveToTextMarkFromString = wdMoveToTextMarkBold
        Case "wdMoveToTextMarkItalic": WdMoveToTextMarkFromString = wdMoveToTextMarkItalic
        Case "wdMoveToTextMarkUnderline": WdMoveToTextMarkFromString = wdMoveToTextMarkUnderline
        Case "wdMoveToTextMarkDoubleUnderline": WdMoveToTextMarkFromString = wdMoveToTextMarkDoubleUnderline
        Case "wdMoveToTextMarkColorOnly": WdMoveToTextMarkFromString = wdMoveToTextMarkColorOnly
        Case "wdMoveToTextMarkStrikeThrough": WdMoveToTextMarkFromString = wdMoveToTextMarkStrikeThrough
        Case "wdMoveToTextMarkDoubleStrikeThrough": WdMoveToTextMarkFromString = wdMoveToTextMarkDoubleStrikeThrough
    End Select
End Function

Function WdMoveToTextMarkToString(value As WdMoveToTextMark) As String
    Select Case value
        Case wdMoveToTextMarkNone: WdMoveToTextMarkToString = "wdMoveToTextMarkNone"
        Case wdMoveToTextMarkBold: WdMoveToTextMarkToString = "wdMoveToTextMarkBold"
        Case wdMoveToTextMarkItalic: WdMoveToTextMarkToString = "wdMoveToTextMarkItalic"
        Case wdMoveToTextMarkUnderline: WdMoveToTextMarkToString = "wdMoveToTextMarkUnderline"
        Case wdMoveToTextMarkDoubleUnderline: WdMoveToTextMarkToString = "wdMoveToTextMarkDoubleUnderline"
        Case wdMoveToTextMarkColorOnly: WdMoveToTextMarkToString = "wdMoveToTextMarkColorOnly"
        Case wdMoveToTextMarkStrikeThrough: WdMoveToTextMarkToString = "wdMoveToTextMarkStrikeThrough"
        Case wdMoveToTextMarkDoubleStrikeThrough: WdMoveToTextMarkToString = "wdMoveToTextMarkDoubleStrikeThrough"
    End Select
End Function
