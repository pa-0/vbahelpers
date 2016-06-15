Attribute VB_Name = "wWdInsertedTextMark"
Function WdInsertedTextMarkFromString(value As String) As WdInsertedTextMark
    If IsNumeric(value) Then
        WdInsertedTextMarkFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdInsertedTextMarkNone": WdInsertedTextMarkFromString = wdInsertedTextMarkNone
        Case "wdInsertedTextMarkBold": WdInsertedTextMarkFromString = wdInsertedTextMarkBold
        Case "wdInsertedTextMarkItalic": WdInsertedTextMarkFromString = wdInsertedTextMarkItalic
        Case "wdInsertedTextMarkUnderline": WdInsertedTextMarkFromString = wdInsertedTextMarkUnderline
        Case "wdInsertedTextMarkDoubleUnderline": WdInsertedTextMarkFromString = wdInsertedTextMarkDoubleUnderline
        Case "wdInsertedTextMarkColorOnly": WdInsertedTextMarkFromString = wdInsertedTextMarkColorOnly
        Case "wdInsertedTextMarkStrikeThrough": WdInsertedTextMarkFromString = wdInsertedTextMarkStrikeThrough
        Case "wdInsertedTextMarkDoubleStrikeThrough": WdInsertedTextMarkFromString = wdInsertedTextMarkDoubleStrikeThrough
    End Select
End Function

Function WdInsertedTextMarkToString(value As WdInsertedTextMark) As String
    Select Case value
        Case wdInsertedTextMarkNone: WdInsertedTextMarkToString = "wdInsertedTextMarkNone"
        Case wdInsertedTextMarkBold: WdInsertedTextMarkToString = "wdInsertedTextMarkBold"
        Case wdInsertedTextMarkItalic: WdInsertedTextMarkToString = "wdInsertedTextMarkItalic"
        Case wdInsertedTextMarkUnderline: WdInsertedTextMarkToString = "wdInsertedTextMarkUnderline"
        Case wdInsertedTextMarkDoubleUnderline: WdInsertedTextMarkToString = "wdInsertedTextMarkDoubleUnderline"
        Case wdInsertedTextMarkColorOnly: WdInsertedTextMarkToString = "wdInsertedTextMarkColorOnly"
        Case wdInsertedTextMarkStrikeThrough: WdInsertedTextMarkToString = "wdInsertedTextMarkStrikeThrough"
        Case wdInsertedTextMarkDoubleStrikeThrough: WdInsertedTextMarkToString = "wdInsertedTextMarkDoubleStrikeThrough"
    End Select
End Function
