Attribute VB_Name = "wWdDeletedTextMark"
Function WdDeletedTextMarkFromString(value As String) As WdDeletedTextMark
    If IsNumeric(value) Then
        WdDeletedTextMarkFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdDeletedTextMarkHidden": WdDeletedTextMarkFromString = wdDeletedTextMarkHidden
        Case "wdDeletedTextMarkStrikeThrough": WdDeletedTextMarkFromString = wdDeletedTextMarkStrikeThrough
        Case "wdDeletedTextMarkCaret": WdDeletedTextMarkFromString = wdDeletedTextMarkCaret
        Case "wdDeletedTextMarkPound": WdDeletedTextMarkFromString = wdDeletedTextMarkPound
        Case "wdDeletedTextMarkNone": WdDeletedTextMarkFromString = wdDeletedTextMarkNone
        Case "wdDeletedTextMarkBold": WdDeletedTextMarkFromString = wdDeletedTextMarkBold
        Case "wdDeletedTextMarkItalic": WdDeletedTextMarkFromString = wdDeletedTextMarkItalic
        Case "wdDeletedTextMarkUnderline": WdDeletedTextMarkFromString = wdDeletedTextMarkUnderline
        Case "wdDeletedTextMarkDoubleUnderline": WdDeletedTextMarkFromString = wdDeletedTextMarkDoubleUnderline
        Case "wdDeletedTextMarkColorOnly": WdDeletedTextMarkFromString = wdDeletedTextMarkColorOnly
        Case "wdDeletedTextMarkDoubleStrikeThrough": WdDeletedTextMarkFromString = wdDeletedTextMarkDoubleStrikeThrough
    End Select
End Function

Function WdDeletedTextMarkToString(value As WdDeletedTextMark) As String
    Select Case value
        Case wdDeletedTextMarkHidden: WdDeletedTextMarkToString = "wdDeletedTextMarkHidden"
        Case wdDeletedTextMarkStrikeThrough: WdDeletedTextMarkToString = "wdDeletedTextMarkStrikeThrough"
        Case wdDeletedTextMarkCaret: WdDeletedTextMarkToString = "wdDeletedTextMarkCaret"
        Case wdDeletedTextMarkPound: WdDeletedTextMarkToString = "wdDeletedTextMarkPound"
        Case wdDeletedTextMarkNone: WdDeletedTextMarkToString = "wdDeletedTextMarkNone"
        Case wdDeletedTextMarkBold: WdDeletedTextMarkToString = "wdDeletedTextMarkBold"
        Case wdDeletedTextMarkItalic: WdDeletedTextMarkToString = "wdDeletedTextMarkItalic"
        Case wdDeletedTextMarkUnderline: WdDeletedTextMarkToString = "wdDeletedTextMarkUnderline"
        Case wdDeletedTextMarkDoubleUnderline: WdDeletedTextMarkToString = "wdDeletedTextMarkDoubleUnderline"
        Case wdDeletedTextMarkColorOnly: WdDeletedTextMarkToString = "wdDeletedTextMarkColorOnly"
        Case wdDeletedTextMarkDoubleStrikeThrough: WdDeletedTextMarkToString = "wdDeletedTextMarkDoubleStrikeThrough"
    End Select
End Function
