Attribute VB_Name = "wWdMoveFromTextMark"
Function WdMoveFromTextMarkFromString(value As String) As WdMoveFromTextMark
    If IsNumeric(value) Then
        WdMoveFromTextMarkFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdMoveFromTextMarkHidden": WdMoveFromTextMarkFromString = wdMoveFromTextMarkHidden
        Case "wdMoveFromTextMarkDoubleStrikeThrough": WdMoveFromTextMarkFromString = wdMoveFromTextMarkDoubleStrikeThrough
        Case "wdMoveFromTextMarkStrikeThrough": WdMoveFromTextMarkFromString = wdMoveFromTextMarkStrikeThrough
        Case "wdMoveFromTextMarkCaret": WdMoveFromTextMarkFromString = wdMoveFromTextMarkCaret
        Case "wdMoveFromTextMarkPound": WdMoveFromTextMarkFromString = wdMoveFromTextMarkPound
        Case "wdMoveFromTextMarkNone": WdMoveFromTextMarkFromString = wdMoveFromTextMarkNone
        Case "wdMoveFromTextMarkBold": WdMoveFromTextMarkFromString = wdMoveFromTextMarkBold
        Case "wdMoveFromTextMarkItalic": WdMoveFromTextMarkFromString = wdMoveFromTextMarkItalic
        Case "wdMoveFromTextMarkUnderline": WdMoveFromTextMarkFromString = wdMoveFromTextMarkUnderline
        Case "wdMoveFromTextMarkDoubleUnderline": WdMoveFromTextMarkFromString = wdMoveFromTextMarkDoubleUnderline
        Case "wdMoveFromTextMarkColorOnly": WdMoveFromTextMarkFromString = wdMoveFromTextMarkColorOnly
    End Select
End Function

Function WdMoveFromTextMarkToString(value As WdMoveFromTextMark) As String
    Select Case value
        Case wdMoveFromTextMarkHidden: WdMoveFromTextMarkToString = "wdMoveFromTextMarkHidden"
        Case wdMoveFromTextMarkDoubleStrikeThrough: WdMoveFromTextMarkToString = "wdMoveFromTextMarkDoubleStrikeThrough"
        Case wdMoveFromTextMarkStrikeThrough: WdMoveFromTextMarkToString = "wdMoveFromTextMarkStrikeThrough"
        Case wdMoveFromTextMarkCaret: WdMoveFromTextMarkToString = "wdMoveFromTextMarkCaret"
        Case wdMoveFromTextMarkPound: WdMoveFromTextMarkToString = "wdMoveFromTextMarkPound"
        Case wdMoveFromTextMarkNone: WdMoveFromTextMarkToString = "wdMoveFromTextMarkNone"
        Case wdMoveFromTextMarkBold: WdMoveFromTextMarkToString = "wdMoveFromTextMarkBold"
        Case wdMoveFromTextMarkItalic: WdMoveFromTextMarkToString = "wdMoveFromTextMarkItalic"
        Case wdMoveFromTextMarkUnderline: WdMoveFromTextMarkToString = "wdMoveFromTextMarkUnderline"
        Case wdMoveFromTextMarkDoubleUnderline: WdMoveFromTextMarkToString = "wdMoveFromTextMarkDoubleUnderline"
        Case wdMoveFromTextMarkColorOnly: WdMoveFromTextMarkToString = "wdMoveFromTextMarkColorOnly"
    End Select
End Function
