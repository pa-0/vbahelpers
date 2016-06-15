Attribute VB_Name = "wWdRevisedPropertiesMark"
Function WdRevisedPropertiesMarkFromString(value As String) As WdRevisedPropertiesMark
    If IsNumeric(value) Then
        WdRevisedPropertiesMarkFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdRevisedPropertiesMarkNone": WdRevisedPropertiesMarkFromString = wdRevisedPropertiesMarkNone
        Case "wdRevisedPropertiesMarkBold": WdRevisedPropertiesMarkFromString = wdRevisedPropertiesMarkBold
        Case "wdRevisedPropertiesMarkItalic": WdRevisedPropertiesMarkFromString = wdRevisedPropertiesMarkItalic
        Case "wdRevisedPropertiesMarkUnderline": WdRevisedPropertiesMarkFromString = wdRevisedPropertiesMarkUnderline
        Case "wdRevisedPropertiesMarkDoubleUnderline": WdRevisedPropertiesMarkFromString = wdRevisedPropertiesMarkDoubleUnderline
        Case "wdRevisedPropertiesMarkColorOnly": WdRevisedPropertiesMarkFromString = wdRevisedPropertiesMarkColorOnly
        Case "wdRevisedPropertiesMarkStrikeThrough": WdRevisedPropertiesMarkFromString = wdRevisedPropertiesMarkStrikeThrough
        Case "wdRevisedPropertiesMarkDoubleStrikeThrough": WdRevisedPropertiesMarkFromString = wdRevisedPropertiesMarkDoubleStrikeThrough
    End Select
End Function

Function WdRevisedPropertiesMarkToString(value As WdRevisedPropertiesMark) As String
    Select Case value
        Case wdRevisedPropertiesMarkNone: WdRevisedPropertiesMarkToString = "wdRevisedPropertiesMarkNone"
        Case wdRevisedPropertiesMarkBold: WdRevisedPropertiesMarkToString = "wdRevisedPropertiesMarkBold"
        Case wdRevisedPropertiesMarkItalic: WdRevisedPropertiesMarkToString = "wdRevisedPropertiesMarkItalic"
        Case wdRevisedPropertiesMarkUnderline: WdRevisedPropertiesMarkToString = "wdRevisedPropertiesMarkUnderline"
        Case wdRevisedPropertiesMarkDoubleUnderline: WdRevisedPropertiesMarkToString = "wdRevisedPropertiesMarkDoubleUnderline"
        Case wdRevisedPropertiesMarkColorOnly: WdRevisedPropertiesMarkToString = "wdRevisedPropertiesMarkColorOnly"
        Case wdRevisedPropertiesMarkStrikeThrough: WdRevisedPropertiesMarkToString = "wdRevisedPropertiesMarkStrikeThrough"
        Case wdRevisedPropertiesMarkDoubleStrikeThrough: WdRevisedPropertiesMarkToString = "wdRevisedPropertiesMarkDoubleStrikeThrough"
    End Select
End Function
