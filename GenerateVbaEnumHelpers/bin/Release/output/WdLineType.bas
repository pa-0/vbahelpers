Attribute VB_Name = "wWdLineType"
Function WdLineTypeFromString(value As String) As WdLineType
    If IsNumeric(value) Then
        WdLineTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdTextLine": WdLineTypeFromString = wdTextLine
        Case "wdTableRow": WdLineTypeFromString = wdTableRow
    End Select
End Function

Function WdLineTypeToString(value As WdLineType) As String
    Select Case value
        Case wdTextLine: WdLineTypeToString = "wdTextLine"
        Case wdTableRow: WdLineTypeToString = "wdTableRow"
    End Select
End Function
