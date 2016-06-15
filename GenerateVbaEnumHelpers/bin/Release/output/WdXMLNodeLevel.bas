Attribute VB_Name = "wWdXMLNodeLevel"
Function WdXMLNodeLevelFromString(value As String) As WdXMLNodeLevel
    If IsNumeric(value) Then
        WdXMLNodeLevelFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdXMLNodeLevelInline": WdXMLNodeLevelFromString = wdXMLNodeLevelInline
        Case "wdXMLNodeLevelParagraph": WdXMLNodeLevelFromString = wdXMLNodeLevelParagraph
        Case "wdXMLNodeLevelRow": WdXMLNodeLevelFromString = wdXMLNodeLevelRow
        Case "wdXMLNodeLevelCell": WdXMLNodeLevelFromString = wdXMLNodeLevelCell
    End Select
End Function

Function WdXMLNodeLevelToString(value As WdXMLNodeLevel) As String
    Select Case value
        Case wdXMLNodeLevelInline: WdXMLNodeLevelToString = "wdXMLNodeLevelInline"
        Case wdXMLNodeLevelParagraph: WdXMLNodeLevelToString = "wdXMLNodeLevelParagraph"
        Case wdXMLNodeLevelRow: WdXMLNodeLevelToString = "wdXMLNodeLevelRow"
        Case wdXMLNodeLevelCell: WdXMLNodeLevelToString = "wdXMLNodeLevelCell"
    End Select
End Function
