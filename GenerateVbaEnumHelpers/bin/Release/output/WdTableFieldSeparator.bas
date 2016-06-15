Attribute VB_Name = "wWdTableFieldSeparator"
Function WdTableFieldSeparatorFromString(value As String) As WdTableFieldSeparator
    If IsNumeric(value) Then
        WdTableFieldSeparatorFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdSeparateByParagraphs": WdTableFieldSeparatorFromString = wdSeparateByParagraphs
        Case "wdSeparateByTabs": WdTableFieldSeparatorFromString = wdSeparateByTabs
        Case "wdSeparateByCommas": WdTableFieldSeparatorFromString = wdSeparateByCommas
        Case "wdSeparateByDefaultListSeparator": WdTableFieldSeparatorFromString = wdSeparateByDefaultListSeparator
    End Select
End Function

Function WdTableFieldSeparatorToString(value As WdTableFieldSeparator) As String
    Select Case value
        Case wdSeparateByParagraphs: WdTableFieldSeparatorToString = "wdSeparateByParagraphs"
        Case wdSeparateByTabs: WdTableFieldSeparatorToString = "wdSeparateByTabs"
        Case wdSeparateByCommas: WdTableFieldSeparatorToString = "wdSeparateByCommas"
        Case wdSeparateByDefaultListSeparator: WdTableFieldSeparatorToString = "wdSeparateByDefaultListSeparator"
    End Select
End Function
