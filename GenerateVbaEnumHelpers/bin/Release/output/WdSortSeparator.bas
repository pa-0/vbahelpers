Attribute VB_Name = "wWdSortSeparator"
Function WdSortSeparatorFromString(value As String) As WdSortSeparator
    If IsNumeric(value) Then
        WdSortSeparatorFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdSortSeparateByTabs": WdSortSeparatorFromString = wdSortSeparateByTabs
        Case "wdSortSeparateByCommas": WdSortSeparatorFromString = wdSortSeparateByCommas
        Case "wdSortSeparateByDefaultTableSeparator": WdSortSeparatorFromString = wdSortSeparateByDefaultTableSeparator
    End Select
End Function

Function WdSortSeparatorToString(value As WdSortSeparator) As String
    Select Case value
        Case wdSortSeparateByTabs: WdSortSeparatorToString = "wdSortSeparateByTabs"
        Case wdSortSeparateByCommas: WdSortSeparatorToString = "wdSortSeparateByCommas"
        Case wdSortSeparateByDefaultTableSeparator: WdSortSeparatorToString = "wdSortSeparateByDefaultTableSeparator"
    End Select
End Function
