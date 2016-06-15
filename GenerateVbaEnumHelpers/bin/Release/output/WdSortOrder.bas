Attribute VB_Name = "wWdSortOrder"
Function WdSortOrderFromString(value As String) As WdSortOrder
    If IsNumeric(value) Then
        WdSortOrderFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdSortOrderAscending": WdSortOrderFromString = wdSortOrderAscending
        Case "wdSortOrderDescending": WdSortOrderFromString = wdSortOrderDescending
    End Select
End Function

Function WdSortOrderToString(value As WdSortOrder) As String
    Select Case value
        Case wdSortOrderAscending: WdSortOrderToString = "wdSortOrderAscending"
        Case wdSortOrderDescending: WdSortOrderToString = "wdSortOrderDescending"
    End Select
End Function
