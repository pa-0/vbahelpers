Attribute VB_Name = "wMsoSortOrder"
Function MsoSortOrderFromString(value As String) As MsoSortOrder
    If IsNumeric(value) Then
        MsoSortOrderFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoSortOrderAscending": MsoSortOrderFromString = msoSortOrderAscending
        Case "msoSortOrderDescending": MsoSortOrderFromString = msoSortOrderDescending
    End Select
End Function

Function MsoSortOrderToString(value As MsoSortOrder) As String
    Select Case value
        Case msoSortOrderAscending: MsoSortOrderToString = "msoSortOrderAscending"
        Case msoSortOrderDescending: MsoSortOrderToString = "msoSortOrderDescending"
    End Select
End Function
