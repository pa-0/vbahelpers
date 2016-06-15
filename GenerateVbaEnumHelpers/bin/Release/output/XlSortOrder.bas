Attribute VB_Name = "wXlSortOrder"
Function XlSortOrderFromString(value As String) As XlSortOrder
    If IsNumeric(value) Then
        XlSortOrderFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlAscending": XlSortOrderFromString = xlAscending
        Case "xlDescending": XlSortOrderFromString = xlDescending
    End Select
End Function

Function XlSortOrderToString(value As XlSortOrder) As String
    Select Case value
        Case xlAscending: XlSortOrderToString = "xlAscending"
        Case xlDescending: XlSortOrderToString = "xlDescending"
    End Select
End Function
