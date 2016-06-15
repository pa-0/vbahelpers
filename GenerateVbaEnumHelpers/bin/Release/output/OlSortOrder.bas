Attribute VB_Name = "wOlSortOrder"
Function OlSortOrderFromString(value As String) As OlSortOrder
    If IsNumeric(value) Then
        OlSortOrderFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "olSortNone": OlSortOrderFromString = olSortNone
        Case "olAscending": OlSortOrderFromString = olAscending
        Case "olDescending": OlSortOrderFromString = olDescending
    End Select
End Function

Function OlSortOrderToString(value As OlSortOrder) As String
    Select Case value
        Case olSortNone: OlSortOrderToString = "olSortNone"
        Case olAscending: OlSortOrderToString = "olAscending"
        Case olDescending: OlSortOrderToString = "olDescending"
    End Select
End Function
