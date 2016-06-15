Attribute VB_Name = "wWdBookmarkSortBy"
Function WdBookmarkSortByFromString(value As String) As WdBookmarkSortBy
    If IsNumeric(value) Then
        WdBookmarkSortByFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdSortByName": WdBookmarkSortByFromString = wdSortByName
        Case "wdSortByLocation": WdBookmarkSortByFromString = wdSortByLocation
    End Select
End Function

Function WdBookmarkSortByToString(value As WdBookmarkSortBy) As String
    Select Case value
        Case wdSortByName: WdBookmarkSortByToString = "wdSortByName"
        Case wdSortByLocation: WdBookmarkSortByToString = "wdSortByLocation"
    End Select
End Function
