Attribute VB_Name = "wWdIndexSortBy"
Function WdIndexSortByFromString(value As String) As WdIndexSortBy
    If IsNumeric(value) Then
        WdIndexSortByFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdIndexSortByStroke": WdIndexSortByFromString = wdIndexSortByStroke
        Case "wdIndexSortBySyllable": WdIndexSortByFromString = wdIndexSortBySyllable
    End Select
End Function

Function WdIndexSortByToString(value As WdIndexSortBy) As String
    Select Case value
        Case wdIndexSortByStroke: WdIndexSortByToString = "wdIndexSortByStroke"
        Case wdIndexSortBySyllable: WdIndexSortByToString = "wdIndexSortBySyllable"
    End Select
End Function
