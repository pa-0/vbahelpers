Attribute VB_Name = "wXlSlicerSort"
Function XlSlicerSortFromString(value As String) As XlSlicerSort
    If IsNumeric(value) Then
        XlSlicerSortFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlSlicerSortDataSourceOrder": XlSlicerSortFromString = xlSlicerSortDataSourceOrder
        Case "xlSlicerSortAscending": XlSlicerSortFromString = xlSlicerSortAscending
        Case "xlSlicerSortDescending": XlSlicerSortFromString = xlSlicerSortDescending
    End Select
End Function

Function XlSlicerSortToString(value As XlSlicerSort) As String
    Select Case value
        Case xlSlicerSortDataSourceOrder: XlSlicerSortToString = "xlSlicerSortDataSourceOrder"
        Case xlSlicerSortAscending: XlSlicerSortToString = "xlSlicerSortAscending"
        Case xlSlicerSortDescending: XlSlicerSortToString = "xlSlicerSortDescending"
    End Select
End Function
