Attribute VB_Name = "wXlSortType"
Function XlSortTypeFromString(value As String) As XlSortType
    If IsNumeric(value) Then
        XlSortTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlSortValues": XlSortTypeFromString = xlSortValues
        Case "xlSortLabels": XlSortTypeFromString = xlSortLabels
    End Select
End Function

Function XlSortTypeToString(value As XlSortType) As String
    Select Case value
        Case xlSortValues: XlSortTypeToString = "xlSortValues"
        Case xlSortLabels: XlSortTypeToString = "xlSortLabels"
    End Select
End Function
