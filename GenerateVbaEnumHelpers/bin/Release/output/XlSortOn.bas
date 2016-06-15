Attribute VB_Name = "wXlSortOn"
Function XlSortOnFromString(value As String) As XlSortOn
    If IsNumeric(value) Then
        XlSortOnFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlSortOnValues": XlSortOnFromString = xlSortOnValues
        Case "xlSortOnCellColor": XlSortOnFromString = xlSortOnCellColor
        Case "xlSortOnFontColor": XlSortOnFromString = xlSortOnFontColor
        Case "xlSortOnIcon": XlSortOnFromString = xlSortOnIcon
    End Select
End Function

Function XlSortOnToString(value As XlSortOn) As String
    Select Case value
        Case xlSortOnValues: XlSortOnToString = "xlSortOnValues"
        Case xlSortOnCellColor: XlSortOnToString = "xlSortOnCellColor"
        Case xlSortOnFontColor: XlSortOnToString = "xlSortOnFontColor"
        Case xlSortOnIcon: XlSortOnToString = "xlSortOnIcon"
    End Select
End Function
