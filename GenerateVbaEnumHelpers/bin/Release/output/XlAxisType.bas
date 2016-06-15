Attribute VB_Name = "wXlAxisType"
Function XlAxisTypeFromString(value As String) As XlAxisType
    If IsNumeric(value) Then
        XlAxisTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlCategory": XlAxisTypeFromString = xlCategory
        Case "xlValue": XlAxisTypeFromString = xlValue
        Case "xlSeriesAxis": XlAxisTypeFromString = xlSeriesAxis
    End Select
End Function

Function XlAxisTypeToString(value As XlAxisType) As String
    Select Case value
        Case xlCategory: XlAxisTypeToString = "xlCategory"
        Case xlValue: XlAxisTypeToString = "xlValue"
        Case xlSeriesAxis: XlAxisTypeToString = "xlSeriesAxis"
    End Select
End Function
