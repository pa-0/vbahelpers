Attribute VB_Name = "wXlDataSeriesType"
Function XlDataSeriesTypeFromString(value As String) As XlDataSeriesType
    If IsNumeric(value) Then
        XlDataSeriesTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlGrowth": XlDataSeriesTypeFromString = xlGrowth
        Case "xlChronological": XlDataSeriesTypeFromString = xlChronological
        Case "xlAutoFill": XlDataSeriesTypeFromString = xlAutoFill
        Case "xlDataSeriesLinear": XlDataSeriesTypeFromString = xlDataSeriesLinear
    End Select
End Function

Function XlDataSeriesTypeToString(value As XlDataSeriesType) As String
    Select Case value
        Case xlGrowth: XlDataSeriesTypeToString = "xlGrowth"
        Case xlChronological: XlDataSeriesTypeToString = "xlChronological"
        Case xlAutoFill: XlDataSeriesTypeToString = "xlAutoFill"
        Case xlDataSeriesLinear: XlDataSeriesTypeToString = "xlDataSeriesLinear"
    End Select
End Function
