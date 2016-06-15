Attribute VB_Name = "wXlAutoFillType"
Function XlAutoFillTypeFromString(value As String) As XlAutoFillType
    If IsNumeric(value) Then
        XlAutoFillTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlFillDefault": XlAutoFillTypeFromString = xlFillDefault
        Case "xlFillCopy": XlAutoFillTypeFromString = xlFillCopy
        Case "xlFillSeries": XlAutoFillTypeFromString = xlFillSeries
        Case "xlFillFormats": XlAutoFillTypeFromString = xlFillFormats
        Case "xlFillValues": XlAutoFillTypeFromString = xlFillValues
        Case "xlFillDays": XlAutoFillTypeFromString = xlFillDays
        Case "xlFillWeekdays": XlAutoFillTypeFromString = xlFillWeekdays
        Case "xlFillMonths": XlAutoFillTypeFromString = xlFillMonths
        Case "xlFillYears": XlAutoFillTypeFromString = xlFillYears
        Case "xlLinearTrend": XlAutoFillTypeFromString = xlLinearTrend
        Case "xlGrowthTrend": XlAutoFillTypeFromString = xlGrowthTrend
    End Select
End Function

Function XlAutoFillTypeToString(value As XlAutoFillType) As String
    Select Case value
        Case xlFillDefault: XlAutoFillTypeToString = "xlFillDefault"
        Case xlFillCopy: XlAutoFillTypeToString = "xlFillCopy"
        Case xlFillSeries: XlAutoFillTypeToString = "xlFillSeries"
        Case xlFillFormats: XlAutoFillTypeToString = "xlFillFormats"
        Case xlFillValues: XlAutoFillTypeToString = "xlFillValues"
        Case xlFillDays: XlAutoFillTypeToString = "xlFillDays"
        Case xlFillWeekdays: XlAutoFillTypeToString = "xlFillWeekdays"
        Case xlFillMonths: XlAutoFillTypeToString = "xlFillMonths"
        Case xlFillYears: XlAutoFillTypeToString = "xlFillYears"
        Case xlLinearTrend: XlAutoFillTypeToString = "xlLinearTrend"
        Case xlGrowthTrend: XlAutoFillTypeToString = "xlGrowthTrend"
    End Select
End Function
