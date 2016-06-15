Attribute VB_Name = "wXlChartSplitType"
Function XlChartSplitTypeFromString(value As String) As XlChartSplitType
    If IsNumeric(value) Then
        XlChartSplitTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlSplitByPosition": XlChartSplitTypeFromString = xlSplitByPosition
        Case "xlSplitByValue": XlChartSplitTypeFromString = xlSplitByValue
        Case "xlSplitByPercentValue": XlChartSplitTypeFromString = xlSplitByPercentValue
        Case "xlSplitByCustomSplit": XlChartSplitTypeFromString = xlSplitByCustomSplit
    End Select
End Function

Function XlChartSplitTypeToString(value As XlChartSplitType) As String
    Select Case value
        Case xlSplitByPosition: XlChartSplitTypeToString = "xlSplitByPosition"
        Case xlSplitByValue: XlChartSplitTypeToString = "xlSplitByValue"
        Case xlSplitByPercentValue: XlChartSplitTypeToString = "xlSplitByPercentValue"
        Case xlSplitByCustomSplit: XlChartSplitTypeToString = "xlSplitByCustomSplit"
    End Select
End Function
