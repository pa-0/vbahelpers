Attribute VB_Name = "wXlSummaryReportType"
Function XlSummaryReportTypeFromString(value As String) As XlSummaryReportType
    If IsNumeric(value) Then
        XlSummaryReportTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlStandardSummary": XlSummaryReportTypeFromString = xlStandardSummary
        Case "xlSummaryPivotTable": XlSummaryReportTypeFromString = xlSummaryPivotTable
    End Select
End Function

Function XlSummaryReportTypeToString(value As XlSummaryReportType) As String
    Select Case value
        Case xlStandardSummary: XlSummaryReportTypeToString = "xlStandardSummary"
        Case xlSummaryPivotTable: XlSummaryReportTypeToString = "xlSummaryPivotTable"
    End Select
End Function
