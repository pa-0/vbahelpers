Attribute VB_Name = "wXlSummaryColumn"
Function XlSummaryColumnFromString(value As String) As XlSummaryColumn
    If IsNumeric(value) Then
        XlSummaryColumnFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlSummaryOnRight": XlSummaryColumnFromString = xlSummaryOnRight
        Case "xlSummaryOnLeft": XlSummaryColumnFromString = xlSummaryOnLeft
    End Select
End Function

Function XlSummaryColumnToString(value As XlSummaryColumn) As String
    Select Case value
        Case xlSummaryOnRight: XlSummaryColumnToString = "xlSummaryOnRight"
        Case xlSummaryOnLeft: XlSummaryColumnToString = "xlSummaryOnLeft"
    End Select
End Function
