Attribute VB_Name = "wXlSummaryRow"
Function XlSummaryRowFromString(value As String) As XlSummaryRow
    If IsNumeric(value) Then
        XlSummaryRowFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlSummaryAbove": XlSummaryRowFromString = xlSummaryAbove
        Case "xlSummaryBelow": XlSummaryRowFromString = xlSummaryBelow
    End Select
End Function

Function XlSummaryRowToString(value As XlSummaryRow) As String
    Select Case value
        Case xlSummaryAbove: XlSummaryRowToString = "xlSummaryAbove"
        Case xlSummaryBelow: XlSummaryRowToString = "xlSummaryBelow"
    End Select
End Function
