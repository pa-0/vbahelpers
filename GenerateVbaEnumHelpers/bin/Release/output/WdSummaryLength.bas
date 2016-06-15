Attribute VB_Name = "wWdSummaryLength"
Function WdSummaryLengthFromString(value As String) As WdSummaryLength
    If IsNumeric(value) Then
        WdSummaryLengthFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wd75Percent": WdSummaryLengthFromString = wd75Percent
        Case "wd50Percent": WdSummaryLengthFromString = wd50Percent
        Case "wd25Percent": WdSummaryLengthFromString = wd25Percent
        Case "wd10Percent": WdSummaryLengthFromString = wd10Percent
        Case "wd500Words": WdSummaryLengthFromString = wd500Words
        Case "wd100Words": WdSummaryLengthFromString = wd100Words
        Case "wd20Sentences": WdSummaryLengthFromString = wd20Sentences
        Case "wd10Sentences": WdSummaryLengthFromString = wd10Sentences
    End Select
End Function

Function WdSummaryLengthToString(value As WdSummaryLength) As String
    Select Case value
        Case wd75Percent: WdSummaryLengthToString = "wd75Percent"
        Case wd50Percent: WdSummaryLengthToString = "wd50Percent"
        Case wd25Percent: WdSummaryLengthToString = "wd25Percent"
        Case wd10Percent: WdSummaryLengthToString = "wd10Percent"
        Case wd500Words: WdSummaryLengthToString = "wd500Words"
        Case wd100Words: WdSummaryLengthToString = "wd100Words"
        Case wd20Sentences: WdSummaryLengthToString = "wd20Sentences"
        Case wd10Sentences: WdSummaryLengthToString = "wd10Sentences"
    End Select
End Function
