Attribute VB_Name = "wXlSortDataOption"
Function XlSortDataOptionFromString(value As String) As XlSortDataOption
    If IsNumeric(value) Then
        XlSortDataOptionFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlSortNormal": XlSortDataOptionFromString = xlSortNormal
        Case "xlSortTextAsNumbers": XlSortDataOptionFromString = xlSortTextAsNumbers
    End Select
End Function

Function XlSortDataOptionToString(value As XlSortDataOption) As String
    Select Case value
        Case xlSortNormal: XlSortDataOptionToString = "xlSortNormal"
        Case xlSortTextAsNumbers: XlSortDataOptionToString = "xlSortTextAsNumbers"
    End Select
End Function
