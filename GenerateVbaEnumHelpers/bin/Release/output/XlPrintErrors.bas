Attribute VB_Name = "wXlPrintErrors"
Function XlPrintErrorsFromString(value As String) As XlPrintErrors
    If IsNumeric(value) Then
        XlPrintErrorsFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlPrintErrorsDisplayed": XlPrintErrorsFromString = xlPrintErrorsDisplayed
        Case "xlPrintErrorsBlank": XlPrintErrorsFromString = xlPrintErrorsBlank
        Case "xlPrintErrorsDash": XlPrintErrorsFromString = xlPrintErrorsDash
        Case "xlPrintErrorsNA": XlPrintErrorsFromString = xlPrintErrorsNA
    End Select
End Function

Function XlPrintErrorsToString(value As XlPrintErrors) As String
    Select Case value
        Case xlPrintErrorsDisplayed: XlPrintErrorsToString = "xlPrintErrorsDisplayed"
        Case xlPrintErrorsBlank: XlPrintErrorsToString = "xlPrintErrorsBlank"
        Case xlPrintErrorsDash: XlPrintErrorsToString = "xlPrintErrorsDash"
        Case xlPrintErrorsNA: XlPrintErrorsToString = "xlPrintErrorsNA"
    End Select
End Function
