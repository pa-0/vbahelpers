Attribute VB_Name = "wXlXmlImportResult"
Function XlXmlImportResultFromString(value As String) As XlXmlImportResult
    If IsNumeric(value) Then
        XlXmlImportResultFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlXmlImportSuccess": XlXmlImportResultFromString = xlXmlImportSuccess
        Case "xlXmlImportElementsTruncated": XlXmlImportResultFromString = xlXmlImportElementsTruncated
        Case "xlXmlImportValidationFailed": XlXmlImportResultFromString = xlXmlImportValidationFailed
    End Select
End Function

Function XlXmlImportResultToString(value As XlXmlImportResult) As String
    Select Case value
        Case xlXmlImportSuccess: XlXmlImportResultToString = "xlXmlImportSuccess"
        Case xlXmlImportElementsTruncated: XlXmlImportResultToString = "xlXmlImportElementsTruncated"
        Case xlXmlImportValidationFailed: XlXmlImportResultToString = "xlXmlImportValidationFailed"
    End Select
End Function
