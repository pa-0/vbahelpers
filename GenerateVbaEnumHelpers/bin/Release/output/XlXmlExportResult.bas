Attribute VB_Name = "wXlXmlExportResult"
Function XlXmlExportResultFromString(value As String) As XlXmlExportResult
    If IsNumeric(value) Then
        XlXmlExportResultFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlXmlExportSuccess": XlXmlExportResultFromString = xlXmlExportSuccess
        Case "xlXmlExportValidationFailed": XlXmlExportResultFromString = xlXmlExportValidationFailed
    End Select
End Function

Function XlXmlExportResultToString(value As XlXmlExportResult) As String
    Select Case value
        Case xlXmlExportSuccess: XlXmlExportResultToString = "xlXmlExportSuccess"
        Case xlXmlExportValidationFailed: XlXmlExportResultToString = "xlXmlExportValidationFailed"
    End Select
End Function
