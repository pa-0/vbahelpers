Attribute VB_Name = "wXlRangeValueDataType"
Function XlRangeValueDataTypeFromString(value As String) As XlRangeValueDataType
    If IsNumeric(value) Then
        XlRangeValueDataTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlRangeValueDefault": XlRangeValueDataTypeFromString = xlRangeValueDefault
        Case "xlRangeValueXMLSpreadsheet": XlRangeValueDataTypeFromString = xlRangeValueXMLSpreadsheet
        Case "xlRangeValueMSPersistXML": XlRangeValueDataTypeFromString = xlRangeValueMSPersistXML
    End Select
End Function

Function XlRangeValueDataTypeToString(value As XlRangeValueDataType) As String
    Select Case value
        Case xlRangeValueDefault: XlRangeValueDataTypeToString = "xlRangeValueDefault"
        Case xlRangeValueXMLSpreadsheet: XlRangeValueDataTypeToString = "xlRangeValueXMLSpreadsheet"
        Case xlRangeValueMSPersistXML: XlRangeValueDataTypeToString = "xlRangeValueMSPersistXML"
    End Select
End Function
