Attribute VB_Name = "wXlQueryType"
Function XlQueryTypeFromString(value As String) As XlQueryType
    If IsNumeric(value) Then
        XlQueryTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlODBCQuery": XlQueryTypeFromString = xlODBCQuery
        Case "xlDAORecordset": XlQueryTypeFromString = xlDAORecordset
        Case "xlWebQuery": XlQueryTypeFromString = xlWebQuery
        Case "xlOLEDBQuery": XlQueryTypeFromString = xlOLEDBQuery
        Case "xlTextImport": XlQueryTypeFromString = xlTextImport
        Case "xlADORecordset": XlQueryTypeFromString = xlADORecordset
    End Select
End Function

Function XlQueryTypeToString(value As XlQueryType) As String
    Select Case value
        Case xlODBCQuery: XlQueryTypeToString = "xlODBCQuery"
        Case xlDAORecordset: XlQueryTypeToString = "xlDAORecordset"
        Case xlWebQuery: XlQueryTypeToString = "xlWebQuery"
        Case xlOLEDBQuery: XlQueryTypeToString = "xlOLEDBQuery"
        Case xlTextImport: XlQueryTypeToString = "xlTextImport"
        Case xlADORecordset: XlQueryTypeToString = "xlADORecordset"
    End Select
End Function
