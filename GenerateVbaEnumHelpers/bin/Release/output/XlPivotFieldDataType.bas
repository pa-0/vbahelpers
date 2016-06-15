Attribute VB_Name = "wXlPivotFieldDataType"
Function XlPivotFieldDataTypeFromString(value As String) As XlPivotFieldDataType
    If IsNumeric(value) Then
        XlPivotFieldDataTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlDate": XlPivotFieldDataTypeFromString = xlDate
        Case "xlText": XlPivotFieldDataTypeFromString = xlText
        Case "xlNumber": XlPivotFieldDataTypeFromString = xlNumber
    End Select
End Function

Function XlPivotFieldDataTypeToString(value As XlPivotFieldDataType) As String
    Select Case value
        Case xlDate: XlPivotFieldDataTypeToString = "xlDate"
        Case xlText: XlPivotFieldDataTypeToString = "xlText"
        Case xlNumber: XlPivotFieldDataTypeToString = "xlNumber"
    End Select
End Function
