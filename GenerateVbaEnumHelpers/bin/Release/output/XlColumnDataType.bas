Attribute VB_Name = "wXlColumnDataType"
Function XlColumnDataTypeFromString(value As String) As XlColumnDataType
    If IsNumeric(value) Then
        XlColumnDataTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlGeneralFormat": XlColumnDataTypeFromString = xlGeneralFormat
        Case "xlTextFormat": XlColumnDataTypeFromString = xlTextFormat
        Case "xlMDYFormat": XlColumnDataTypeFromString = xlMDYFormat
        Case "xlDMYFormat": XlColumnDataTypeFromString = xlDMYFormat
        Case "xlYMDFormat": XlColumnDataTypeFromString = xlYMDFormat
        Case "xlMYDFormat": XlColumnDataTypeFromString = xlMYDFormat
        Case "xlDYMFormat": XlColumnDataTypeFromString = xlDYMFormat
        Case "xlYDMFormat": XlColumnDataTypeFromString = xlYDMFormat
        Case "xlSkipColumn": XlColumnDataTypeFromString = xlSkipColumn
        Case "xlEMDFormat": XlColumnDataTypeFromString = xlEMDFormat
    End Select
End Function

Function XlColumnDataTypeToString(value As XlColumnDataType) As String
    Select Case value
        Case xlGeneralFormat: XlColumnDataTypeToString = "xlGeneralFormat"
        Case xlTextFormat: XlColumnDataTypeToString = "xlTextFormat"
        Case xlMDYFormat: XlColumnDataTypeToString = "xlMDYFormat"
        Case xlDMYFormat: XlColumnDataTypeToString = "xlDMYFormat"
        Case xlYMDFormat: XlColumnDataTypeToString = "xlYMDFormat"
        Case xlMYDFormat: XlColumnDataTypeToString = "xlMYDFormat"
        Case xlDYMFormat: XlColumnDataTypeToString = "xlDYMFormat"
        Case xlYDMFormat: XlColumnDataTypeToString = "xlYDMFormat"
        Case xlSkipColumn: XlColumnDataTypeToString = "xlSkipColumn"
        Case xlEMDFormat: XlColumnDataTypeToString = "xlEMDFormat"
    End Select
End Function
