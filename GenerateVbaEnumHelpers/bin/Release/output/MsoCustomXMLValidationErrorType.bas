Attribute VB_Name = "wMsoCustomXMLValidationErrorType"
Function MsoCustomXMLValidationErrorTypeFromString(value As String) As MsoCustomXMLValidationErrorType
    If IsNumeric(value) Then
        MsoCustomXMLValidationErrorTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoCustomXMLValidationErrorSchemaGenerated": MsoCustomXMLValidationErrorTypeFromString = msoCustomXMLValidationErrorSchemaGenerated
        Case "msoCustomXMLValidationErrorAutomaticallyCleared": MsoCustomXMLValidationErrorTypeFromString = msoCustomXMLValidationErrorAutomaticallyCleared
        Case "msoCustomXMLValidationErrorManual": MsoCustomXMLValidationErrorTypeFromString = msoCustomXMLValidationErrorManual
    End Select
End Function

Function MsoCustomXMLValidationErrorTypeToString(value As MsoCustomXMLValidationErrorType) As String
    Select Case value
        Case msoCustomXMLValidationErrorSchemaGenerated: MsoCustomXMLValidationErrorTypeToString = "msoCustomXMLValidationErrorSchemaGenerated"
        Case msoCustomXMLValidationErrorAutomaticallyCleared: MsoCustomXMLValidationErrorTypeToString = "msoCustomXMLValidationErrorAutomaticallyCleared"
        Case msoCustomXMLValidationErrorManual: MsoCustomXMLValidationErrorTypeToString = "msoCustomXMLValidationErrorManual"
    End Select
End Function
