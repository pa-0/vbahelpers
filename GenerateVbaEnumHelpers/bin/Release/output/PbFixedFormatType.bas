Attribute VB_Name = "wPbFixedFormatType"
Function PbFixedFormatTypeFromString(value As String) As PbFixedFormatType
    If IsNumeric(value) Then
        PbFixedFormatTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "pbFixedFormatTypeXPS": PbFixedFormatTypeFromString = pbFixedFormatTypeXPS
        Case "pbFixedFormatTypePDF": PbFixedFormatTypeFromString = pbFixedFormatTypePDF
    End Select
End Function

Function PbFixedFormatTypeToString(value As PbFixedFormatType) As String
    Select Case value
        Case pbFixedFormatTypeXPS: PbFixedFormatTypeToString = "pbFixedFormatTypeXPS"
        Case pbFixedFormatTypePDF: PbFixedFormatTypeToString = "pbFixedFormatTypePDF"
    End Select
End Function
