Attribute VB_Name = "wPpFixedFormatType"
Function PpFixedFormatTypeFromString(value As String) As PpFixedFormatType
    If IsNumeric(value) Then
        PpFixedFormatTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "ppFixedFormatTypeXPS": PpFixedFormatTypeFromString = ppFixedFormatTypeXPS
        Case "ppFixedFormatTypePDF": PpFixedFormatTypeFromString = ppFixedFormatTypePDF
    End Select
End Function

Function PpFixedFormatTypeToString(value As PpFixedFormatType) As String
    Select Case value
        Case ppFixedFormatTypeXPS: PpFixedFormatTypeToString = "ppFixedFormatTypeXPS"
        Case ppFixedFormatTypePDF: PpFixedFormatTypeToString = "ppFixedFormatTypePDF"
    End Select
End Function
