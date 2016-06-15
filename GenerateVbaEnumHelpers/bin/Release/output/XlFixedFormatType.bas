Attribute VB_Name = "wXlFixedFormatType"
Function XlFixedFormatTypeFromString(value As String) As XlFixedFormatType
    If IsNumeric(value) Then
        XlFixedFormatTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlTypePDF": XlFixedFormatTypeFromString = xlTypePDF
        Case "xlTypeXPS": XlFixedFormatTypeFromString = xlTypeXPS
    End Select
End Function

Function XlFixedFormatTypeToString(value As XlFixedFormatType) As String
    Select Case value
        Case xlTypePDF: XlFixedFormatTypeToString = "xlTypePDF"
        Case xlTypeXPS: XlFixedFormatTypeToString = "xlTypeXPS"
    End Select
End Function
