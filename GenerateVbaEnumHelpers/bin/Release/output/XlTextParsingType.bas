Attribute VB_Name = "wXlTextParsingType"
Function XlTextParsingTypeFromString(value As String) As XlTextParsingType
    If IsNumeric(value) Then
        XlTextParsingTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlDelimited": XlTextParsingTypeFromString = xlDelimited
        Case "xlFixedWidth": XlTextParsingTypeFromString = xlFixedWidth
    End Select
End Function

Function XlTextParsingTypeToString(value As XlTextParsingType) As String
    Select Case value
        Case xlDelimited: XlTextParsingTypeToString = "xlDelimited"
        Case xlFixedWidth: XlTextParsingTypeToString = "xlFixedWidth"
    End Select
End Function
