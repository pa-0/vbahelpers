Attribute VB_Name = "wXlTextQualifier"
Function XlTextQualifierFromString(value As String) As XlTextQualifier
    If IsNumeric(value) Then
        XlTextQualifierFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlTextQualifierDoubleQuote": XlTextQualifierFromString = xlTextQualifierDoubleQuote
        Case "xlTextQualifierSingleQuote": XlTextQualifierFromString = xlTextQualifierSingleQuote
        Case "xlTextQualifierNone": XlTextQualifierFromString = xlTextQualifierNone
    End Select
End Function

Function XlTextQualifierToString(value As XlTextQualifier) As String
    Select Case value
        Case xlTextQualifierDoubleQuote: XlTextQualifierToString = "xlTextQualifierDoubleQuote"
        Case xlTextQualifierSingleQuote: XlTextQualifierToString = "xlTextQualifierSingleQuote"
        Case xlTextQualifierNone: XlTextQualifierToString = "xlTextQualifierNone"
    End Select
End Function
