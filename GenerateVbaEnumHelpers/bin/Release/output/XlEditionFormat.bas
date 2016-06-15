Attribute VB_Name = "wXlEditionFormat"
Function XlEditionFormatFromString(value As String) As XlEditionFormat
    If IsNumeric(value) Then
        XlEditionFormatFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlPICT": XlEditionFormatFromString = xlPICT
        Case "xlBIFF": XlEditionFormatFromString = xlBIFF
        Case "xlRTF": XlEditionFormatFromString = xlRTF
        Case "xlVALU": XlEditionFormatFromString = xlVALU
    End Select
End Function

Function XlEditionFormatToString(value As XlEditionFormat) As String
    Select Case value
        Case xlPICT: XlEditionFormatToString = "xlPICT"
        Case xlBIFF: XlEditionFormatToString = "xlBIFF"
        Case xlRTF: XlEditionFormatToString = "xlRTF"
        Case xlVALU: XlEditionFormatToString = "xlVALU"
    End Select
End Function
