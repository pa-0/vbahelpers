Attribute VB_Name = "wXlPasteType"
Function XlPasteTypeFromString(value As String) As XlPasteType
    If IsNumeric(value) Then
        XlPasteTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlPasteValidation": XlPasteTypeFromString = xlPasteValidation
        Case "xlPasteAllExceptBorders": XlPasteTypeFromString = xlPasteAllExceptBorders
        Case "xlPasteColumnWidths": XlPasteTypeFromString = xlPasteColumnWidths
        Case "xlPasteFormulasAndNumberFormats": XlPasteTypeFromString = xlPasteFormulasAndNumberFormats
        Case "xlPasteValuesAndNumberFormats": XlPasteTypeFromString = xlPasteValuesAndNumberFormats
        Case "xlPasteAllUsingSourceTheme": XlPasteTypeFromString = xlPasteAllUsingSourceTheme
        Case "xlPasteAllMergingConditionalFormats": XlPasteTypeFromString = xlPasteAllMergingConditionalFormats
        Case "xlPasteValues": XlPasteTypeFromString = xlPasteValues
        Case "xlPasteComments": XlPasteTypeFromString = xlPasteComments
        Case "xlPasteFormulas": XlPasteTypeFromString = xlPasteFormulas
        Case "xlPasteFormats": XlPasteTypeFromString = xlPasteFormats
        Case "xlPasteAll": XlPasteTypeFromString = xlPasteAll
    End Select
End Function

Function XlPasteTypeToString(value As XlPasteType) As String
    Select Case value
        Case xlPasteValidation: XlPasteTypeToString = "xlPasteValidation"
        Case xlPasteAllExceptBorders: XlPasteTypeToString = "xlPasteAllExceptBorders"
        Case xlPasteColumnWidths: XlPasteTypeToString = "xlPasteColumnWidths"
        Case xlPasteFormulasAndNumberFormats: XlPasteTypeToString = "xlPasteFormulasAndNumberFormats"
        Case xlPasteValuesAndNumberFormats: XlPasteTypeToString = "xlPasteValuesAndNumberFormats"
        Case xlPasteAllUsingSourceTheme: XlPasteTypeToString = "xlPasteAllUsingSourceTheme"
        Case xlPasteAllMergingConditionalFormats: XlPasteTypeToString = "xlPasteAllMergingConditionalFormats"
        Case xlPasteValues: XlPasteTypeToString = "xlPasteValues"
        Case xlPasteComments: XlPasteTypeToString = "xlPasteComments"
        Case xlPasteFormulas: XlPasteTypeToString = "xlPasteFormulas"
        Case xlPasteFormats: XlPasteTypeToString = "xlPasteFormats"
        Case xlPasteAll: XlPasteTypeToString = "xlPasteAll"
    End Select
End Function
