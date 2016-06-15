Attribute VB_Name = "wXlListDataType"
Function XlListDataTypeFromString(value As String) As XlListDataType
    If IsNumeric(value) Then
        XlListDataTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlListDataTypeNone": XlListDataTypeFromString = xlListDataTypeNone
        Case "xlListDataTypeText": XlListDataTypeFromString = xlListDataTypeText
        Case "xlListDataTypeMultiLineText": XlListDataTypeFromString = xlListDataTypeMultiLineText
        Case "xlListDataTypeNumber": XlListDataTypeFromString = xlListDataTypeNumber
        Case "xlListDataTypeCurrency": XlListDataTypeFromString = xlListDataTypeCurrency
        Case "xlListDataTypeDateTime": XlListDataTypeFromString = xlListDataTypeDateTime
        Case "xlListDataTypeChoice": XlListDataTypeFromString = xlListDataTypeChoice
        Case "xlListDataTypeChoiceMulti": XlListDataTypeFromString = xlListDataTypeChoiceMulti
        Case "xlListDataTypeListLookup": XlListDataTypeFromString = xlListDataTypeListLookup
        Case "xlListDataTypeCheckbox": XlListDataTypeFromString = xlListDataTypeCheckbox
        Case "xlListDataTypeHyperLink": XlListDataTypeFromString = xlListDataTypeHyperLink
        Case "xlListDataTypeCounter": XlListDataTypeFromString = xlListDataTypeCounter
        Case "xlListDataTypeMultiLineRichText": XlListDataTypeFromString = xlListDataTypeMultiLineRichText
    End Select
End Function

Function XlListDataTypeToString(value As XlListDataType) As String
    Select Case value
        Case xlListDataTypeNone: XlListDataTypeToString = "xlListDataTypeNone"
        Case xlListDataTypeText: XlListDataTypeToString = "xlListDataTypeText"
        Case xlListDataTypeMultiLineText: XlListDataTypeToString = "xlListDataTypeMultiLineText"
        Case xlListDataTypeNumber: XlListDataTypeToString = "xlListDataTypeNumber"
        Case xlListDataTypeCurrency: XlListDataTypeToString = "xlListDataTypeCurrency"
        Case xlListDataTypeDateTime: XlListDataTypeToString = "xlListDataTypeDateTime"
        Case xlListDataTypeChoice: XlListDataTypeToString = "xlListDataTypeChoice"
        Case xlListDataTypeChoiceMulti: XlListDataTypeToString = "xlListDataTypeChoiceMulti"
        Case xlListDataTypeListLookup: XlListDataTypeToString = "xlListDataTypeListLookup"
        Case xlListDataTypeCheckbox: XlListDataTypeToString = "xlListDataTypeCheckbox"
        Case xlListDataTypeHyperLink: XlListDataTypeToString = "xlListDataTypeHyperLink"
        Case xlListDataTypeCounter: XlListDataTypeToString = "xlListDataTypeCounter"
        Case xlListDataTypeMultiLineRichText: XlListDataTypeToString = "xlListDataTypeMultiLineRichText"
    End Select
End Function
