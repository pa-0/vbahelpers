Attribute VB_Name = "wXlFormatConditionType"
Function XlFormatConditionTypeFromString(value As String) As XlFormatConditionType
    If IsNumeric(value) Then
        XlFormatConditionTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlCellValue": XlFormatConditionTypeFromString = xlCellValue
        Case "xlExpression": XlFormatConditionTypeFromString = xlExpression
        Case "xlColorScale": XlFormatConditionTypeFromString = xlColorScale
        Case "xlDatabar": XlFormatConditionTypeFromString = xlDatabar
        Case "xlTop10": XlFormatConditionTypeFromString = xlTop10
        Case "xlIconSets": XlFormatConditionTypeFromString = xlIconSets
        Case "xlUniqueValues": XlFormatConditionTypeFromString = xlUniqueValues
        Case "xlTextString": XlFormatConditionTypeFromString = xlTextString
        Case "xlBlanksCondition": XlFormatConditionTypeFromString = xlBlanksCondition
        Case "xlTimePeriod": XlFormatConditionTypeFromString = xlTimePeriod
        Case "xlAboveAverageCondition": XlFormatConditionTypeFromString = xlAboveAverageCondition
        Case "xlNoBlanksCondition": XlFormatConditionTypeFromString = xlNoBlanksCondition
        Case "xlErrorsCondition": XlFormatConditionTypeFromString = xlErrorsCondition
        Case "xlNoErrorsCondition": XlFormatConditionTypeFromString = xlNoErrorsCondition
    End Select
End Function

Function XlFormatConditionTypeToString(value As XlFormatConditionType) As String
    Select Case value
        Case xlCellValue: XlFormatConditionTypeToString = "xlCellValue"
        Case xlExpression: XlFormatConditionTypeToString = "xlExpression"
        Case xlColorScale: XlFormatConditionTypeToString = "xlColorScale"
        Case xlDatabar: XlFormatConditionTypeToString = "xlDatabar"
        Case xlTop10: XlFormatConditionTypeToString = "xlTop10"
        Case xlIconSets: XlFormatConditionTypeToString = "xlIconSets"
        Case xlUniqueValues: XlFormatConditionTypeToString = "xlUniqueValues"
        Case xlTextString: XlFormatConditionTypeToString = "xlTextString"
        Case xlBlanksCondition: XlFormatConditionTypeToString = "xlBlanksCondition"
        Case xlTimePeriod: XlFormatConditionTypeToString = "xlTimePeriod"
        Case xlAboveAverageCondition: XlFormatConditionTypeToString = "xlAboveAverageCondition"
        Case xlNoBlanksCondition: XlFormatConditionTypeToString = "xlNoBlanksCondition"
        Case xlErrorsCondition: XlFormatConditionTypeToString = "xlErrorsCondition"
        Case xlNoErrorsCondition: XlFormatConditionTypeToString = "xlNoErrorsCondition"
    End Select
End Function
