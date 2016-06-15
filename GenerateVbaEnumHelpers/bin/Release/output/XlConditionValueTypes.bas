Attribute VB_Name = "wXlConditionValueTypes"
Function XlConditionValueTypesFromString(value As String) As XlConditionValueTypes
    If IsNumeric(value) Then
        XlConditionValueTypesFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlConditionValueNumber": XlConditionValueTypesFromString = xlConditionValueNumber
        Case "xlConditionValueLowestValue": XlConditionValueTypesFromString = xlConditionValueLowestValue
        Case "xlConditionValueHighestValue": XlConditionValueTypesFromString = xlConditionValueHighestValue
        Case "xlConditionValuePercent": XlConditionValueTypesFromString = xlConditionValuePercent
        Case "xlConditionValueFormula": XlConditionValueTypesFromString = xlConditionValueFormula
        Case "xlConditionValuePercentile": XlConditionValueTypesFromString = xlConditionValuePercentile
        Case "xlConditionValueAutomaticMin": XlConditionValueTypesFromString = xlConditionValueAutomaticMin
        Case "xlConditionValueAutomaticMax": XlConditionValueTypesFromString = xlConditionValueAutomaticMax
        Case "xlConditionValueNone": XlConditionValueTypesFromString = xlConditionValueNone
    End Select
End Function

Function XlConditionValueTypesToString(value As XlConditionValueTypes) As String
    Select Case value
        Case xlConditionValueNumber: XlConditionValueTypesToString = "xlConditionValueNumber"
        Case xlConditionValueLowestValue: XlConditionValueTypesToString = "xlConditionValueLowestValue"
        Case xlConditionValueHighestValue: XlConditionValueTypesToString = "xlConditionValueHighestValue"
        Case xlConditionValuePercent: XlConditionValueTypesToString = "xlConditionValuePercent"
        Case xlConditionValueFormula: XlConditionValueTypesToString = "xlConditionValueFormula"
        Case xlConditionValuePercentile: XlConditionValueTypesToString = "xlConditionValuePercentile"
        Case xlConditionValueAutomaticMin: XlConditionValueTypesToString = "xlConditionValueAutomaticMin"
        Case xlConditionValueAutomaticMax: XlConditionValueTypesToString = "xlConditionValueAutomaticMax"
        Case xlConditionValueNone: XlConditionValueTypesToString = "xlConditionValueNone"
    End Select
End Function
