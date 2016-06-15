Attribute VB_Name = "wOlUserPropertyType"
Function OlUserPropertyTypeFromString(value As String) As OlUserPropertyType
    If IsNumeric(value) Then
        OlUserPropertyTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "olOutlookInternal": OlUserPropertyTypeFromString = olOutlookInternal
        Case "olText": OlUserPropertyTypeFromString = olText
        Case "olNumber": OlUserPropertyTypeFromString = olNumber
        Case "olDateTime": OlUserPropertyTypeFromString = olDateTime
        Case "olYesNo": OlUserPropertyTypeFromString = olYesNo
        Case "olDuration": OlUserPropertyTypeFromString = olDuration
        Case "olKeywords": OlUserPropertyTypeFromString = olKeywords
        Case "olPercent": OlUserPropertyTypeFromString = olPercent
        Case "olCurrency": OlUserPropertyTypeFromString = olCurrency
        Case "olFormula": OlUserPropertyTypeFromString = olFormula
        Case "olCombination": OlUserPropertyTypeFromString = olCombination
        Case "olInteger": OlUserPropertyTypeFromString = olInteger
        Case "olEnumeration": OlUserPropertyTypeFromString = olEnumeration
        Case "olSmartFrom": OlUserPropertyTypeFromString = olSmartFrom
    End Select
End Function

Function OlUserPropertyTypeToString(value As OlUserPropertyType) As String
    Select Case value
        Case olOutlookInternal: OlUserPropertyTypeToString = "olOutlookInternal"
        Case olText: OlUserPropertyTypeToString = "olText"
        Case olNumber: OlUserPropertyTypeToString = "olNumber"
        Case olDateTime: OlUserPropertyTypeToString = "olDateTime"
        Case olYesNo: OlUserPropertyTypeToString = "olYesNo"
        Case olDuration: OlUserPropertyTypeToString = "olDuration"
        Case olKeywords: OlUserPropertyTypeToString = "olKeywords"
        Case olPercent: OlUserPropertyTypeToString = "olPercent"
        Case olCurrency: OlUserPropertyTypeToString = "olCurrency"
        Case olFormula: OlUserPropertyTypeToString = "olFormula"
        Case olCombination: OlUserPropertyTypeToString = "olCombination"
        Case olInteger: OlUserPropertyTypeToString = "olInteger"
        Case olEnumeration: OlUserPropertyTypeToString = "olEnumeration"
        Case olSmartFrom: OlUserPropertyTypeToString = "olSmartFrom"
    End Select
End Function
