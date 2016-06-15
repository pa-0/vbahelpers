Attribute VB_Name = "wWdTextFormFieldType"
Function WdTextFormFieldTypeFromString(value As String) As WdTextFormFieldType
    If IsNumeric(value) Then
        WdTextFormFieldTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdRegularText": WdTextFormFieldTypeFromString = wdRegularText
        Case "wdNumberText": WdTextFormFieldTypeFromString = wdNumberText
        Case "wdDateText": WdTextFormFieldTypeFromString = wdDateText
        Case "wdCurrentDateText": WdTextFormFieldTypeFromString = wdCurrentDateText
        Case "wdCurrentTimeText": WdTextFormFieldTypeFromString = wdCurrentTimeText
        Case "wdCalculationText": WdTextFormFieldTypeFromString = wdCalculationText
    End Select
End Function

Function WdTextFormFieldTypeToString(value As WdTextFormFieldType) As String
    Select Case value
        Case wdRegularText: WdTextFormFieldTypeToString = "wdRegularText"
        Case wdNumberText: WdTextFormFieldTypeToString = "wdNumberText"
        Case wdDateText: WdTextFormFieldTypeToString = "wdDateText"
        Case wdCurrentDateText: WdTextFormFieldTypeToString = "wdCurrentDateText"
        Case wdCurrentTimeText: WdTextFormFieldTypeToString = "wdCurrentTimeText"
        Case wdCalculationText: WdTextFormFieldTypeToString = "wdCalculationText"
    End Select
End Function
