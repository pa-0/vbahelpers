Attribute VB_Name = "wWdChevronConvertRule"
Function WdChevronConvertRuleFromString(value As String) As WdChevronConvertRule
    If IsNumeric(value) Then
        WdChevronConvertRuleFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdNeverConvert": WdChevronConvertRuleFromString = wdNeverConvert
        Case "wdAlwaysConvert": WdChevronConvertRuleFromString = wdAlwaysConvert
        Case "wdAskToNotConvert": WdChevronConvertRuleFromString = wdAskToNotConvert
        Case "wdAskToConvert": WdChevronConvertRuleFromString = wdAskToConvert
    End Select
End Function

Function WdChevronConvertRuleToString(value As WdChevronConvertRule) As String
    Select Case value
        Case wdNeverConvert: WdChevronConvertRuleToString = "wdNeverConvert"
        Case wdAlwaysConvert: WdChevronConvertRuleToString = "wdAlwaysConvert"
        Case wdAskToNotConvert: WdChevronConvertRuleToString = "wdAskToNotConvert"
        Case wdAskToConvert: WdChevronConvertRuleToString = "wdAskToConvert"
    End Select
End Function
