Attribute VB_Name = "wWdNumberingRule"
Function WdNumberingRuleFromString(value As String) As WdNumberingRule
    If IsNumeric(value) Then
        WdNumberingRuleFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdRestartContinuous": WdNumberingRuleFromString = wdRestartContinuous
        Case "wdRestartSection": WdNumberingRuleFromString = wdRestartSection
        Case "wdRestartPage": WdNumberingRuleFromString = wdRestartPage
    End Select
End Function

Function WdNumberingRuleToString(value As WdNumberingRule) As String
    Select Case value
        Case wdRestartContinuous: WdNumberingRuleToString = "wdRestartContinuous"
        Case wdRestartSection: WdNumberingRuleToString = "wdRestartSection"
        Case wdRestartPage: WdNumberingRuleToString = "wdRestartPage"
    End Select
End Function
