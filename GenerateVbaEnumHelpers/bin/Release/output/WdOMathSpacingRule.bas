Attribute VB_Name = "wWdOMathSpacingRule"
Function WdOMathSpacingRuleFromString(value As String) As WdOMathSpacingRule
    If IsNumeric(value) Then
        WdOMathSpacingRuleFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdOMathSpacingSingle": WdOMathSpacingRuleFromString = wdOMathSpacingSingle
        Case "wdOMathSpacing1pt5": WdOMathSpacingRuleFromString = wdOMathSpacing1pt5
        Case "wdOMathSpacingDouble": WdOMathSpacingRuleFromString = wdOMathSpacingDouble
        Case "wdOMathSpacingExactly": WdOMathSpacingRuleFromString = wdOMathSpacingExactly
        Case "wdOMathSpacingMultiple": WdOMathSpacingRuleFromString = wdOMathSpacingMultiple
    End Select
End Function

Function WdOMathSpacingRuleToString(value As WdOMathSpacingRule) As String
    Select Case value
        Case wdOMathSpacingSingle: WdOMathSpacingRuleToString = "wdOMathSpacingSingle"
        Case wdOMathSpacing1pt5: WdOMathSpacingRuleToString = "wdOMathSpacing1pt5"
        Case wdOMathSpacingDouble: WdOMathSpacingRuleToString = "wdOMathSpacingDouble"
        Case wdOMathSpacingExactly: WdOMathSpacingRuleToString = "wdOMathSpacingExactly"
        Case wdOMathSpacingMultiple: WdOMathSpacingRuleToString = "wdOMathSpacingMultiple"
    End Select
End Function
