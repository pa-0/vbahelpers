Attribute VB_Name = "wPbLineSpacingRule"
Function PbLineSpacingRuleFromString(value As String) As PbLineSpacingRule
    If IsNumeric(value) Then
        PbLineSpacingRuleFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "pbLineSpacingSingle": PbLineSpacingRuleFromString = pbLineSpacingSingle
        Case "pbLineSpacing1pt5": PbLineSpacingRuleFromString = pbLineSpacing1pt5
        Case "pbLineSpacingDouble": PbLineSpacingRuleFromString = pbLineSpacingDouble
        Case "pbLineSpacingExactly": PbLineSpacingRuleFromString = pbLineSpacingExactly
        Case "pbLineSpacingMultiple": PbLineSpacingRuleFromString = pbLineSpacingMultiple
        Case "pbLineSpacingMixed": PbLineSpacingRuleFromString = pbLineSpacingMixed
    End Select
End Function

Function PbLineSpacingRuleToString(value As PbLineSpacingRule) As String
    Select Case value
        Case pbLineSpacingSingle: PbLineSpacingRuleToString = "pbLineSpacingSingle"
        Case pbLineSpacing1pt5: PbLineSpacingRuleToString = "pbLineSpacing1pt5"
        Case pbLineSpacingDouble: PbLineSpacingRuleToString = "pbLineSpacingDouble"
        Case pbLineSpacingExactly: PbLineSpacingRuleToString = "pbLineSpacingExactly"
        Case pbLineSpacingMultiple: PbLineSpacingRuleToString = "pbLineSpacingMultiple"
        Case pbLineSpacingMixed: PbLineSpacingRuleToString = "pbLineSpacingMixed"
    End Select
End Function
