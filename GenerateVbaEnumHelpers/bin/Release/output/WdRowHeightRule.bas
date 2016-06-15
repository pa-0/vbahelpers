Attribute VB_Name = "wWdRowHeightRule"
Function WdRowHeightRuleFromString(value As String) As WdRowHeightRule
    If IsNumeric(value) Then
        WdRowHeightRuleFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdRowHeightAuto": WdRowHeightRuleFromString = wdRowHeightAuto
        Case "wdRowHeightAtLeast": WdRowHeightRuleFromString = wdRowHeightAtLeast
        Case "wdRowHeightExactly": WdRowHeightRuleFromString = wdRowHeightExactly
    End Select
End Function

Function WdRowHeightRuleToString(value As WdRowHeightRule) As String
    Select Case value
        Case wdRowHeightAuto: WdRowHeightRuleToString = "wdRowHeightAuto"
        Case wdRowHeightAtLeast: WdRowHeightRuleToString = "wdRowHeightAtLeast"
        Case wdRowHeightExactly: WdRowHeightRuleToString = "wdRowHeightExactly"
    End Select
End Function
