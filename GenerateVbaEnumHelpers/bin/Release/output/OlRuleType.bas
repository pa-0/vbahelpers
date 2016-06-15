Attribute VB_Name = "wOlRuleType"
Function OlRuleTypeFromString(value As String) As OlRuleType
    If IsNumeric(value) Then
        OlRuleTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "olRuleReceive": OlRuleTypeFromString = olRuleReceive
        Case "olRuleSend": OlRuleTypeFromString = olRuleSend
    End Select
End Function

Function OlRuleTypeToString(value As OlRuleType) As String
    Select Case value
        Case olRuleReceive: OlRuleTypeToString = "olRuleReceive"
        Case olRuleSend: OlRuleTypeToString = "olRuleSend"
    End Select
End Function
