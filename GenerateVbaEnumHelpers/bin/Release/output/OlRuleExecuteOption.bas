Attribute VB_Name = "wOlRuleExecuteOption"
Function OlRuleExecuteOptionFromString(value As String) As OlRuleExecuteOption
    If IsNumeric(value) Then
        OlRuleExecuteOptionFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "olRuleExecuteAllMessages": OlRuleExecuteOptionFromString = olRuleExecuteAllMessages
        Case "olRuleExecuteReadMessages": OlRuleExecuteOptionFromString = olRuleExecuteReadMessages
        Case "olRuleExecuteUnreadMessages": OlRuleExecuteOptionFromString = olRuleExecuteUnreadMessages
    End Select
End Function

Function OlRuleExecuteOptionToString(value As OlRuleExecuteOption) As String
    Select Case value
        Case olRuleExecuteAllMessages: OlRuleExecuteOptionToString = "olRuleExecuteAllMessages"
        Case olRuleExecuteReadMessages: OlRuleExecuteOptionToString = "olRuleExecuteReadMessages"
        Case olRuleExecuteUnreadMessages: OlRuleExecuteOptionToString = "olRuleExecuteUnreadMessages"
    End Select
End Function
