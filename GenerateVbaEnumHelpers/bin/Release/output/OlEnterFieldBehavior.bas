Attribute VB_Name = "wOlEnterFieldBehavior"
Function OlEnterFieldBehaviorFromString(value As String) As OlEnterFieldBehavior
    If IsNumeric(value) Then
        OlEnterFieldBehaviorFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "olEnterFieldBehaviorSelectAll": OlEnterFieldBehaviorFromString = olEnterFieldBehaviorSelectAll
        Case "olEnterFieldBehaviorRecallSelection": OlEnterFieldBehaviorFromString = olEnterFieldBehaviorRecallSelection
    End Select
End Function

Function OlEnterFieldBehaviorToString(value As OlEnterFieldBehavior) As String
    Select Case value
        Case olEnterFieldBehaviorSelectAll: OlEnterFieldBehaviorToString = "olEnterFieldBehaviorSelectAll"
        Case olEnterFieldBehaviorRecallSelection: OlEnterFieldBehaviorToString = "olEnterFieldBehaviorRecallSelection"
    End Select
End Function
