Attribute VB_Name = "wWdMailerPriority"
Function WdMailerPriorityFromString(value As String) As WdMailerPriority
    If IsNumeric(value) Then
        WdMailerPriorityFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdPriorityNormal": WdMailerPriorityFromString = wdPriorityNormal
        Case "wdPriorityLow": WdMailerPriorityFromString = wdPriorityLow
        Case "wdPriorityHigh": WdMailerPriorityFromString = wdPriorityHigh
    End Select
End Function

Function WdMailerPriorityToString(value As WdMailerPriority) As String
    Select Case value
        Case wdPriorityNormal: WdMailerPriorityToString = "wdPriorityNormal"
        Case wdPriorityLow: WdMailerPriorityToString = "wdPriorityLow"
        Case wdPriorityHigh: WdMailerPriorityToString = "wdPriorityHigh"
    End Select
End Function
