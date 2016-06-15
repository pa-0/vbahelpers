Attribute VB_Name = "wWdMailSystem"
Function WdMailSystemFromString(value As String) As WdMailSystem
    If IsNumeric(value) Then
        WdMailSystemFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdNoMailSystem": WdMailSystemFromString = wdNoMailSystem
        Case "wdMAPI": WdMailSystemFromString = wdMAPI
        Case "wdPowerTalk": WdMailSystemFromString = wdPowerTalk
        Case "wdMAPIandPowerTalk": WdMailSystemFromString = wdMAPIandPowerTalk
    End Select
End Function

Function WdMailSystemToString(value As WdMailSystem) As String
    Select Case value
        Case wdNoMailSystem: WdMailSystemToString = "wdNoMailSystem"
        Case wdMAPI: WdMailSystemToString = "wdMAPI"
        Case wdPowerTalk: WdMailSystemToString = "wdPowerTalk"
        Case wdMAPIandPowerTalk: WdMailSystemToString = "wdMAPIandPowerTalk"
    End Select
End Function
