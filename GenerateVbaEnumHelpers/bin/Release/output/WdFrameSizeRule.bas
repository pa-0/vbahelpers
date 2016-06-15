Attribute VB_Name = "wWdFrameSizeRule"
Function WdFrameSizeRuleFromString(value As String) As WdFrameSizeRule
    If IsNumeric(value) Then
        WdFrameSizeRuleFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdFrameAuto": WdFrameSizeRuleFromString = wdFrameAuto
        Case "wdFrameAtLeast": WdFrameSizeRuleFromString = wdFrameAtLeast
        Case "wdFrameExact": WdFrameSizeRuleFromString = wdFrameExact
    End Select
End Function

Function WdFrameSizeRuleToString(value As WdFrameSizeRule) As String
    Select Case value
        Case wdFrameAuto: WdFrameSizeRuleToString = "wdFrameAuto"
        Case wdFrameAtLeast: WdFrameSizeRuleToString = "wdFrameAtLeast"
        Case wdFrameExact: WdFrameSizeRuleToString = "wdFrameExact"
    End Select
End Function
