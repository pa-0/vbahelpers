Attribute VB_Name = "wWdOMathBreakSub"
Function WdOMathBreakSubFromString(value As String) As WdOMathBreakSub
    If IsNumeric(value) Then
        WdOMathBreakSubFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdOMathBreakSubMinusMinus": WdOMathBreakSubFromString = wdOMathBreakSubMinusMinus
        Case "wdOMathBreakSubPlusMinus": WdOMathBreakSubFromString = wdOMathBreakSubPlusMinus
        Case "wdOMathBreakSubMinusPlus": WdOMathBreakSubFromString = wdOMathBreakSubMinusPlus
    End Select
End Function

Function WdOMathBreakSubToString(value As WdOMathBreakSub) As String
    Select Case value
        Case wdOMathBreakSubMinusMinus: WdOMathBreakSubToString = "wdOMathBreakSubMinusMinus"
        Case wdOMathBreakSubPlusMinus: WdOMathBreakSubToString = "wdOMathBreakSubPlusMinus"
        Case wdOMathBreakSubMinusPlus: WdOMathBreakSubToString = "wdOMathBreakSubMinusPlus"
    End Select
End Function
