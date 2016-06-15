Attribute VB_Name = "wWdOMathBreakBin"
Function WdOMathBreakBinFromString(value As String) As WdOMathBreakBin
    If IsNumeric(value) Then
        WdOMathBreakBinFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdOMathBreakBinBefore": WdOMathBreakBinFromString = wdOMathBreakBinBefore
        Case "wdOMathBreakBinAfter": WdOMathBreakBinFromString = wdOMathBreakBinAfter
        Case "wdOMathBreakBinRepeat": WdOMathBreakBinFromString = wdOMathBreakBinRepeat
    End Select
End Function

Function WdOMathBreakBinToString(value As WdOMathBreakBin) As String
    Select Case value
        Case wdOMathBreakBinBefore: WdOMathBreakBinToString = "wdOMathBreakBinBefore"
        Case wdOMathBreakBinAfter: WdOMathBreakBinToString = "wdOMathBreakBinAfter"
        Case wdOMathBreakBinRepeat: WdOMathBreakBinToString = "wdOMathBreakBinRepeat"
    End Select
End Function
