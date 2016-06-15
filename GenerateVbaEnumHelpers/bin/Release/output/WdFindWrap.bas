Attribute VB_Name = "wWdFindWrap"
Function WdFindWrapFromString(value As String) As WdFindWrap
    If IsNumeric(value) Then
        WdFindWrapFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdFindStop": WdFindWrapFromString = wdFindStop
        Case "wdFindContinue": WdFindWrapFromString = wdFindContinue
        Case "wdFindAsk": WdFindWrapFromString = wdFindAsk
    End Select
End Function

Function WdFindWrapToString(value As WdFindWrap) As String
    Select Case value
        Case wdFindStop: WdFindWrapToString = "wdFindStop"
        Case wdFindContinue: WdFindWrapToString = "wdFindContinue"
        Case wdFindAsk: WdFindWrapToString = "wdFindAsk"
    End Select
End Function
