Attribute VB_Name = "wWdEnableCancelKey"
Function WdEnableCancelKeyFromString(value As String) As WdEnableCancelKey
    If IsNumeric(value) Then
        WdEnableCancelKeyFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdCancelDisabled": WdEnableCancelKeyFromString = wdCancelDisabled
        Case "wdCancelInterrupt": WdEnableCancelKeyFromString = wdCancelInterrupt
    End Select
End Function

Function WdEnableCancelKeyToString(value As WdEnableCancelKey) As String
    Select Case value
        Case wdCancelDisabled: WdEnableCancelKeyToString = "wdCancelDisabled"
        Case wdCancelInterrupt: WdEnableCancelKeyToString = "wdCancelInterrupt"
    End Select
End Function
