Attribute VB_Name = "wWdContinue"
Function WdContinueFromString(value As String) As WdContinue
    If IsNumeric(value) Then
        WdContinueFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdContinueDisabled": WdContinueFromString = wdContinueDisabled
        Case "wdResetList": WdContinueFromString = wdResetList
        Case "wdContinueList": WdContinueFromString = wdContinueList
    End Select
End Function

Function WdContinueToString(value As WdContinue) As String
    Select Case value
        Case wdContinueDisabled: WdContinueToString = "wdContinueDisabled"
        Case wdResetList: WdContinueToString = "wdResetList"
        Case wdContinueList: WdContinueToString = "wdContinueList"
    End Select
End Function
