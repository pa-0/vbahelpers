Attribute VB_Name = "wWdListApplyTo"
Function WdListApplyToFromString(value As String) As WdListApplyTo
    If IsNumeric(value) Then
        WdListApplyToFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdListApplyToWholeList": WdListApplyToFromString = wdListApplyToWholeList
        Case "wdListApplyToThisPointForward": WdListApplyToFromString = wdListApplyToThisPointForward
        Case "wdListApplyToSelection": WdListApplyToFromString = wdListApplyToSelection
    End Select
End Function

Function WdListApplyToToString(value As WdListApplyTo) As String
    Select Case value
        Case wdListApplyToWholeList: WdListApplyToToString = "wdListApplyToWholeList"
        Case wdListApplyToThisPointForward: WdListApplyToToString = "wdListApplyToThisPointForward"
        Case wdListApplyToSelection: WdListApplyToToString = "wdListApplyToSelection"
    End Select
End Function
