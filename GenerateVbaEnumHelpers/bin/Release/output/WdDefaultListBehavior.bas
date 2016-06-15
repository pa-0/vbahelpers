Attribute VB_Name = "wWdDefaultListBehavior"
Function WdDefaultListBehaviorFromString(value As String) As WdDefaultListBehavior
    If IsNumeric(value) Then
        WdDefaultListBehaviorFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdWord8ListBehavior": WdDefaultListBehaviorFromString = wdWord8ListBehavior
        Case "wdWord9ListBehavior": WdDefaultListBehaviorFromString = wdWord9ListBehavior
        Case "wdWord10ListBehavior": WdDefaultListBehaviorFromString = wdWord10ListBehavior
    End Select
End Function

Function WdDefaultListBehaviorToString(value As WdDefaultListBehavior) As String
    Select Case value
        Case wdWord8ListBehavior: WdDefaultListBehaviorToString = "wdWord8ListBehavior"
        Case wdWord9ListBehavior: WdDefaultListBehaviorToString = "wdWord9ListBehavior"
        Case wdWord10ListBehavior: WdDefaultListBehaviorToString = "wdWord10ListBehavior"
    End Select
End Function
