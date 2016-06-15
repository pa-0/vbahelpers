Attribute VB_Name = "wWdDefaultTableBehavior"
Function WdDefaultTableBehaviorFromString(value As String) As WdDefaultTableBehavior
    If IsNumeric(value) Then
        WdDefaultTableBehaviorFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdWord8TableBehavior": WdDefaultTableBehaviorFromString = wdWord8TableBehavior
        Case "wdWord9TableBehavior": WdDefaultTableBehaviorFromString = wdWord9TableBehavior
    End Select
End Function

Function WdDefaultTableBehaviorToString(value As WdDefaultTableBehavior) As String
    Select Case value
        Case wdWord8TableBehavior: WdDefaultTableBehaviorToString = "wdWord8TableBehavior"
        Case wdWord9TableBehavior: WdDefaultTableBehaviorToString = "wdWord9TableBehavior"
    End Select
End Function
