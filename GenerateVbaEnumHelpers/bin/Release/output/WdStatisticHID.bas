Attribute VB_Name = "wWdStatisticHID"
Function WdStatisticHIDFromString(value As String) As WdStatisticHID
    If IsNumeric(value) Then
        WdStatisticHIDFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "emptyenum": WdStatisticHIDFromString = emptyenum
    End Select
End Function

Function WdStatisticHIDToString(value As WdStatisticHID) As String
    Select Case value
        Case emptyenum: WdStatisticHIDToString = "emptyenum"
    End Select
End Function
