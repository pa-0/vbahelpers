Attribute VB_Name = "wWdListNumberStyleHID"
Function WdListNumberStyleHIDFromString(value As String) As WdListNumberStyleHID
    If IsNumeric(value) Then
        WdListNumberStyleHIDFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "emptyenum": WdListNumberStyleHIDFromString = emptyenum
    End Select
End Function

Function WdListNumberStyleHIDToString(value As WdListNumberStyleHID) As String
    Select Case value
        Case emptyenum: WdListNumberStyleHIDToString = "emptyenum"
    End Select
End Function
