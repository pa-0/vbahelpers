Attribute VB_Name = "wWdPageNumberStyleHID"
Function WdPageNumberStyleHIDFromString(value As String) As WdPageNumberStyleHID
    If IsNumeric(value) Then
        WdPageNumberStyleHIDFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "emptyenum": WdPageNumberStyleHIDFromString = emptyenum
    End Select
End Function

Function WdPageNumberStyleHIDToString(value As WdPageNumberStyleHID) As String
    Select Case value
        Case emptyenum: WdPageNumberStyleHIDToString = "emptyenum"
    End Select
End Function
