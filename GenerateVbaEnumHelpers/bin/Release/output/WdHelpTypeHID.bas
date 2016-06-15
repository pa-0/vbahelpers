Attribute VB_Name = "wWdHelpTypeHID"
Function WdHelpTypeHIDFromString(value As String) As WdHelpTypeHID
    If IsNumeric(value) Then
        WdHelpTypeHIDFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "emptyenum": WdHelpTypeHIDFromString = emptyenum
    End Select
End Function

Function WdHelpTypeHIDToString(value As WdHelpTypeHID) As String
    Select Case value
        Case emptyenum: WdHelpTypeHIDToString = "emptyenum"
    End Select
End Function
