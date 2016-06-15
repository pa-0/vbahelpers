Attribute VB_Name = "wWdBorderTypeHID"
Function WdBorderTypeHIDFromString(value As String) As WdBorderTypeHID
    If IsNumeric(value) Then
        WdBorderTypeHIDFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "emptyenum": WdBorderTypeHIDFromString = emptyenum
    End Select
End Function

Function WdBorderTypeHIDToString(value As WdBorderTypeHID) As String
    Select Case value
        Case emptyenum: WdBorderTypeHIDToString = "emptyenum"
    End Select
End Function
