Attribute VB_Name = "wWdDictionaryTypeHID"
Function WdDictionaryTypeHIDFromString(value As String) As WdDictionaryTypeHID
    If IsNumeric(value) Then
        WdDictionaryTypeHIDFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "emptyenum": WdDictionaryTypeHIDFromString = emptyenum
    End Select
End Function

Function WdDictionaryTypeHIDToString(value As WdDictionaryTypeHID) As String
    Select Case value
        Case emptyenum: WdDictionaryTypeHIDToString = "emptyenum"
    End Select
End Function
