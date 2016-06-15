Attribute VB_Name = "wWdSortFieldTypeHID"
Function WdSortFieldTypeHIDFromString(value As String) As WdSortFieldTypeHID
    If IsNumeric(value) Then
        WdSortFieldTypeHIDFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "emptyenum": WdSortFieldTypeHIDFromString = emptyenum
    End Select
End Function

Function WdSortFieldTypeHIDToString(value As WdSortFieldTypeHID) As String
    Select Case value
        Case emptyenum: WdSortFieldTypeHIDToString = "emptyenum"
    End Select
End Function
