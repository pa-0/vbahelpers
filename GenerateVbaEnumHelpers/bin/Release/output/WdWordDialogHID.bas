Attribute VB_Name = "wWdWordDialogHID"
Function WdWordDialogHIDFromString(value As String) As WdWordDialogHID
    If IsNumeric(value) Then
        WdWordDialogHIDFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "emptyenum": WdWordDialogHIDFromString = emptyenum
    End Select
End Function

Function WdWordDialogHIDToString(value As WdWordDialogHID) As String
    Select Case value
        Case emptyenum: WdWordDialogHIDToString = "emptyenum"
    End Select
End Function
