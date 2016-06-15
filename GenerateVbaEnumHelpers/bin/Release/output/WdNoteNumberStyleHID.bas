Attribute VB_Name = "wWdNoteNumberStyleHID"
Function WdNoteNumberStyleHIDFromString(value As String) As WdNoteNumberStyleHID
    If IsNumeric(value) Then
        WdNoteNumberStyleHIDFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "emptyenum": WdNoteNumberStyleHIDFromString = emptyenum
    End Select
End Function

Function WdNoteNumberStyleHIDToString(value As WdNoteNumberStyleHID) As String
    Select Case value
        Case emptyenum: WdNoteNumberStyleHIDToString = "emptyenum"
    End Select
End Function
