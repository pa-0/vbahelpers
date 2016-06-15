Attribute VB_Name = "wWdCharacterCaseHID"
Function WdCharacterCaseHIDFromString(value As String) As WdCharacterCaseHID
    If IsNumeric(value) Then
        WdCharacterCaseHIDFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "emptyenum": WdCharacterCaseHIDFromString = emptyenum
    End Select
End Function

Function WdCharacterCaseHIDToString(value As WdCharacterCaseHID) As String
    Select Case value
        Case emptyenum: WdCharacterCaseHIDToString = "emptyenum"
    End Select
End Function
