Attribute VB_Name = "wWdTextOrientationHID"
Function WdTextOrientationHIDFromString(value As String) As WdTextOrientationHID
    If IsNumeric(value) Then
        WdTextOrientationHIDFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "emptyenum": WdTextOrientationHIDFromString = emptyenum
    End Select
End Function

Function WdTextOrientationHIDToString(value As WdTextOrientationHID) As String
    Select Case value
        Case emptyenum: WdTextOrientationHIDToString = "emptyenum"
    End Select
End Function
