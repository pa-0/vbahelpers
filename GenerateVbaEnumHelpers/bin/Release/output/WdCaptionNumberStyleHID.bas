Attribute VB_Name = "wWdCaptionNumberStyleHID"
Function WdCaptionNumberStyleHIDFromString(value As String) As WdCaptionNumberStyleHID
    If IsNumeric(value) Then
        WdCaptionNumberStyleHIDFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "emptyenum": WdCaptionNumberStyleHIDFromString = emptyenum
    End Select
End Function

Function WdCaptionNumberStyleHIDToString(value As WdCaptionNumberStyleHID) As String
    Select Case value
        Case emptyenum: WdCaptionNumberStyleHIDToString = "emptyenum"
    End Select
End Function
