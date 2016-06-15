Attribute VB_Name = "wWdEditionType"
Function WdEditionTypeFromString(value As String) As WdEditionType
    If IsNumeric(value) Then
        WdEditionTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdPublisher": WdEditionTypeFromString = wdPublisher
        Case "wdSubscriber": WdEditionTypeFromString = wdSubscriber
    End Select
End Function

Function WdEditionTypeToString(value As WdEditionType) As String
    Select Case value
        Case wdPublisher: WdEditionTypeToString = "wdPublisher"
        Case wdSubscriber: WdEditionTypeToString = "wdSubscriber"
    End Select
End Function
