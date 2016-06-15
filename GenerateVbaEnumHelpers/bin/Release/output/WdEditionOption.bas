Attribute VB_Name = "wWdEditionOption"
Function WdEditionOptionFromString(value As String) As WdEditionOption
    If IsNumeric(value) Then
        WdEditionOptionFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdCancelPublisher": WdEditionOptionFromString = wdCancelPublisher
        Case "wdSendPublisher": WdEditionOptionFromString = wdSendPublisher
        Case "wdSelectPublisher": WdEditionOptionFromString = wdSelectPublisher
        Case "wdAutomaticUpdate": WdEditionOptionFromString = wdAutomaticUpdate
        Case "wdManualUpdate": WdEditionOptionFromString = wdManualUpdate
        Case "wdChangeAttributes": WdEditionOptionFromString = wdChangeAttributes
        Case "wdUpdateSubscriber": WdEditionOptionFromString = wdUpdateSubscriber
        Case "wdOpenSource": WdEditionOptionFromString = wdOpenSource
    End Select
End Function

Function WdEditionOptionToString(value As WdEditionOption) As String
    Select Case value
        Case wdCancelPublisher: WdEditionOptionToString = "wdCancelPublisher"
        Case wdSendPublisher: WdEditionOptionToString = "wdSendPublisher"
        Case wdSelectPublisher: WdEditionOptionToString = "wdSelectPublisher"
        Case wdAutomaticUpdate: WdEditionOptionToString = "wdAutomaticUpdate"
        Case wdManualUpdate: WdEditionOptionToString = "wdManualUpdate"
        Case wdChangeAttributes: WdEditionOptionToString = "wdChangeAttributes"
        Case wdUpdateSubscriber: WdEditionOptionToString = "wdUpdateSubscriber"
        Case wdOpenSource: WdEditionOptionToString = "wdOpenSource"
    End Select
End Function
