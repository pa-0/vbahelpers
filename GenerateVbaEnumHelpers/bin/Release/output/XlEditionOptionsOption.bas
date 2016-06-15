Attribute VB_Name = "wXlEditionOptionsOption"
Function XlEditionOptionsOptionFromString(value As String) As XlEditionOptionsOption
    If IsNumeric(value) Then
        XlEditionOptionsOptionFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlCancel": XlEditionOptionsOptionFromString = xlCancel
        Case "xlUpdateSubscriber": XlEditionOptionsOptionFromString = xlUpdateSubscriber
        Case "xlSendPublisher": XlEditionOptionsOptionFromString = xlSendPublisher
        Case "xlOpenSource": XlEditionOptionsOptionFromString = xlOpenSource
        Case "xlSelect": XlEditionOptionsOptionFromString = xlSelect
        Case "xlAutomaticUpdate": XlEditionOptionsOptionFromString = xlAutomaticUpdate
        Case "xlManualUpdate": XlEditionOptionsOptionFromString = xlManualUpdate
        Case "xlChangeAttributes": XlEditionOptionsOptionFromString = xlChangeAttributes
    End Select
End Function

Function XlEditionOptionsOptionToString(value As XlEditionOptionsOption) As String
    Select Case value
        Case xlCancel: XlEditionOptionsOptionToString = "xlCancel"
        Case xlUpdateSubscriber: XlEditionOptionsOptionToString = "xlUpdateSubscriber"
        Case xlSendPublisher: XlEditionOptionsOptionToString = "xlSendPublisher"
        Case xlOpenSource: XlEditionOptionsOptionToString = "xlOpenSource"
        Case xlSelect: XlEditionOptionsOptionToString = "xlSelect"
        Case xlAutomaticUpdate: XlEditionOptionsOptionToString = "xlAutomaticUpdate"
        Case xlManualUpdate: XlEditionOptionsOptionToString = "xlManualUpdate"
        Case xlChangeAttributes: XlEditionOptionsOptionToString = "xlChangeAttributes"
    End Select
End Function
