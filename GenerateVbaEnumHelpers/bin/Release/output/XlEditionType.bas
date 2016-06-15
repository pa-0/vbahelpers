Attribute VB_Name = "wXlEditionType"
Function XlEditionTypeFromString(value As String) As XlEditionType
    If IsNumeric(value) Then
        XlEditionTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlPublisher": XlEditionTypeFromString = xlPublisher
        Case "xlSubscriber": XlEditionTypeFromString = xlSubscriber
    End Select
End Function

Function XlEditionTypeToString(value As XlEditionType) As String
    Select Case value
        Case xlPublisher: XlEditionTypeToString = "xlPublisher"
        Case xlSubscriber: XlEditionTypeToString = "xlSubscriber"
    End Select
End Function
