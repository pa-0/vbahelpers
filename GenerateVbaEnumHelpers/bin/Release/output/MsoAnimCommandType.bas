Attribute VB_Name = "wMsoAnimCommandType"
Function MsoAnimCommandTypeFromString(value As String) As MsoAnimCommandType
    If IsNumeric(value) Then
        MsoAnimCommandTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoAnimCommandTypeEvent": MsoAnimCommandTypeFromString = msoAnimCommandTypeEvent
        Case "msoAnimCommandTypeCall": MsoAnimCommandTypeFromString = msoAnimCommandTypeCall
        Case "msoAnimCommandTypeVerb": MsoAnimCommandTypeFromString = msoAnimCommandTypeVerb
    End Select
End Function

Function MsoAnimCommandTypeToString(value As MsoAnimCommandType) As String
    Select Case value
        Case msoAnimCommandTypeEvent: MsoAnimCommandTypeToString = "msoAnimCommandTypeEvent"
        Case msoAnimCommandTypeCall: MsoAnimCommandTypeToString = "msoAnimCommandTypeCall"
        Case msoAnimCommandTypeVerb: MsoAnimCommandTypeToString = "msoAnimCommandTypeVerb"
    End Select
End Function
