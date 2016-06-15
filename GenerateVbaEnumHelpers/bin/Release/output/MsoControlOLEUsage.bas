Attribute VB_Name = "wMsoControlOLEUsage"
Function MsoControlOLEUsageFromString(value As String) As MsoControlOLEUsage
    If IsNumeric(value) Then
        MsoControlOLEUsageFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoControlOLEUsageNeither": MsoControlOLEUsageFromString = msoControlOLEUsageNeither
        Case "msoControlOLEUsageServer": MsoControlOLEUsageFromString = msoControlOLEUsageServer
        Case "msoControlOLEUsageClient": MsoControlOLEUsageFromString = msoControlOLEUsageClient
        Case "msoControlOLEUsageBoth": MsoControlOLEUsageFromString = msoControlOLEUsageBoth
    End Select
End Function

Function MsoControlOLEUsageToString(value As MsoControlOLEUsage) As String
    Select Case value
        Case msoControlOLEUsageNeither: MsoControlOLEUsageToString = "msoControlOLEUsageNeither"
        Case msoControlOLEUsageServer: MsoControlOLEUsageToString = "msoControlOLEUsageServer"
        Case msoControlOLEUsageClient: MsoControlOLEUsageToString = "msoControlOLEUsageClient"
        Case msoControlOLEUsageBoth: MsoControlOLEUsageToString = "msoControlOLEUsageBoth"
    End Select
End Function
