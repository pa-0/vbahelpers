Attribute VB_Name = "wXlMailSystem"
Function XlMailSystemFromString(value As String) As XlMailSystem
    If IsNumeric(value) Then
        XlMailSystemFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlNoMailSystem": XlMailSystemFromString = xlNoMailSystem
        Case "xlMAPI": XlMailSystemFromString = xlMAPI
        Case "xlPowerTalk": XlMailSystemFromString = xlPowerTalk
    End Select
End Function

Function XlMailSystemToString(value As XlMailSystem) As String
    Select Case value
        Case xlNoMailSystem: XlMailSystemToString = "xlNoMailSystem"
        Case xlMAPI: XlMailSystemToString = "xlMAPI"
        Case xlPowerTalk: XlMailSystemToString = "xlPowerTalk"
    End Select
End Function
