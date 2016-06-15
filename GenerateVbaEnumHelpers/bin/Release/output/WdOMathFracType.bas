Attribute VB_Name = "wWdOMathFracType"
Function WdOMathFracTypeFromString(value As String) As WdOMathFracType
    If IsNumeric(value) Then
        WdOMathFracTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdOMathFracBar": WdOMathFracTypeFromString = wdOMathFracBar
        Case "wdOMathFracNoBar": WdOMathFracTypeFromString = wdOMathFracNoBar
        Case "wdOMathFracSkw": WdOMathFracTypeFromString = wdOMathFracSkw
        Case "wdOMathFracLin": WdOMathFracTypeFromString = wdOMathFracLin
    End Select
End Function

Function WdOMathFracTypeToString(value As WdOMathFracType) As String
    Select Case value
        Case wdOMathFracBar: WdOMathFracTypeToString = "wdOMathFracBar"
        Case wdOMathFracNoBar: WdOMathFracTypeToString = "wdOMathFracNoBar"
        Case wdOMathFracSkw: WdOMathFracTypeToString = "wdOMathFracSkw"
        Case wdOMathFracLin: WdOMathFracTypeToString = "wdOMathFracLin"
    End Select
End Function
