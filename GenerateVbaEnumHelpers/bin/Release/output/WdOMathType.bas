Attribute VB_Name = "wWdOMathType"
Function WdOMathTypeFromString(value As String) As WdOMathType
    If IsNumeric(value) Then
        WdOMathTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdOMathDisplay": WdOMathTypeFromString = wdOMathDisplay
        Case "wdOMathInline": WdOMathTypeFromString = wdOMathInline
    End Select
End Function

Function WdOMathTypeToString(value As WdOMathType) As String
    Select Case value
        Case wdOMathDisplay: WdOMathTypeToString = "wdOMathDisplay"
        Case wdOMathInline: WdOMathTypeToString = "wdOMathInline"
    End Select
End Function
