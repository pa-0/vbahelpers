Attribute VB_Name = "wXlBackground"
Function XlBackgroundFromString(value As String) As XlBackground
    If IsNumeric(value) Then
        XlBackgroundFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlBackgroundTransparent": XlBackgroundFromString = xlBackgroundTransparent
        Case "xlBackgroundOpaque": XlBackgroundFromString = xlBackgroundOpaque
        Case "xlBackgroundAutomatic": XlBackgroundFromString = xlBackgroundAutomatic
    End Select
End Function

Function XlBackgroundToString(value As XlBackground) As String
    Select Case value
        Case xlBackgroundTransparent: XlBackgroundToString = "xlBackgroundTransparent"
        Case xlBackgroundOpaque: XlBackgroundToString = "xlBackgroundOpaque"
        Case xlBackgroundAutomatic: XlBackgroundToString = "xlBackgroundAutomatic"
    End Select
End Function
