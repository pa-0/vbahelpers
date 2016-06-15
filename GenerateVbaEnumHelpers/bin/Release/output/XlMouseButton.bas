Attribute VB_Name = "wXlMouseButton"
Function XlMouseButtonFromString(value As String) As XlMouseButton
    If IsNumeric(value) Then
        XlMouseButtonFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlNoButton": XlMouseButtonFromString = xlNoButton
        Case "xlPrimaryButton": XlMouseButtonFromString = xlPrimaryButton
        Case "xlSecondaryButton": XlMouseButtonFromString = xlSecondaryButton
    End Select
End Function

Function XlMouseButtonToString(value As XlMouseButton) As String
    Select Case value
        Case xlNoButton: XlMouseButtonToString = "xlNoButton"
        Case xlPrimaryButton: XlMouseButtonToString = "xlPrimaryButton"
        Case xlSecondaryButton: XlMouseButtonToString = "xlSecondaryButton"
    End Select
End Function
