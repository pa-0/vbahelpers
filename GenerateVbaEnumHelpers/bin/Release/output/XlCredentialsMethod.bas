Attribute VB_Name = "wXlCredentialsMethod"
Function XlCredentialsMethodFromString(value As String) As XlCredentialsMethod
    If IsNumeric(value) Then
        XlCredentialsMethodFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlCredentialsMethodIntegrated": XlCredentialsMethodFromString = xlCredentialsMethodIntegrated
        Case "xlCredentialsMethodNone": XlCredentialsMethodFromString = xlCredentialsMethodNone
        Case "xlCredentialsMethodStored": XlCredentialsMethodFromString = xlCredentialsMethodStored
    End Select
End Function

Function XlCredentialsMethodToString(value As XlCredentialsMethod) As String
    Select Case value
        Case xlCredentialsMethodIntegrated: XlCredentialsMethodToString = "xlCredentialsMethodIntegrated"
        Case xlCredentialsMethodNone: XlCredentialsMethodToString = "xlCredentialsMethodNone"
        Case xlCredentialsMethodStored: XlCredentialsMethodToString = "xlCredentialsMethodStored"
    End Select
End Function
