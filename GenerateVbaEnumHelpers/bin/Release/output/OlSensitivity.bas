Attribute VB_Name = "wOlSensitivity"
Function OlSensitivityFromString(value As String) As OlSensitivity
    If IsNumeric(value) Then
        OlSensitivityFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "olNormal": OlSensitivityFromString = olNormal
        Case "olPersonal": OlSensitivityFromString = olPersonal
        Case "olPrivate": OlSensitivityFromString = olPrivate
        Case "olConfidential": OlSensitivityFromString = olConfidential
    End Select
End Function

Function OlSensitivityToString(value As OlSensitivity) As String
    Select Case value
        Case olNormal: OlSensitivityToString = "olNormal"
        Case olPersonal: OlSensitivityToString = "olPersonal"
        Case olPrivate: OlSensitivityToString = "olPrivate"
        Case olConfidential: OlSensitivityToString = "olConfidential"
    End Select
End Function
