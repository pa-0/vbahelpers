Attribute VB_Name = "wEncryptionCipherMode"
Function EncryptionCipherModeFromString(value As String) As EncryptionCipherMode
    If IsNumeric(value) Then
        EncryptionCipherModeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "cipherModeECB": EncryptionCipherModeFromString = cipherModeECB
        Case "cipherModeCBC": EncryptionCipherModeFromString = cipherModeCBC
    End Select
End Function

Function EncryptionCipherModeToString(value As EncryptionCipherMode) As String
    Select Case value
        Case cipherModeECB: EncryptionCipherModeToString = "cipherModeECB"
        Case cipherModeCBC: EncryptionCipherModeToString = "cipherModeCBC"
    End Select
End Function
