Attribute VB_Name = "wEncryptionProviderDetail"
Function EncryptionProviderDetailFromString(value As String) As EncryptionProviderDetail
    If IsNumeric(value) Then
        EncryptionProviderDetailFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "encprovdetUrl": EncryptionProviderDetailFromString = encprovdetUrl
        Case "encprovdetAlgorithm": EncryptionProviderDetailFromString = encprovdetAlgorithm
        Case "encprovdetBlockCipher": EncryptionProviderDetailFromString = encprovdetBlockCipher
        Case "encprovdetCipherBlockSize": EncryptionProviderDetailFromString = encprovdetCipherBlockSize
        Case "encprovdetCipherMode": EncryptionProviderDetailFromString = encprovdetCipherMode
    End Select
End Function

Function EncryptionProviderDetailToString(value As EncryptionProviderDetail) As String
    Select Case value
        Case encprovdetUrl: EncryptionProviderDetailToString = "encprovdetUrl"
        Case encprovdetAlgorithm: EncryptionProviderDetailToString = "encprovdetAlgorithm"
        Case encprovdetBlockCipher: EncryptionProviderDetailToString = "encprovdetBlockCipher"
        Case encprovdetCipherBlockSize: EncryptionProviderDetailToString = "encprovdetCipherBlockSize"
        Case encprovdetCipherMode: EncryptionProviderDetailToString = "encprovdetCipherMode"
    End Select
End Function
