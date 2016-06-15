Attribute VB_Name = "wSignatureProviderDetail"
Function SignatureProviderDetailFromString(value As String) As SignatureProviderDetail
    If IsNumeric(value) Then
        SignatureProviderDetailFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "sigprovdetUrl": SignatureProviderDetailFromString = sigprovdetUrl
        Case "sigprovdetHashAlgorithm": SignatureProviderDetailFromString = sigprovdetHashAlgorithm
        Case "sigprovdetUIOnly": SignatureProviderDetailFromString = sigprovdetUIOnly
        Case "sigprovdetUseOfficeUI": SignatureProviderDetailFromString = sigprovdetUseOfficeUI
        Case "sigprovdetUseOfficeStampUI": SignatureProviderDetailFromString = sigprovdetUseOfficeStampUI
    End Select
End Function

Function SignatureProviderDetailToString(value As SignatureProviderDetail) As String
    Select Case value
        Case sigprovdetUrl: SignatureProviderDetailToString = "sigprovdetUrl"
        Case sigprovdetHashAlgorithm: SignatureProviderDetailToString = "sigprovdetHashAlgorithm"
        Case sigprovdetUIOnly: SignatureProviderDetailToString = "sigprovdetUIOnly"
        Case sigprovdetUseOfficeUI: SignatureProviderDetailToString = "sigprovdetUseOfficeUI"
        Case sigprovdetUseOfficeStampUI: SignatureProviderDetailToString = "sigprovdetUseOfficeStampUI"
    End Select
End Function
