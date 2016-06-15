Attribute VB_Name = "wCertificateDetail"
Function CertificateDetailFromString(value As String) As CertificateDetail
    If IsNumeric(value) Then
        CertificateDetailFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "certdetAvailable": CertificateDetailFromString = certdetAvailable
        Case "certdetSubject": CertificateDetailFromString = certdetSubject
        Case "certdetIssuer": CertificateDetailFromString = certdetIssuer
        Case "certdetExpirationDate": CertificateDetailFromString = certdetExpirationDate
        Case "certdetThumbprint": CertificateDetailFromString = certdetThumbprint
    End Select
End Function

Function CertificateDetailToString(value As CertificateDetail) As String
    Select Case value
        Case certdetAvailable: CertificateDetailToString = "certdetAvailable"
        Case certdetSubject: CertificateDetailToString = "certdetSubject"
        Case certdetIssuer: CertificateDetailToString = "certdetIssuer"
        Case certdetExpirationDate: CertificateDetailToString = "certdetExpirationDate"
        Case certdetThumbprint: CertificateDetailToString = "certdetThumbprint"
    End Select
End Function
