Attribute VB_Name = "wCertificateVerificationResults"
Function CertificateVerificationResultsFromString(value As String) As CertificateVerificationResults
    If IsNumeric(value) Then
        CertificateVerificationResultsFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "certverresError": CertificateVerificationResultsFromString = certverresError
        Case "certverresVerifying": CertificateVerificationResultsFromString = certverresVerifying
        Case "certverresUnverified": CertificateVerificationResultsFromString = certverresUnverified
        Case "certverresValid": CertificateVerificationResultsFromString = certverresValid
        Case "certverresInvalid": CertificateVerificationResultsFromString = certverresInvalid
        Case "certverresExpired": CertificateVerificationResultsFromString = certverresExpired
        Case "certverresRevoked": CertificateVerificationResultsFromString = certverresRevoked
        Case "certverresUntrusted": CertificateVerificationResultsFromString = certverresUntrusted
    End Select
End Function

Function CertificateVerificationResultsToString(value As CertificateVerificationResults) As String
    Select Case value
        Case certverresError: CertificateVerificationResultsToString = "certverresError"
        Case certverresVerifying: CertificateVerificationResultsToString = "certverresVerifying"
        Case certverresUnverified: CertificateVerificationResultsToString = "certverresUnverified"
        Case certverresValid: CertificateVerificationResultsToString = "certverresValid"
        Case certverresInvalid: CertificateVerificationResultsToString = "certverresInvalid"
        Case certverresExpired: CertificateVerificationResultsToString = "certverresExpired"
        Case certverresRevoked: CertificateVerificationResultsToString = "certverresRevoked"
        Case certverresUntrusted: CertificateVerificationResultsToString = "certverresUntrusted"
    End Select
End Function
