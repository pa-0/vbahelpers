Attribute VB_Name = "wSignatureLineImage"
Function SignatureLineImageFromString(value As String) As SignatureLineImage
    If IsNumeric(value) Then
        SignatureLineImageFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "siglnimgSoftwareRequired": SignatureLineImageFromString = siglnimgSoftwareRequired
        Case "siglnimgUnsigned": SignatureLineImageFromString = siglnimgUnsigned
        Case "siglnimgSignedValid": SignatureLineImageFromString = siglnimgSignedValid
        Case "siglnimgSignedInvalid": SignatureLineImageFromString = siglnimgSignedInvalid
        Case "siglnimgSigned": SignatureLineImageFromString = siglnimgSigned
    End Select
End Function

Function SignatureLineImageToString(value As SignatureLineImage) As String
    Select Case value
        Case siglnimgSoftwareRequired: SignatureLineImageToString = "siglnimgSoftwareRequired"
        Case siglnimgUnsigned: SignatureLineImageToString = "siglnimgUnsigned"
        Case siglnimgSignedValid: SignatureLineImageToString = "siglnimgSignedValid"
        Case siglnimgSignedInvalid: SignatureLineImageToString = "siglnimgSignedInvalid"
        Case siglnimgSigned: SignatureLineImageToString = "siglnimgSigned"
    End Select
End Function
