Attribute VB_Name = "wPbFontLicenseLimitations"
Function PbFontLicenseLimitationsFromString(value As String) As PbFontLicenseLimitations
    If IsNumeric(value) Then
        PbFontLicenseLimitationsFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "pbFontEmbeddable": PbFontLicenseLimitationsFromString = pbFontEmbeddable
        Case "pbFontPrintPreviewEmbeddable": PbFontLicenseLimitationsFromString = pbFontPrintPreviewEmbeddable
        Case "pbFontNotEmbeddable": PbFontLicenseLimitationsFromString = pbFontNotEmbeddable
    End Select
End Function

Function PbFontLicenseLimitationsToString(value As PbFontLicenseLimitations) As String
    Select Case value
        Case pbFontEmbeddable: PbFontLicenseLimitationsToString = "pbFontEmbeddable"
        Case pbFontPrintPreviewEmbeddable: PbFontLicenseLimitationsToString = "pbFontPrintPreviewEmbeddable"
        Case pbFontNotEmbeddable: PbFontLicenseLimitationsToString = "pbFontNotEmbeddable"
    End Select
End Function
