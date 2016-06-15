Attribute VB_Name = "wPbPictureResolution"
Function PbPictureResolutionFromString(value As String) As PbPictureResolution
    If IsNumeric(value) Then
        PbPictureResolutionFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "pbPictureResolutionDefault": PbPictureResolutionFromString = pbPictureResolutionDefault
        Case "pbPictureResolutionWeb_96dpi": PbPictureResolutionFromString = pbPictureResolutionWeb_96dpi
        Case "pbPictureResolutionDesktopPrint_150dpi": PbPictureResolutionFromString = pbPictureResolutionDesktopPrint_150dpi
        Case "pbPictureResolutionCommercialPrint_300dpi": PbPictureResolutionFromString = pbPictureResolutionCommercialPrint_300dpi
    End Select
End Function

Function PbPictureResolutionToString(value As PbPictureResolution) As String
    Select Case value
        Case pbPictureResolutionDefault: PbPictureResolutionToString = "pbPictureResolutionDefault"
        Case pbPictureResolutionWeb_96dpi: PbPictureResolutionToString = "pbPictureResolutionWeb_96dpi"
        Case pbPictureResolutionDesktopPrint_150dpi: PbPictureResolutionToString = "pbPictureResolutionDesktopPrint_150dpi"
        Case pbPictureResolutionCommercialPrint_300dpi: PbPictureResolutionToString = "pbPictureResolutionCommercialPrint_300dpi"
    End Select
End Function
