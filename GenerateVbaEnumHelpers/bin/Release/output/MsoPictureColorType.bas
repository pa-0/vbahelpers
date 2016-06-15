Attribute VB_Name = "wMsoPictureColorType"
Function MsoPictureColorTypeFromString(value As String) As MsoPictureColorType
    If IsNumeric(value) Then
        MsoPictureColorTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoPictureAutomatic": MsoPictureColorTypeFromString = msoPictureAutomatic
        Case "msoPictureGrayscale": MsoPictureColorTypeFromString = msoPictureGrayscale
        Case "msoPictureBlackAndWhite": MsoPictureColorTypeFromString = msoPictureBlackAndWhite
        Case "msoPictureWatermark": MsoPictureColorTypeFromString = msoPictureWatermark
        Case "msoPictureMixed": MsoPictureColorTypeFromString = msoPictureMixed
    End Select
End Function

Function MsoPictureColorTypeToString(value As MsoPictureColorType) As String
    Select Case value
        Case msoPictureAutomatic: MsoPictureColorTypeToString = "msoPictureAutomatic"
        Case msoPictureGrayscale: MsoPictureColorTypeToString = "msoPictureGrayscale"
        Case msoPictureBlackAndWhite: MsoPictureColorTypeToString = "msoPictureBlackAndWhite"
        Case msoPictureWatermark: MsoPictureColorTypeToString = "msoPictureWatermark"
        Case msoPictureMixed: MsoPictureColorTypeToString = "msoPictureMixed"
    End Select
End Function
