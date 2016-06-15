Attribute VB_Name = "wPbImageFormat"
Function PbImageFormatFromString(value As String) As PbImageFormat
    If IsNumeric(value) Then
        PbImageFormatFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "pbImageFormatUNKNOWN": PbImageFormatFromString = pbImageFormatUNKNOWN
        Case "pbImageFormatEMF": PbImageFormatFromString = pbImageFormatEMF
        Case "pbImageFormatWMF": PbImageFormatFromString = pbImageFormatWMF
        Case "pbImageFormatPICT": PbImageFormatFromString = pbImageFormatPICT
        Case "pbImageFormatJPEG": PbImageFormatFromString = pbImageFormatJPEG
        Case "pbImageFormatPNG": PbImageFormatFromString = pbImageFormatPNG
        Case "pbImageFormatDIB": PbImageFormatFromString = pbImageFormatDIB
        Case "pbImageFormatGIF": PbImageFormatFromString = pbImageFormatGIF
        Case "pbImageFormatTIFF": PbImageFormatFromString = pbImageFormatTIFF
        Case "pbImageFormatCMYKJPEG": PbImageFormatFromString = pbImageFormatCMYKJPEG
    End Select
End Function

Function PbImageFormatToString(value As PbImageFormat) As String
    Select Case value
        Case pbImageFormatUNKNOWN: PbImageFormatToString = "pbImageFormatUNKNOWN"
        Case pbImageFormatEMF: PbImageFormatToString = "pbImageFormatEMF"
        Case pbImageFormatWMF: PbImageFormatToString = "pbImageFormatWMF"
        Case pbImageFormatPICT: PbImageFormatToString = "pbImageFormatPICT"
        Case pbImageFormatJPEG: PbImageFormatToString = "pbImageFormatJPEG"
        Case pbImageFormatPNG: PbImageFormatToString = "pbImageFormatPNG"
        Case pbImageFormatDIB: PbImageFormatToString = "pbImageFormatDIB"
        Case pbImageFormatGIF: PbImageFormatToString = "pbImageFormatGIF"
        Case pbImageFormatTIFF: PbImageFormatToString = "pbImageFormatTIFF"
        Case pbImageFormatCMYKJPEG: PbImageFormatToString = "pbImageFormatCMYKJPEG"
    End Select
End Function
