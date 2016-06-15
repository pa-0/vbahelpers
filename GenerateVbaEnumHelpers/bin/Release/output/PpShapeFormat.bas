Attribute VB_Name = "wPpShapeFormat"
Function PpShapeFormatFromString(value As String) As PpShapeFormat
    If IsNumeric(value) Then
        PpShapeFormatFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "ppShapeFormatGIF": PpShapeFormatFromString = ppShapeFormatGIF
        Case "ppShapeFormatJPG": PpShapeFormatFromString = ppShapeFormatJPG
        Case "ppShapeFormatPNG": PpShapeFormatFromString = ppShapeFormatPNG
        Case "ppShapeFormatBMP": PpShapeFormatFromString = ppShapeFormatBMP
        Case "ppShapeFormatWMF": PpShapeFormatFromString = ppShapeFormatWMF
        Case "ppShapeFormatEMF": PpShapeFormatFromString = ppShapeFormatEMF
    End Select
End Function

Function PpShapeFormatToString(value As PpShapeFormat) As String
    Select Case value
        Case ppShapeFormatGIF: PpShapeFormatToString = "ppShapeFormatGIF"
        Case ppShapeFormatJPG: PpShapeFormatToString = "ppShapeFormatJPG"
        Case ppShapeFormatPNG: PpShapeFormatToString = "ppShapeFormatPNG"
        Case ppShapeFormatBMP: PpShapeFormatToString = "ppShapeFormatBMP"
        Case ppShapeFormatWMF: PpShapeFormatToString = "ppShapeFormatWMF"
        Case ppShapeFormatEMF: PpShapeFormatToString = "ppShapeFormatEMF"
    End Select
End Function
