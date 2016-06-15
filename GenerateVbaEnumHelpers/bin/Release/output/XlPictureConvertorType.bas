Attribute VB_Name = "wXlPictureConvertorType"
Function XlPictureConvertorTypeFromString(value As String) As XlPictureConvertorType
    If IsNumeric(value) Then
        XlPictureConvertorTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlBMP": XlPictureConvertorTypeFromString = xlBMP
        Case "xlWMF": XlPictureConvertorTypeFromString = xlWMF
        Case "xlWPG": XlPictureConvertorTypeFromString = xlWPG
        Case "xlDRW": XlPictureConvertorTypeFromString = xlDRW
        Case "xlDXF": XlPictureConvertorTypeFromString = xlDXF
        Case "xlHGL": XlPictureConvertorTypeFromString = xlHGL
        Case "xlCGM": XlPictureConvertorTypeFromString = xlCGM
        Case "xlEPS": XlPictureConvertorTypeFromString = xlEPS
        Case "xlTIF": XlPictureConvertorTypeFromString = xlTIF
        Case "xlPCX": XlPictureConvertorTypeFromString = xlPCX
        Case "xlPIC": XlPictureConvertorTypeFromString = xlPIC
        Case "xlPLT": XlPictureConvertorTypeFromString = xlPLT
        Case "xlPCT": XlPictureConvertorTypeFromString = xlPCT
    End Select
End Function

Function XlPictureConvertorTypeToString(value As XlPictureConvertorType) As String
    Select Case value
        Case xlBMP: XlPictureConvertorTypeToString = "xlBMP"
        Case xlWMF: XlPictureConvertorTypeToString = "xlWMF"
        Case xlWPG: XlPictureConvertorTypeToString = "xlWPG"
        Case xlDRW: XlPictureConvertorTypeToString = "xlDRW"
        Case xlDXF: XlPictureConvertorTypeToString = "xlDXF"
        Case xlHGL: XlPictureConvertorTypeToString = "xlHGL"
        Case xlCGM: XlPictureConvertorTypeToString = "xlCGM"
        Case xlEPS: XlPictureConvertorTypeToString = "xlEPS"
        Case xlTIF: XlPictureConvertorTypeToString = "xlTIF"
        Case xlPCX: XlPictureConvertorTypeToString = "xlPCX"
        Case xlPIC: XlPictureConvertorTypeToString = "xlPIC"
        Case xlPLT: XlPictureConvertorTypeToString = "xlPLT"
        Case xlPCT: XlPictureConvertorTypeToString = "xlPCT"
    End Select
End Function
