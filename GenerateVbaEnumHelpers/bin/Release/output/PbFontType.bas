Attribute VB_Name = "wPbFontType"
Function PbFontTypeFromString(value As String) As PbFontType
    If IsNumeric(value) Then
        PbFontTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "pbFontTrueType": PbFontTypeFromString = pbFontTrueType
        Case "pbFontPrinter": PbFontTypeFromString = pbFontPrinter
        Case "pbFontRaster": PbFontTypeFromString = pbFontRaster
        Case "pbFontUnknown": PbFontTypeFromString = pbFontUnknown
    End Select
End Function

Function PbFontTypeToString(value As PbFontType) As String
    Select Case value
        Case pbFontTrueType: PbFontTypeToString = "pbFontTrueType"
        Case pbFontPrinter: PbFontTypeToString = "pbFontPrinter"
        Case pbFontRaster: PbFontTypeToString = "pbFontRaster"
        Case pbFontUnknown: PbFontTypeToString = "pbFontUnknown"
    End Select
End Function
