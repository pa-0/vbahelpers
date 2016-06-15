Attribute VB_Name = "wPpFrameColors"
Function PpFrameColorsFromString(value As String) As PpFrameColors
    If IsNumeric(value) Then
        PpFrameColorsFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "ppFrameColorsBrowserColors": PpFrameColorsFromString = ppFrameColorsBrowserColors
        Case "ppFrameColorsPresentationSchemeTextColor": PpFrameColorsFromString = ppFrameColorsPresentationSchemeTextColor
        Case "ppFrameColorsPresentationSchemeAccentColor": PpFrameColorsFromString = ppFrameColorsPresentationSchemeAccentColor
        Case "ppFrameColorsWhiteTextOnBlack": PpFrameColorsFromString = ppFrameColorsWhiteTextOnBlack
        Case "ppFrameColorsBlackTextOnWhite": PpFrameColorsFromString = ppFrameColorsBlackTextOnWhite
    End Select
End Function

Function PpFrameColorsToString(value As PpFrameColors) As String
    Select Case value
        Case ppFrameColorsBrowserColors: PpFrameColorsToString = "ppFrameColorsBrowserColors"
        Case ppFrameColorsPresentationSchemeTextColor: PpFrameColorsToString = "ppFrameColorsPresentationSchemeTextColor"
        Case ppFrameColorsPresentationSchemeAccentColor: PpFrameColorsToString = "ppFrameColorsPresentationSchemeAccentColor"
        Case ppFrameColorsWhiteTextOnBlack: PpFrameColorsToString = "ppFrameColorsWhiteTextOnBlack"
        Case ppFrameColorsBlackTextOnWhite: PpFrameColorsToString = "ppFrameColorsBlackTextOnWhite"
    End Select
End Function
