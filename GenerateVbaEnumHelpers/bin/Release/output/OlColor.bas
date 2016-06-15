Attribute VB_Name = "wOlColor"
Function OlColorFromString(value As String) As OlColor
    If IsNumeric(value) Then
        OlColorFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "olAutoColor": OlColorFromString = olAutoColor
        Case "olColorBlack": OlColorFromString = olColorBlack
        Case "olColorMaroon": OlColorFromString = olColorMaroon
        Case "olColorGreen": OlColorFromString = olColorGreen
        Case "olColorOlive": OlColorFromString = olColorOlive
        Case "olColorNavy": OlColorFromString = olColorNavy
        Case "olColorPurple": OlColorFromString = olColorPurple
        Case "olColorTeal": OlColorFromString = olColorTeal
        Case "olColorGray": OlColorFromString = olColorGray
        Case "olColorSilver": OlColorFromString = olColorSilver
        Case "olColorRed": OlColorFromString = olColorRed
        Case "olColorLime": OlColorFromString = olColorLime
        Case "olColorYellow": OlColorFromString = olColorYellow
        Case "olColorBlue": OlColorFromString = olColorBlue
        Case "olColorFuchsia": OlColorFromString = olColorFuchsia
        Case "olColorAqua": OlColorFromString = olColorAqua
        Case "olColorWhite": OlColorFromString = olColorWhite
    End Select
End Function

Function OlColorToString(value As OlColor) As String
    Select Case value
        Case olAutoColor: OlColorToString = "olAutoColor"
        Case olColorBlack: OlColorToString = "olColorBlack"
        Case olColorMaroon: OlColorToString = "olColorMaroon"
        Case olColorGreen: OlColorToString = "olColorGreen"
        Case olColorOlive: OlColorToString = "olColorOlive"
        Case olColorNavy: OlColorToString = "olColorNavy"
        Case olColorPurple: OlColorToString = "olColorPurple"
        Case olColorTeal: OlColorToString = "olColorTeal"
        Case olColorGray: OlColorToString = "olColorGray"
        Case olColorSilver: OlColorToString = "olColorSilver"
        Case olColorRed: OlColorToString = "olColorRed"
        Case olColorLime: OlColorToString = "olColorLime"
        Case olColorYellow: OlColorToString = "olColorYellow"
        Case olColorBlue: OlColorToString = "olColorBlue"
        Case olColorFuchsia: OlColorToString = "olColorFuchsia"
        Case olColorAqua: OlColorToString = "olColorAqua"
        Case olColorWhite: OlColorToString = "olColorWhite"
    End Select
End Function
