Attribute VB_Name = "wWdColor"
Function WdColorFromString(value As String) As WdColor
    If IsNumeric(value) Then
        WdColorFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdColorBlack": WdColorFromString = wdColorBlack
        Case "wdColorDarkRed": WdColorFromString = wdColorDarkRed
        Case "wdColorRed": WdColorFromString = wdColorRed
        Case "wdColorDarkGreen": WdColorFromString = wdColorDarkGreen
        Case "wdColorOliveGreen": WdColorFromString = wdColorOliveGreen
        Case "wdColorBrown": WdColorFromString = wdColorBrown
        Case "wdColorOrange": WdColorFromString = wdColorOrange
        Case "wdColorGreen": WdColorFromString = wdColorGreen
        Case "wdColorDarkYellow": WdColorFromString = wdColorDarkYellow
        Case "wdColorLightOrange": WdColorFromString = wdColorLightOrange
        Case "wdColorLime": WdColorFromString = wdColorLime
        Case "wdColorGold": WdColorFromString = wdColorGold
        Case "wdColorBrightGreen": WdColorFromString = wdColorBrightGreen
        Case "wdColorYellow": WdColorFromString = wdColorYellow
        Case "wdColorGray95": WdColorFromString = wdColorGray95
        Case "wdColorGray90": WdColorFromString = wdColorGray90
        Case "wdColorGray875": WdColorFromString = wdColorGray875
        Case "wdColorGray85": WdColorFromString = wdColorGray85
        Case "wdColorGray80": WdColorFromString = wdColorGray80
        Case "wdColorGray75": WdColorFromString = wdColorGray75
        Case "wdColorGray70": WdColorFromString = wdColorGray70
        Case "wdColorGray65": WdColorFromString = wdColorGray65
        Case "wdColorGray625": WdColorFromString = wdColorGray625
        Case "wdColorDarkTeal": WdColorFromString = wdColorDarkTeal
        Case "wdColorPlum": WdColorFromString = wdColorPlum
        Case "wdColorGray60": WdColorFromString = wdColorGray60
        Case "wdColorSeaGreen": WdColorFromString = wdColorSeaGreen
        Case "wdColorGray55": WdColorFromString = wdColorGray55
        Case "wdColorDarkBlue": WdColorFromString = wdColorDarkBlue
        Case "wdColorViolet": WdColorFromString = wdColorViolet
        Case "wdColorTeal": WdColorFromString = wdColorTeal
        Case "wdColorGray50": WdColorFromString = wdColorGray50
        Case "wdColorGray45": WdColorFromString = wdColorGray45
        Case "wdColorIndigo": WdColorFromString = wdColorIndigo
        Case "wdColorBlueGray": WdColorFromString = wdColorBlueGray
        Case "wdColorGray40": WdColorFromString = wdColorGray40
        Case "wdColorTan": WdColorFromString = wdColorTan
        Case "wdColorLightYellow": WdColorFromString = wdColorLightYellow
        Case "wdColorGray375": WdColorFromString = wdColorGray375
        Case "wdColorGray35": WdColorFromString = wdColorGray35
        Case "wdColorGray30": WdColorFromString = wdColorGray30
        Case "wdColorGray25": WdColorFromString = wdColorGray25
        Case "wdColorRose": WdColorFromString = wdColorRose
        Case "wdColorAqua": WdColorFromString = wdColorAqua
        Case "wdColorGray20": WdColorFromString = wdColorGray20
        Case "wdColorLightGreen": WdColorFromString = wdColorLightGreen
        Case "wdColorGray15": WdColorFromString = wdColorGray15
        Case "wdColorGray125": WdColorFromString = wdColorGray125
        Case "wdColorGray10": WdColorFromString = wdColorGray10
        Case "wdColorGray05": WdColorFromString = wdColorGray05
        Case "wdColorBlue": WdColorFromString = wdColorBlue
        Case "wdColorPink": WdColorFromString = wdColorPink
        Case "wdColorLightBlue": WdColorFromString = wdColorLightBlue
        Case "wdColorLavender": WdColorFromString = wdColorLavender
        Case "wdColorSkyBlue": WdColorFromString = wdColorSkyBlue
        Case "wdColorPaleBlue": WdColorFromString = wdColorPaleBlue
        Case "wdColorTurquoise": WdColorFromString = wdColorTurquoise
        Case "wdColorLightTurquoise": WdColorFromString = wdColorLightTurquoise
        Case "wdColorWhite": WdColorFromString = wdColorWhite
        Case "wdColorAutomatic": WdColorFromString = wdColorAutomatic
    End Select
End Function

Function WdColorToString(value As WdColor) As String
    Select Case value
        Case wdColorBlack: WdColorToString = "wdColorBlack"
        Case wdColorDarkRed: WdColorToString = "wdColorDarkRed"
        Case wdColorRed: WdColorToString = "wdColorRed"
        Case wdColorDarkGreen: WdColorToString = "wdColorDarkGreen"
        Case wdColorOliveGreen: WdColorToString = "wdColorOliveGreen"
        Case wdColorBrown: WdColorToString = "wdColorBrown"
        Case wdColorOrange: WdColorToString = "wdColorOrange"
        Case wdColorGreen: WdColorToString = "wdColorGreen"
        Case wdColorDarkYellow: WdColorToString = "wdColorDarkYellow"
        Case wdColorLightOrange: WdColorToString = "wdColorLightOrange"
        Case wdColorLime: WdColorToString = "wdColorLime"
        Case wdColorGold: WdColorToString = "wdColorGold"
        Case wdColorBrightGreen: WdColorToString = "wdColorBrightGreen"
        Case wdColorYellow: WdColorToString = "wdColorYellow"
        Case wdColorGray95: WdColorToString = "wdColorGray95"
        Case wdColorGray90: WdColorToString = "wdColorGray90"
        Case wdColorGray875: WdColorToString = "wdColorGray875"
        Case wdColorGray85: WdColorToString = "wdColorGray85"
        Case wdColorGray80: WdColorToString = "wdColorGray80"
        Case wdColorGray75: WdColorToString = "wdColorGray75"
        Case wdColorGray70: WdColorToString = "wdColorGray70"
        Case wdColorGray65: WdColorToString = "wdColorGray65"
        Case wdColorGray625: WdColorToString = "wdColorGray625"
        Case wdColorDarkTeal: WdColorToString = "wdColorDarkTeal"
        Case wdColorPlum: WdColorToString = "wdColorPlum"
        Case wdColorGray60: WdColorToString = "wdColorGray60"
        Case wdColorSeaGreen: WdColorToString = "wdColorSeaGreen"
        Case wdColorGray55: WdColorToString = "wdColorGray55"
        Case wdColorDarkBlue: WdColorToString = "wdColorDarkBlue"
        Case wdColorViolet: WdColorToString = "wdColorViolet"
        Case wdColorTeal: WdColorToString = "wdColorTeal"
        Case wdColorGray50: WdColorToString = "wdColorGray50"
        Case wdColorGray45: WdColorToString = "wdColorGray45"
        Case wdColorIndigo: WdColorToString = "wdColorIndigo"
        Case wdColorBlueGray: WdColorToString = "wdColorBlueGray"
        Case wdColorGray40: WdColorToString = "wdColorGray40"
        Case wdColorTan: WdColorToString = "wdColorTan"
        Case wdColorLightYellow: WdColorToString = "wdColorLightYellow"
        Case wdColorGray375: WdColorToString = "wdColorGray375"
        Case wdColorGray35: WdColorToString = "wdColorGray35"
        Case wdColorGray30: WdColorToString = "wdColorGray30"
        Case wdColorGray25: WdColorToString = "wdColorGray25"
        Case wdColorRose: WdColorToString = "wdColorRose"
        Case wdColorAqua: WdColorToString = "wdColorAqua"
        Case wdColorGray20: WdColorToString = "wdColorGray20"
        Case wdColorLightGreen: WdColorToString = "wdColorLightGreen"
        Case wdColorGray15: WdColorToString = "wdColorGray15"
        Case wdColorGray125: WdColorToString = "wdColorGray125"
        Case wdColorGray10: WdColorToString = "wdColorGray10"
        Case wdColorGray05: WdColorToString = "wdColorGray05"
        Case wdColorBlue: WdColorToString = "wdColorBlue"
        Case wdColorPink: WdColorToString = "wdColorPink"
        Case wdColorLightBlue: WdColorToString = "wdColorLightBlue"
        Case wdColorLavender: WdColorToString = "wdColorLavender"
        Case wdColorSkyBlue: WdColorToString = "wdColorSkyBlue"
        Case wdColorPaleBlue: WdColorToString = "wdColorPaleBlue"
        Case wdColorTurquoise: WdColorToString = "wdColorTurquoise"
        Case wdColorLightTurquoise: WdColorToString = "wdColorLightTurquoise"
        Case wdColorWhite: WdColorToString = "wdColorWhite"
        Case wdColorAutomatic: WdColorToString = "wdColorAutomatic"
    End Select
End Function
