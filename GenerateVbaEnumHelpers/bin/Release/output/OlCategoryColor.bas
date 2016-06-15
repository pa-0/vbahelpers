Attribute VB_Name = "wOlCategoryColor"
Function OlCategoryColorFromString(value As String) As OlCategoryColor
    If IsNumeric(value) Then
        OlCategoryColorFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "olCategoryColorNone": OlCategoryColorFromString = olCategoryColorNone
        Case "olCategoryColorRed": OlCategoryColorFromString = olCategoryColorRed
        Case "olCategoryColorOrange": OlCategoryColorFromString = olCategoryColorOrange
        Case "olCategoryColorPeach": OlCategoryColorFromString = olCategoryColorPeach
        Case "olCategoryColorYellow": OlCategoryColorFromString = olCategoryColorYellow
        Case "olCategoryColorGreen": OlCategoryColorFromString = olCategoryColorGreen
        Case "olCategoryColorTeal": OlCategoryColorFromString = olCategoryColorTeal
        Case "olCategoryColorOlive": OlCategoryColorFromString = olCategoryColorOlive
        Case "olCategoryColorBlue": OlCategoryColorFromString = olCategoryColorBlue
        Case "olCategoryColorPurple": OlCategoryColorFromString = olCategoryColorPurple
        Case "olCategoryColorMaroon": OlCategoryColorFromString = olCategoryColorMaroon
        Case "olCategoryColorSteel": OlCategoryColorFromString = olCategoryColorSteel
        Case "olCategoryColorDarkSteel": OlCategoryColorFromString = olCategoryColorDarkSteel
        Case "olCategoryColorGray": OlCategoryColorFromString = olCategoryColorGray
        Case "olCategoryColorDarkGray": OlCategoryColorFromString = olCategoryColorDarkGray
        Case "olCategoryColorBlack": OlCategoryColorFromString = olCategoryColorBlack
        Case "olCategoryColorDarkRed": OlCategoryColorFromString = olCategoryColorDarkRed
        Case "olCategoryColorDarkOrange": OlCategoryColorFromString = olCategoryColorDarkOrange
        Case "olCategoryColorDarkPeach": OlCategoryColorFromString = olCategoryColorDarkPeach
        Case "olCategoryColorDarkYellow": OlCategoryColorFromString = olCategoryColorDarkYellow
        Case "olCategoryColorDarkGreen": OlCategoryColorFromString = olCategoryColorDarkGreen
        Case "olCategoryColorDarkTeal": OlCategoryColorFromString = olCategoryColorDarkTeal
        Case "olCategoryColorDarkOlive": OlCategoryColorFromString = olCategoryColorDarkOlive
        Case "olCategoryColorDarkBlue": OlCategoryColorFromString = olCategoryColorDarkBlue
        Case "olCategoryColorDarkPurple": OlCategoryColorFromString = olCategoryColorDarkPurple
        Case "olCategoryColorDarkMaroon": OlCategoryColorFromString = olCategoryColorDarkMaroon
    End Select
End Function

Function OlCategoryColorToString(value As OlCategoryColor) As String
    Select Case value
        Case olCategoryColorNone: OlCategoryColorToString = "olCategoryColorNone"
        Case olCategoryColorRed: OlCategoryColorToString = "olCategoryColorRed"
        Case olCategoryColorOrange: OlCategoryColorToString = "olCategoryColorOrange"
        Case olCategoryColorPeach: OlCategoryColorToString = "olCategoryColorPeach"
        Case olCategoryColorYellow: OlCategoryColorToString = "olCategoryColorYellow"
        Case olCategoryColorGreen: OlCategoryColorToString = "olCategoryColorGreen"
        Case olCategoryColorTeal: OlCategoryColorToString = "olCategoryColorTeal"
        Case olCategoryColorOlive: OlCategoryColorToString = "olCategoryColorOlive"
        Case olCategoryColorBlue: OlCategoryColorToString = "olCategoryColorBlue"
        Case olCategoryColorPurple: OlCategoryColorToString = "olCategoryColorPurple"
        Case olCategoryColorMaroon: OlCategoryColorToString = "olCategoryColorMaroon"
        Case olCategoryColorSteel: OlCategoryColorToString = "olCategoryColorSteel"
        Case olCategoryColorDarkSteel: OlCategoryColorToString = "olCategoryColorDarkSteel"
        Case olCategoryColorGray: OlCategoryColorToString = "olCategoryColorGray"
        Case olCategoryColorDarkGray: OlCategoryColorToString = "olCategoryColorDarkGray"
        Case olCategoryColorBlack: OlCategoryColorToString = "olCategoryColorBlack"
        Case olCategoryColorDarkRed: OlCategoryColorToString = "olCategoryColorDarkRed"
        Case olCategoryColorDarkOrange: OlCategoryColorToString = "olCategoryColorDarkOrange"
        Case olCategoryColorDarkPeach: OlCategoryColorToString = "olCategoryColorDarkPeach"
        Case olCategoryColorDarkYellow: OlCategoryColorToString = "olCategoryColorDarkYellow"
        Case olCategoryColorDarkGreen: OlCategoryColorToString = "olCategoryColorDarkGreen"
        Case olCategoryColorDarkTeal: OlCategoryColorToString = "olCategoryColorDarkTeal"
        Case olCategoryColorDarkOlive: OlCategoryColorToString = "olCategoryColorDarkOlive"
        Case olCategoryColorDarkBlue: OlCategoryColorToString = "olCategoryColorDarkBlue"
        Case olCategoryColorDarkPurple: OlCategoryColorToString = "olCategoryColorDarkPurple"
        Case olCategoryColorDarkMaroon: OlCategoryColorToString = "olCategoryColorDarkMaroon"
    End Select
End Function
