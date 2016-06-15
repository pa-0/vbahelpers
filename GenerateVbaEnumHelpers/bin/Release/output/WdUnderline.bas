Attribute VB_Name = "wWdUnderline"
Function WdUnderlineFromString(value As String) As WdUnderline
    If IsNumeric(value) Then
        WdUnderlineFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdUnderlineNone": WdUnderlineFromString = wdUnderlineNone
        Case "wdUnderlineSingle": WdUnderlineFromString = wdUnderlineSingle
        Case "wdUnderlineWords": WdUnderlineFromString = wdUnderlineWords
        Case "wdUnderlineDouble": WdUnderlineFromString = wdUnderlineDouble
        Case "wdUnderlineDotted": WdUnderlineFromString = wdUnderlineDotted
        Case "wdUnderlineThick": WdUnderlineFromString = wdUnderlineThick
        Case "wdUnderlineDash": WdUnderlineFromString = wdUnderlineDash
        Case "wdUnderlineDotDash": WdUnderlineFromString = wdUnderlineDotDash
        Case "wdUnderlineDotDotDash": WdUnderlineFromString = wdUnderlineDotDotDash
        Case "wdUnderlineWavy": WdUnderlineFromString = wdUnderlineWavy
        Case "wdUnderlineDottedHeavy": WdUnderlineFromString = wdUnderlineDottedHeavy
        Case "wdUnderlineDashHeavy": WdUnderlineFromString = wdUnderlineDashHeavy
        Case "wdUnderlineDotDashHeavy": WdUnderlineFromString = wdUnderlineDotDashHeavy
        Case "wdUnderlineDotDotDashHeavy": WdUnderlineFromString = wdUnderlineDotDotDashHeavy
        Case "wdUnderlineWavyHeavy": WdUnderlineFromString = wdUnderlineWavyHeavy
        Case "wdUnderlineDashLong": WdUnderlineFromString = wdUnderlineDashLong
        Case "wdUnderlineWavyDouble": WdUnderlineFromString = wdUnderlineWavyDouble
        Case "wdUnderlineDashLongHeavy": WdUnderlineFromString = wdUnderlineDashLongHeavy
    End Select
End Function

Function WdUnderlineToString(value As WdUnderline) As String
    Select Case value
        Case wdUnderlineNone: WdUnderlineToString = "wdUnderlineNone"
        Case wdUnderlineSingle: WdUnderlineToString = "wdUnderlineSingle"
        Case wdUnderlineWords: WdUnderlineToString = "wdUnderlineWords"
        Case wdUnderlineDouble: WdUnderlineToString = "wdUnderlineDouble"
        Case wdUnderlineDotted: WdUnderlineToString = "wdUnderlineDotted"
        Case wdUnderlineThick: WdUnderlineToString = "wdUnderlineThick"
        Case wdUnderlineDash: WdUnderlineToString = "wdUnderlineDash"
        Case wdUnderlineDotDash: WdUnderlineToString = "wdUnderlineDotDash"
        Case wdUnderlineDotDotDash: WdUnderlineToString = "wdUnderlineDotDotDash"
        Case wdUnderlineWavy: WdUnderlineToString = "wdUnderlineWavy"
        Case wdUnderlineDottedHeavy: WdUnderlineToString = "wdUnderlineDottedHeavy"
        Case wdUnderlineDashHeavy: WdUnderlineToString = "wdUnderlineDashHeavy"
        Case wdUnderlineDotDashHeavy: WdUnderlineToString = "wdUnderlineDotDashHeavy"
        Case wdUnderlineDotDotDashHeavy: WdUnderlineToString = "wdUnderlineDotDotDashHeavy"
        Case wdUnderlineWavyHeavy: WdUnderlineToString = "wdUnderlineWavyHeavy"
        Case wdUnderlineDashLong: WdUnderlineToString = "wdUnderlineDashLong"
        Case wdUnderlineWavyDouble: WdUnderlineToString = "wdUnderlineWavyDouble"
        Case wdUnderlineDashLongHeavy: WdUnderlineToString = "wdUnderlineDashLongHeavy"
    End Select
End Function
