Attribute VB_Name = "wPbUnderlineType"
Function PbUnderlineTypeFromString(value As String) As PbUnderlineType
    If IsNumeric(value) Then
        PbUnderlineTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "pbUnderlineNone": PbUnderlineTypeFromString = pbUnderlineNone
        Case "pbUnderlineSingle": PbUnderlineTypeFromString = pbUnderlineSingle
        Case "pbUnderlineWordsOnly": PbUnderlineTypeFromString = pbUnderlineWordsOnly
        Case "pbUnderlineDouble": PbUnderlineTypeFromString = pbUnderlineDouble
        Case "pbUnderlineDotted": PbUnderlineTypeFromString = pbUnderlineDotted
        Case "pbUnderlineThick": PbUnderlineTypeFromString = pbUnderlineThick
        Case "pbUnderlineDash": PbUnderlineTypeFromString = pbUnderlineDash
        Case "pbUnderlineDotDash": PbUnderlineTypeFromString = pbUnderlineDotDash
        Case "pbUnderlineDotDotDash": PbUnderlineTypeFromString = pbUnderlineDotDotDash
        Case "pbUnderlineWavy": PbUnderlineTypeFromString = pbUnderlineWavy
        Case "pbUnderlineWavyHeavy": PbUnderlineTypeFromString = pbUnderlineWavyHeavy
        Case "pbUnderlineDotHeavy": PbUnderlineTypeFromString = pbUnderlineDotHeavy
        Case "pbUnderlineDashHeavy": PbUnderlineTypeFromString = pbUnderlineDashHeavy
        Case "pbUnderlineDotDashHeavy": PbUnderlineTypeFromString = pbUnderlineDotDashHeavy
        Case "pbUnderlineDotDotDashHeavy": PbUnderlineTypeFromString = pbUnderlineDotDotDashHeavy
        Case "pbUnderlineDashLong": PbUnderlineTypeFromString = pbUnderlineDashLong
        Case "pbUnderlineDashLongHeavy": PbUnderlineTypeFromString = pbUnderlineDashLongHeavy
        Case "pbUnderlineWavyDouble": PbUnderlineTypeFromString = pbUnderlineWavyDouble
        Case "pbUnderlineMixed": PbUnderlineTypeFromString = pbUnderlineMixed
    End Select
End Function

Function PbUnderlineTypeToString(value As PbUnderlineType) As String
    Select Case value
        Case pbUnderlineNone: PbUnderlineTypeToString = "pbUnderlineNone"
        Case pbUnderlineSingle: PbUnderlineTypeToString = "pbUnderlineSingle"
        Case pbUnderlineWordsOnly: PbUnderlineTypeToString = "pbUnderlineWordsOnly"
        Case pbUnderlineDouble: PbUnderlineTypeToString = "pbUnderlineDouble"
        Case pbUnderlineDotted: PbUnderlineTypeToString = "pbUnderlineDotted"
        Case pbUnderlineThick: PbUnderlineTypeToString = "pbUnderlineThick"
        Case pbUnderlineDash: PbUnderlineTypeToString = "pbUnderlineDash"
        Case pbUnderlineDotDash: PbUnderlineTypeToString = "pbUnderlineDotDash"
        Case pbUnderlineDotDotDash: PbUnderlineTypeToString = "pbUnderlineDotDotDash"
        Case pbUnderlineWavy: PbUnderlineTypeToString = "pbUnderlineWavy"
        Case pbUnderlineWavyHeavy: PbUnderlineTypeToString = "pbUnderlineWavyHeavy"
        Case pbUnderlineDotHeavy: PbUnderlineTypeToString = "pbUnderlineDotHeavy"
        Case pbUnderlineDashHeavy: PbUnderlineTypeToString = "pbUnderlineDashHeavy"
        Case pbUnderlineDotDashHeavy: PbUnderlineTypeToString = "pbUnderlineDotDashHeavy"
        Case pbUnderlineDotDotDashHeavy: PbUnderlineTypeToString = "pbUnderlineDotDotDashHeavy"
        Case pbUnderlineDashLong: PbUnderlineTypeToString = "pbUnderlineDashLong"
        Case pbUnderlineDashLongHeavy: PbUnderlineTypeToString = "pbUnderlineDashLongHeavy"
        Case pbUnderlineWavyDouble: PbUnderlineTypeToString = "pbUnderlineWavyDouble"
        Case pbUnderlineMixed: PbUnderlineTypeToString = "pbUnderlineMixed"
    End Select
End Function
