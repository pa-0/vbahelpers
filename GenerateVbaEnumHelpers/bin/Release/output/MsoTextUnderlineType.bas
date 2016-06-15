Attribute VB_Name = "wMsoTextUnderlineType"
Function MsoTextUnderlineTypeFromString(value As String) As MsoTextUnderlineType
    If IsNumeric(value) Then
        MsoTextUnderlineTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoNoUnderline": MsoTextUnderlineTypeFromString = msoNoUnderline
        Case "msoUnderlineWords": MsoTextUnderlineTypeFromString = msoUnderlineWords
        Case "msoUnderlineSingleLine": MsoTextUnderlineTypeFromString = msoUnderlineSingleLine
        Case "msoUnderlineDoubleLine": MsoTextUnderlineTypeFromString = msoUnderlineDoubleLine
        Case "msoUnderlineHeavyLine": MsoTextUnderlineTypeFromString = msoUnderlineHeavyLine
        Case "msoUnderlineDottedLine": MsoTextUnderlineTypeFromString = msoUnderlineDottedLine
        Case "msoUnderlineDottedHeavyLine": MsoTextUnderlineTypeFromString = msoUnderlineDottedHeavyLine
        Case "msoUnderlineDashLine": MsoTextUnderlineTypeFromString = msoUnderlineDashLine
        Case "msoUnderlineDashHeavyLine": MsoTextUnderlineTypeFromString = msoUnderlineDashHeavyLine
        Case "msoUnderlineDashLongLine": MsoTextUnderlineTypeFromString = msoUnderlineDashLongLine
        Case "msoUnderlineDashLongHeavyLine": MsoTextUnderlineTypeFromString = msoUnderlineDashLongHeavyLine
        Case "msoUnderlineDotDashLine": MsoTextUnderlineTypeFromString = msoUnderlineDotDashLine
        Case "msoUnderlineDotDashHeavyLine": MsoTextUnderlineTypeFromString = msoUnderlineDotDashHeavyLine
        Case "msoUnderlineDotDotDashLine": MsoTextUnderlineTypeFromString = msoUnderlineDotDotDashLine
        Case "msoUnderlineDotDotDashHeavyLine": MsoTextUnderlineTypeFromString = msoUnderlineDotDotDashHeavyLine
        Case "msoUnderlineWavyLine": MsoTextUnderlineTypeFromString = msoUnderlineWavyLine
        Case "msoUnderlineWavyHeavyLine": MsoTextUnderlineTypeFromString = msoUnderlineWavyHeavyLine
        Case "msoUnderlineWavyDoubleLine": MsoTextUnderlineTypeFromString = msoUnderlineWavyDoubleLine
        Case "msoUnderlineMixed": MsoTextUnderlineTypeFromString = msoUnderlineMixed
    End Select
End Function

Function MsoTextUnderlineTypeToString(value As MsoTextUnderlineType) As String
    Select Case value
        Case msoNoUnderline: MsoTextUnderlineTypeToString = "msoNoUnderline"
        Case msoUnderlineWords: MsoTextUnderlineTypeToString = "msoUnderlineWords"
        Case msoUnderlineSingleLine: MsoTextUnderlineTypeToString = "msoUnderlineSingleLine"
        Case msoUnderlineDoubleLine: MsoTextUnderlineTypeToString = "msoUnderlineDoubleLine"
        Case msoUnderlineHeavyLine: MsoTextUnderlineTypeToString = "msoUnderlineHeavyLine"
        Case msoUnderlineDottedLine: MsoTextUnderlineTypeToString = "msoUnderlineDottedLine"
        Case msoUnderlineDottedHeavyLine: MsoTextUnderlineTypeToString = "msoUnderlineDottedHeavyLine"
        Case msoUnderlineDashLine: MsoTextUnderlineTypeToString = "msoUnderlineDashLine"
        Case msoUnderlineDashHeavyLine: MsoTextUnderlineTypeToString = "msoUnderlineDashHeavyLine"
        Case msoUnderlineDashLongLine: MsoTextUnderlineTypeToString = "msoUnderlineDashLongLine"
        Case msoUnderlineDashLongHeavyLine: MsoTextUnderlineTypeToString = "msoUnderlineDashLongHeavyLine"
        Case msoUnderlineDotDashLine: MsoTextUnderlineTypeToString = "msoUnderlineDotDashLine"
        Case msoUnderlineDotDashHeavyLine: MsoTextUnderlineTypeToString = "msoUnderlineDotDashHeavyLine"
        Case msoUnderlineDotDotDashLine: MsoTextUnderlineTypeToString = "msoUnderlineDotDotDashLine"
        Case msoUnderlineDotDotDashHeavyLine: MsoTextUnderlineTypeToString = "msoUnderlineDotDotDashHeavyLine"
        Case msoUnderlineWavyLine: MsoTextUnderlineTypeToString = "msoUnderlineWavyLine"
        Case msoUnderlineWavyHeavyLine: MsoTextUnderlineTypeToString = "msoUnderlineWavyHeavyLine"
        Case msoUnderlineWavyDoubleLine: MsoTextUnderlineTypeToString = "msoUnderlineWavyDoubleLine"
        Case msoUnderlineMixed: MsoTextUnderlineTypeToString = "msoUnderlineMixed"
    End Select
End Function
