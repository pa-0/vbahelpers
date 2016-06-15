Attribute VB_Name = "wMsoPatternType"
Function MsoPatternTypeFromString(value As String) As MsoPatternType
    If IsNumeric(value) Then
        MsoPatternTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoPattern5Percent": MsoPatternTypeFromString = msoPattern5Percent
        Case "msoPattern10Percent": MsoPatternTypeFromString = msoPattern10Percent
        Case "msoPattern20Percent": MsoPatternTypeFromString = msoPattern20Percent
        Case "msoPattern25Percent": MsoPatternTypeFromString = msoPattern25Percent
        Case "msoPattern30Percent": MsoPatternTypeFromString = msoPattern30Percent
        Case "msoPattern40Percent": MsoPatternTypeFromString = msoPattern40Percent
        Case "msoPattern50Percent": MsoPatternTypeFromString = msoPattern50Percent
        Case "msoPattern60Percent": MsoPatternTypeFromString = msoPattern60Percent
        Case "msoPattern70Percent": MsoPatternTypeFromString = msoPattern70Percent
        Case "msoPattern75Percent": MsoPatternTypeFromString = msoPattern75Percent
        Case "msoPattern80Percent": MsoPatternTypeFromString = msoPattern80Percent
        Case "msoPattern90Percent": MsoPatternTypeFromString = msoPattern90Percent
        Case "msoPatternDarkHorizontal": MsoPatternTypeFromString = msoPatternDarkHorizontal
        Case "msoPatternDarkVertical": MsoPatternTypeFromString = msoPatternDarkVertical
        Case "msoPatternDarkDownwardDiagonal": MsoPatternTypeFromString = msoPatternDarkDownwardDiagonal
        Case "msoPatternDarkUpwardDiagonal": MsoPatternTypeFromString = msoPatternDarkUpwardDiagonal
        Case "msoPatternSmallCheckerBoard": MsoPatternTypeFromString = msoPatternSmallCheckerBoard
        Case "msoPatternTrellis": MsoPatternTypeFromString = msoPatternTrellis
        Case "msoPatternLightHorizontal": MsoPatternTypeFromString = msoPatternLightHorizontal
        Case "msoPatternLightVertical": MsoPatternTypeFromString = msoPatternLightVertical
        Case "msoPatternLightDownwardDiagonal": MsoPatternTypeFromString = msoPatternLightDownwardDiagonal
        Case "msoPatternLightUpwardDiagonal": MsoPatternTypeFromString = msoPatternLightUpwardDiagonal
        Case "msoPatternSmallGrid": MsoPatternTypeFromString = msoPatternSmallGrid
        Case "msoPatternDottedDiamond": MsoPatternTypeFromString = msoPatternDottedDiamond
        Case "msoPatternWideDownwardDiagonal": MsoPatternTypeFromString = msoPatternWideDownwardDiagonal
        Case "msoPatternWideUpwardDiagonal": MsoPatternTypeFromString = msoPatternWideUpwardDiagonal
        Case "msoPatternDashedUpwardDiagonal": MsoPatternTypeFromString = msoPatternDashedUpwardDiagonal
        Case "msoPatternDashedDownwardDiagonal": MsoPatternTypeFromString = msoPatternDashedDownwardDiagonal
        Case "msoPatternNarrowVertical": MsoPatternTypeFromString = msoPatternNarrowVertical
        Case "msoPatternNarrowHorizontal": MsoPatternTypeFromString = msoPatternNarrowHorizontal
        Case "msoPatternDashedVertical": MsoPatternTypeFromString = msoPatternDashedVertical
        Case "msoPatternDashedHorizontal": MsoPatternTypeFromString = msoPatternDashedHorizontal
        Case "msoPatternLargeConfetti": MsoPatternTypeFromString = msoPatternLargeConfetti
        Case "msoPatternLargeGrid": MsoPatternTypeFromString = msoPatternLargeGrid
        Case "msoPatternHorizontalBrick": MsoPatternTypeFromString = msoPatternHorizontalBrick
        Case "msoPatternLargeCheckerBoard": MsoPatternTypeFromString = msoPatternLargeCheckerBoard
        Case "msoPatternSmallConfetti": MsoPatternTypeFromString = msoPatternSmallConfetti
        Case "msoPatternZigZag": MsoPatternTypeFromString = msoPatternZigZag
        Case "msoPatternSolidDiamond": MsoPatternTypeFromString = msoPatternSolidDiamond
        Case "msoPatternDiagonalBrick": MsoPatternTypeFromString = msoPatternDiagonalBrick
        Case "msoPatternOutlinedDiamond": MsoPatternTypeFromString = msoPatternOutlinedDiamond
        Case "msoPatternPlaid": MsoPatternTypeFromString = msoPatternPlaid
        Case "msoPatternSphere": MsoPatternTypeFromString = msoPatternSphere
        Case "msoPatternWeave": MsoPatternTypeFromString = msoPatternWeave
        Case "msoPatternDottedGrid": MsoPatternTypeFromString = msoPatternDottedGrid
        Case "msoPatternDivot": MsoPatternTypeFromString = msoPatternDivot
        Case "msoPatternShingle": MsoPatternTypeFromString = msoPatternShingle
        Case "msoPatternWave": MsoPatternTypeFromString = msoPatternWave
        Case "msoPatternHorizontal": MsoPatternTypeFromString = msoPatternHorizontal
        Case "msoPatternVertical": MsoPatternTypeFromString = msoPatternVertical
        Case "msoPatternCross": MsoPatternTypeFromString = msoPatternCross
        Case "msoPatternDownwardDiagonal": MsoPatternTypeFromString = msoPatternDownwardDiagonal
        Case "msoPatternUpwardDiagonal": MsoPatternTypeFromString = msoPatternUpwardDiagonal
        Case "msoPatternDiagonalCross": MsoPatternTypeFromString = msoPatternDiagonalCross
        Case "msoPatternMixed": MsoPatternTypeFromString = msoPatternMixed
    End Select
End Function

Function MsoPatternTypeToString(value As MsoPatternType) As String
    Select Case value
        Case msoPattern5Percent: MsoPatternTypeToString = "msoPattern5Percent"
        Case msoPattern10Percent: MsoPatternTypeToString = "msoPattern10Percent"
        Case msoPattern20Percent: MsoPatternTypeToString = "msoPattern20Percent"
        Case msoPattern25Percent: MsoPatternTypeToString = "msoPattern25Percent"
        Case msoPattern30Percent: MsoPatternTypeToString = "msoPattern30Percent"
        Case msoPattern40Percent: MsoPatternTypeToString = "msoPattern40Percent"
        Case msoPattern50Percent: MsoPatternTypeToString = "msoPattern50Percent"
        Case msoPattern60Percent: MsoPatternTypeToString = "msoPattern60Percent"
        Case msoPattern70Percent: MsoPatternTypeToString = "msoPattern70Percent"
        Case msoPattern75Percent: MsoPatternTypeToString = "msoPattern75Percent"
        Case msoPattern80Percent: MsoPatternTypeToString = "msoPattern80Percent"
        Case msoPattern90Percent: MsoPatternTypeToString = "msoPattern90Percent"
        Case msoPatternDarkHorizontal: MsoPatternTypeToString = "msoPatternDarkHorizontal"
        Case msoPatternDarkVertical: MsoPatternTypeToString = "msoPatternDarkVertical"
        Case msoPatternDarkDownwardDiagonal: MsoPatternTypeToString = "msoPatternDarkDownwardDiagonal"
        Case msoPatternDarkUpwardDiagonal: MsoPatternTypeToString = "msoPatternDarkUpwardDiagonal"
        Case msoPatternSmallCheckerBoard: MsoPatternTypeToString = "msoPatternSmallCheckerBoard"
        Case msoPatternTrellis: MsoPatternTypeToString = "msoPatternTrellis"
        Case msoPatternLightHorizontal: MsoPatternTypeToString = "msoPatternLightHorizontal"
        Case msoPatternLightVertical: MsoPatternTypeToString = "msoPatternLightVertical"
        Case msoPatternLightDownwardDiagonal: MsoPatternTypeToString = "msoPatternLightDownwardDiagonal"
        Case msoPatternLightUpwardDiagonal: MsoPatternTypeToString = "msoPatternLightUpwardDiagonal"
        Case msoPatternSmallGrid: MsoPatternTypeToString = "msoPatternSmallGrid"
        Case msoPatternDottedDiamond: MsoPatternTypeToString = "msoPatternDottedDiamond"
        Case msoPatternWideDownwardDiagonal: MsoPatternTypeToString = "msoPatternWideDownwardDiagonal"
        Case msoPatternWideUpwardDiagonal: MsoPatternTypeToString = "msoPatternWideUpwardDiagonal"
        Case msoPatternDashedUpwardDiagonal: MsoPatternTypeToString = "msoPatternDashedUpwardDiagonal"
        Case msoPatternDashedDownwardDiagonal: MsoPatternTypeToString = "msoPatternDashedDownwardDiagonal"
        Case msoPatternNarrowVertical: MsoPatternTypeToString = "msoPatternNarrowVertical"
        Case msoPatternNarrowHorizontal: MsoPatternTypeToString = "msoPatternNarrowHorizontal"
        Case msoPatternDashedVertical: MsoPatternTypeToString = "msoPatternDashedVertical"
        Case msoPatternDashedHorizontal: MsoPatternTypeToString = "msoPatternDashedHorizontal"
        Case msoPatternLargeConfetti: MsoPatternTypeToString = "msoPatternLargeConfetti"
        Case msoPatternLargeGrid: MsoPatternTypeToString = "msoPatternLargeGrid"
        Case msoPatternHorizontalBrick: MsoPatternTypeToString = "msoPatternHorizontalBrick"
        Case msoPatternLargeCheckerBoard: MsoPatternTypeToString = "msoPatternLargeCheckerBoard"
        Case msoPatternSmallConfetti: MsoPatternTypeToString = "msoPatternSmallConfetti"
        Case msoPatternZigZag: MsoPatternTypeToString = "msoPatternZigZag"
        Case msoPatternSolidDiamond: MsoPatternTypeToString = "msoPatternSolidDiamond"
        Case msoPatternDiagonalBrick: MsoPatternTypeToString = "msoPatternDiagonalBrick"
        Case msoPatternOutlinedDiamond: MsoPatternTypeToString = "msoPatternOutlinedDiamond"
        Case msoPatternPlaid: MsoPatternTypeToString = "msoPatternPlaid"
        Case msoPatternSphere: MsoPatternTypeToString = "msoPatternSphere"
        Case msoPatternWeave: MsoPatternTypeToString = "msoPatternWeave"
        Case msoPatternDottedGrid: MsoPatternTypeToString = "msoPatternDottedGrid"
        Case msoPatternDivot: MsoPatternTypeToString = "msoPatternDivot"
        Case msoPatternShingle: MsoPatternTypeToString = "msoPatternShingle"
        Case msoPatternWave: MsoPatternTypeToString = "msoPatternWave"
        Case msoPatternHorizontal: MsoPatternTypeToString = "msoPatternHorizontal"
        Case msoPatternVertical: MsoPatternTypeToString = "msoPatternVertical"
        Case msoPatternCross: MsoPatternTypeToString = "msoPatternCross"
        Case msoPatternDownwardDiagonal: MsoPatternTypeToString = "msoPatternDownwardDiagonal"
        Case msoPatternUpwardDiagonal: MsoPatternTypeToString = "msoPatternUpwardDiagonal"
        Case msoPatternDiagonalCross: MsoPatternTypeToString = "msoPatternDiagonalCross"
        Case msoPatternMixed: MsoPatternTypeToString = "msoPatternMixed"
    End Select
End Function
