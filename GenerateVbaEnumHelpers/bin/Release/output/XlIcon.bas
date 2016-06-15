Attribute VB_Name = "wXlIcon"
Function XlIconFromString(value As String) As XlIcon
    If IsNumeric(value) Then
        XlIconFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlIconGreenUpArrow": XlIconFromString = xlIconGreenUpArrow
        Case "xlIconYellowSideArrow": XlIconFromString = xlIconYellowSideArrow
        Case "xlIconRedDownArrow": XlIconFromString = xlIconRedDownArrow
        Case "xlIconGrayUpArrow": XlIconFromString = xlIconGrayUpArrow
        Case "xlIconGraySideArrow": XlIconFromString = xlIconGraySideArrow
        Case "xlIconGrayDownArrow": XlIconFromString = xlIconGrayDownArrow
        Case "xlIconGreenFlag": XlIconFromString = xlIconGreenFlag
        Case "xlIconYellowFlag": XlIconFromString = xlIconYellowFlag
        Case "xlIconRedFlag": XlIconFromString = xlIconRedFlag
        Case "xlIconGreenCircle": XlIconFromString = xlIconGreenCircle
        Case "xlIconYellowCircle": XlIconFromString = xlIconYellowCircle
        Case "xlIconRedCircleWithBorder": XlIconFromString = xlIconRedCircleWithBorder
        Case "xlIconBlackCircleWithBorder": XlIconFromString = xlIconBlackCircleWithBorder
        Case "xlIconGreenTrafficLight": XlIconFromString = xlIconGreenTrafficLight
        Case "xlIconYellowTrafficLight": XlIconFromString = xlIconYellowTrafficLight
        Case "xlIconRedTrafficLight": XlIconFromString = xlIconRedTrafficLight
        Case "xlIconYellowTriangle": XlIconFromString = xlIconYellowTriangle
        Case "xlIconRedDiamond": XlIconFromString = xlIconRedDiamond
        Case "xlIconGreenCheckSymbol": XlIconFromString = xlIconGreenCheckSymbol
        Case "xlIconYellowExclamationSymbol": XlIconFromString = xlIconYellowExclamationSymbol
        Case "xlIconRedCrossSymbol": XlIconFromString = xlIconRedCrossSymbol
        Case "xlIconGreenCheck": XlIconFromString = xlIconGreenCheck
        Case "xlIconYellowExclamation": XlIconFromString = xlIconYellowExclamation
        Case "xlIconRedCross": XlIconFromString = xlIconRedCross
        Case "xlIconYellowUpInclineArrow": XlIconFromString = xlIconYellowUpInclineArrow
        Case "xlIconYellowDownInclineArrow": XlIconFromString = xlIconYellowDownInclineArrow
        Case "xlIconGrayUpInclineArrow": XlIconFromString = xlIconGrayUpInclineArrow
        Case "xlIconGrayDownInclineArrow": XlIconFromString = xlIconGrayDownInclineArrow
        Case "xlIconRedCircle": XlIconFromString = xlIconRedCircle
        Case "xlIconPinkCircle": XlIconFromString = xlIconPinkCircle
        Case "xlIconGrayCircle": XlIconFromString = xlIconGrayCircle
        Case "xlIconBlackCircle": XlIconFromString = xlIconBlackCircle
        Case "xlIconCircleWithOneWhiteQuarter": XlIconFromString = xlIconCircleWithOneWhiteQuarter
        Case "xlIconCircleWithTwoWhiteQuarters": XlIconFromString = xlIconCircleWithTwoWhiteQuarters
        Case "xlIconCircleWithThreeWhiteQuarters": XlIconFromString = xlIconCircleWithThreeWhiteQuarters
        Case "xlIconWhiteCircleAllWhiteQuarters": XlIconFromString = xlIconWhiteCircleAllWhiteQuarters
        Case "xlIcon0Bars": XlIconFromString = xlIcon0Bars
        Case "xlIcon1Bar": XlIconFromString = xlIcon1Bar
        Case "xlIcon2Bars": XlIconFromString = xlIcon2Bars
        Case "xlIcon3Bars": XlIconFromString = xlIcon3Bars
        Case "xlIcon4Bars": XlIconFromString = xlIcon4Bars
        Case "xlIconGoldStar": XlIconFromString = xlIconGoldStar
        Case "xlIconHalfGoldStar": XlIconFromString = xlIconHalfGoldStar
        Case "xlIconSilverStar": XlIconFromString = xlIconSilverStar
        Case "xlIconGreenUpTriangle": XlIconFromString = xlIconGreenUpTriangle
        Case "xlIconYellowDash": XlIconFromString = xlIconYellowDash
        Case "xlIconRedDownTriangle": XlIconFromString = xlIconRedDownTriangle
        Case "xlIcon4FilledBoxes": XlIconFromString = xlIcon4FilledBoxes
        Case "xlIcon3FilledBoxes": XlIconFromString = xlIcon3FilledBoxes
        Case "xlIcon2FilledBoxes": XlIconFromString = xlIcon2FilledBoxes
        Case "xlIcon1FilledBox": XlIconFromString = xlIcon1FilledBox
        Case "xlIcon0FilledBoxes": XlIconFromString = xlIcon0FilledBoxes
        Case "xlIconNoCellIcon": XlIconFromString = xlIconNoCellIcon
    End Select
End Function

Function XlIconToString(value As XlIcon) As String
    Select Case value
        Case xlIconGreenUpArrow: XlIconToString = "xlIconGreenUpArrow"
        Case xlIconYellowSideArrow: XlIconToString = "xlIconYellowSideArrow"
        Case xlIconRedDownArrow: XlIconToString = "xlIconRedDownArrow"
        Case xlIconGrayUpArrow: XlIconToString = "xlIconGrayUpArrow"
        Case xlIconGraySideArrow: XlIconToString = "xlIconGraySideArrow"
        Case xlIconGrayDownArrow: XlIconToString = "xlIconGrayDownArrow"
        Case xlIconGreenFlag: XlIconToString = "xlIconGreenFlag"
        Case xlIconYellowFlag: XlIconToString = "xlIconYellowFlag"
        Case xlIconRedFlag: XlIconToString = "xlIconRedFlag"
        Case xlIconGreenCircle: XlIconToString = "xlIconGreenCircle"
        Case xlIconYellowCircle: XlIconToString = "xlIconYellowCircle"
        Case xlIconRedCircleWithBorder: XlIconToString = "xlIconRedCircleWithBorder"
        Case xlIconBlackCircleWithBorder: XlIconToString = "xlIconBlackCircleWithBorder"
        Case xlIconGreenTrafficLight: XlIconToString = "xlIconGreenTrafficLight"
        Case xlIconYellowTrafficLight: XlIconToString = "xlIconYellowTrafficLight"
        Case xlIconRedTrafficLight: XlIconToString = "xlIconRedTrafficLight"
        Case xlIconYellowTriangle: XlIconToString = "xlIconYellowTriangle"
        Case xlIconRedDiamond: XlIconToString = "xlIconRedDiamond"
        Case xlIconGreenCheckSymbol: XlIconToString = "xlIconGreenCheckSymbol"
        Case xlIconYellowExclamationSymbol: XlIconToString = "xlIconYellowExclamationSymbol"
        Case xlIconRedCrossSymbol: XlIconToString = "xlIconRedCrossSymbol"
        Case xlIconGreenCheck: XlIconToString = "xlIconGreenCheck"
        Case xlIconYellowExclamation: XlIconToString = "xlIconYellowExclamation"
        Case xlIconRedCross: XlIconToString = "xlIconRedCross"
        Case xlIconYellowUpInclineArrow: XlIconToString = "xlIconYellowUpInclineArrow"
        Case xlIconYellowDownInclineArrow: XlIconToString = "xlIconYellowDownInclineArrow"
        Case xlIconGrayUpInclineArrow: XlIconToString = "xlIconGrayUpInclineArrow"
        Case xlIconGrayDownInclineArrow: XlIconToString = "xlIconGrayDownInclineArrow"
        Case xlIconRedCircle: XlIconToString = "xlIconRedCircle"
        Case xlIconPinkCircle: XlIconToString = "xlIconPinkCircle"
        Case xlIconGrayCircle: XlIconToString = "xlIconGrayCircle"
        Case xlIconBlackCircle: XlIconToString = "xlIconBlackCircle"
        Case xlIconCircleWithOneWhiteQuarter: XlIconToString = "xlIconCircleWithOneWhiteQuarter"
        Case xlIconCircleWithTwoWhiteQuarters: XlIconToString = "xlIconCircleWithTwoWhiteQuarters"
        Case xlIconCircleWithThreeWhiteQuarters: XlIconToString = "xlIconCircleWithThreeWhiteQuarters"
        Case xlIconWhiteCircleAllWhiteQuarters: XlIconToString = "xlIconWhiteCircleAllWhiteQuarters"
        Case xlIcon0Bars: XlIconToString = "xlIcon0Bars"
        Case xlIcon1Bar: XlIconToString = "xlIcon1Bar"
        Case xlIcon2Bars: XlIconToString = "xlIcon2Bars"
        Case xlIcon3Bars: XlIconToString = "xlIcon3Bars"
        Case xlIcon4Bars: XlIconToString = "xlIcon4Bars"
        Case xlIconGoldStar: XlIconToString = "xlIconGoldStar"
        Case xlIconHalfGoldStar: XlIconToString = "xlIconHalfGoldStar"
        Case xlIconSilverStar: XlIconToString = "xlIconSilverStar"
        Case xlIconGreenUpTriangle: XlIconToString = "xlIconGreenUpTriangle"
        Case xlIconYellowDash: XlIconToString = "xlIconYellowDash"
        Case xlIconRedDownTriangle: XlIconToString = "xlIconRedDownTriangle"
        Case xlIcon4FilledBoxes: XlIconToString = "xlIcon4FilledBoxes"
        Case xlIcon3FilledBoxes: XlIconToString = "xlIcon3FilledBoxes"
        Case xlIcon2FilledBoxes: XlIconToString = "xlIcon2FilledBoxes"
        Case xlIcon1FilledBox: XlIconToString = "xlIcon1FilledBox"
        Case xlIcon0FilledBoxes: XlIconToString = "xlIcon0FilledBoxes"
        Case xlIconNoCellIcon: XlIconToString = "xlIconNoCellIcon"
    End Select
End Function
