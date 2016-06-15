Attribute VB_Name = "wMsoAnimDirection"
Function MsoAnimDirectionFromString(value As String) As MsoAnimDirection
    If IsNumeric(value) Then
        MsoAnimDirectionFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoAnimDirectionNone": MsoAnimDirectionFromString = msoAnimDirectionNone
        Case "msoAnimDirectionUp": MsoAnimDirectionFromString = msoAnimDirectionUp
        Case "msoAnimDirectionRight": MsoAnimDirectionFromString = msoAnimDirectionRight
        Case "msoAnimDirectionDown": MsoAnimDirectionFromString = msoAnimDirectionDown
        Case "msoAnimDirectionLeft": MsoAnimDirectionFromString = msoAnimDirectionLeft
        Case "msoAnimDirectionOrdinalMask": MsoAnimDirectionFromString = msoAnimDirectionOrdinalMask
        Case "msoAnimDirectionUpLeft": MsoAnimDirectionFromString = msoAnimDirectionUpLeft
        Case "msoAnimDirectionUpRight": MsoAnimDirectionFromString = msoAnimDirectionUpRight
        Case "msoAnimDirectionDownRight": MsoAnimDirectionFromString = msoAnimDirectionDownRight
        Case "msoAnimDirectionDownLeft": MsoAnimDirectionFromString = msoAnimDirectionDownLeft
        Case "msoAnimDirectionTop": MsoAnimDirectionFromString = msoAnimDirectionTop
        Case "msoAnimDirectionBottom": MsoAnimDirectionFromString = msoAnimDirectionBottom
        Case "msoAnimDirectionTopLeft": MsoAnimDirectionFromString = msoAnimDirectionTopLeft
        Case "msoAnimDirectionTopRight": MsoAnimDirectionFromString = msoAnimDirectionTopRight
        Case "msoAnimDirectionBottomRight": MsoAnimDirectionFromString = msoAnimDirectionBottomRight
        Case "msoAnimDirectionBottomLeft": MsoAnimDirectionFromString = msoAnimDirectionBottomLeft
        Case "msoAnimDirectionHorizontal": MsoAnimDirectionFromString = msoAnimDirectionHorizontal
        Case "msoAnimDirectionVertical": MsoAnimDirectionFromString = msoAnimDirectionVertical
        Case "msoAnimDirectionAcross": MsoAnimDirectionFromString = msoAnimDirectionAcross
        Case "msoAnimDirectionIn": MsoAnimDirectionFromString = msoAnimDirectionIn
        Case "msoAnimDirectionOut": MsoAnimDirectionFromString = msoAnimDirectionOut
        Case "msoAnimDirectionClockwise": MsoAnimDirectionFromString = msoAnimDirectionClockwise
        Case "msoAnimDirectionCounterclockwise": MsoAnimDirectionFromString = msoAnimDirectionCounterclockwise
        Case "msoAnimDirectionHorizontalIn": MsoAnimDirectionFromString = msoAnimDirectionHorizontalIn
        Case "msoAnimDirectionHorizontalOut": MsoAnimDirectionFromString = msoAnimDirectionHorizontalOut
        Case "msoAnimDirectionVerticalIn": MsoAnimDirectionFromString = msoAnimDirectionVerticalIn
        Case "msoAnimDirectionVerticalOut": MsoAnimDirectionFromString = msoAnimDirectionVerticalOut
        Case "msoAnimDirectionSlightly": MsoAnimDirectionFromString = msoAnimDirectionSlightly
        Case "msoAnimDirectionCenter": MsoAnimDirectionFromString = msoAnimDirectionCenter
        Case "msoAnimDirectionInSlightly": MsoAnimDirectionFromString = msoAnimDirectionInSlightly
        Case "msoAnimDirectionInCenter": MsoAnimDirectionFromString = msoAnimDirectionInCenter
        Case "msoAnimDirectionInBottom": MsoAnimDirectionFromString = msoAnimDirectionInBottom
        Case "msoAnimDirectionOutSlightly": MsoAnimDirectionFromString = msoAnimDirectionOutSlightly
        Case "msoAnimDirectionOutCenter": MsoAnimDirectionFromString = msoAnimDirectionOutCenter
        Case "msoAnimDirectionOutBottom": MsoAnimDirectionFromString = msoAnimDirectionOutBottom
        Case "msoAnimDirectionFontBold": MsoAnimDirectionFromString = msoAnimDirectionFontBold
        Case "msoAnimDirectionFontItalic": MsoAnimDirectionFromString = msoAnimDirectionFontItalic
        Case "msoAnimDirectionFontUnderline": MsoAnimDirectionFromString = msoAnimDirectionFontUnderline
        Case "msoAnimDirectionFontStrikethrough": MsoAnimDirectionFromString = msoAnimDirectionFontStrikethrough
        Case "msoAnimDirectionFontShadow": MsoAnimDirectionFromString = msoAnimDirectionFontShadow
        Case "msoAnimDirectionFontAllCaps": MsoAnimDirectionFromString = msoAnimDirectionFontAllCaps
        Case "msoAnimDirectionInstant": MsoAnimDirectionFromString = msoAnimDirectionInstant
        Case "msoAnimDirectionGradual": MsoAnimDirectionFromString = msoAnimDirectionGradual
        Case "msoAnimDirectionCycleClockwise": MsoAnimDirectionFromString = msoAnimDirectionCycleClockwise
        Case "msoAnimDirectionCycleCounterclockwise": MsoAnimDirectionFromString = msoAnimDirectionCycleCounterclockwise
    End Select
End Function

Function MsoAnimDirectionToString(value As MsoAnimDirection) As String
    Select Case value
        Case msoAnimDirectionNone: MsoAnimDirectionToString = "msoAnimDirectionNone"
        Case msoAnimDirectionUp: MsoAnimDirectionToString = "msoAnimDirectionUp"
        Case msoAnimDirectionRight: MsoAnimDirectionToString = "msoAnimDirectionRight"
        Case msoAnimDirectionDown: MsoAnimDirectionToString = "msoAnimDirectionDown"
        Case msoAnimDirectionLeft: MsoAnimDirectionToString = "msoAnimDirectionLeft"
        Case msoAnimDirectionOrdinalMask: MsoAnimDirectionToString = "msoAnimDirectionOrdinalMask"
        Case msoAnimDirectionUpLeft: MsoAnimDirectionToString = "msoAnimDirectionUpLeft"
        Case msoAnimDirectionUpRight: MsoAnimDirectionToString = "msoAnimDirectionUpRight"
        Case msoAnimDirectionDownRight: MsoAnimDirectionToString = "msoAnimDirectionDownRight"
        Case msoAnimDirectionDownLeft: MsoAnimDirectionToString = "msoAnimDirectionDownLeft"
        Case msoAnimDirectionTop: MsoAnimDirectionToString = "msoAnimDirectionTop"
        Case msoAnimDirectionBottom: MsoAnimDirectionToString = "msoAnimDirectionBottom"
        Case msoAnimDirectionTopLeft: MsoAnimDirectionToString = "msoAnimDirectionTopLeft"
        Case msoAnimDirectionTopRight: MsoAnimDirectionToString = "msoAnimDirectionTopRight"
        Case msoAnimDirectionBottomRight: MsoAnimDirectionToString = "msoAnimDirectionBottomRight"
        Case msoAnimDirectionBottomLeft: MsoAnimDirectionToString = "msoAnimDirectionBottomLeft"
        Case msoAnimDirectionHorizontal: MsoAnimDirectionToString = "msoAnimDirectionHorizontal"
        Case msoAnimDirectionVertical: MsoAnimDirectionToString = "msoAnimDirectionVertical"
        Case msoAnimDirectionAcross: MsoAnimDirectionToString = "msoAnimDirectionAcross"
        Case msoAnimDirectionIn: MsoAnimDirectionToString = "msoAnimDirectionIn"
        Case msoAnimDirectionOut: MsoAnimDirectionToString = "msoAnimDirectionOut"
        Case msoAnimDirectionClockwise: MsoAnimDirectionToString = "msoAnimDirectionClockwise"
        Case msoAnimDirectionCounterclockwise: MsoAnimDirectionToString = "msoAnimDirectionCounterclockwise"
        Case msoAnimDirectionHorizontalIn: MsoAnimDirectionToString = "msoAnimDirectionHorizontalIn"
        Case msoAnimDirectionHorizontalOut: MsoAnimDirectionToString = "msoAnimDirectionHorizontalOut"
        Case msoAnimDirectionVerticalIn: MsoAnimDirectionToString = "msoAnimDirectionVerticalIn"
        Case msoAnimDirectionVerticalOut: MsoAnimDirectionToString = "msoAnimDirectionVerticalOut"
        Case msoAnimDirectionSlightly: MsoAnimDirectionToString = "msoAnimDirectionSlightly"
        Case msoAnimDirectionCenter: MsoAnimDirectionToString = "msoAnimDirectionCenter"
        Case msoAnimDirectionInSlightly: MsoAnimDirectionToString = "msoAnimDirectionInSlightly"
        Case msoAnimDirectionInCenter: MsoAnimDirectionToString = "msoAnimDirectionInCenter"
        Case msoAnimDirectionInBottom: MsoAnimDirectionToString = "msoAnimDirectionInBottom"
        Case msoAnimDirectionOutSlightly: MsoAnimDirectionToString = "msoAnimDirectionOutSlightly"
        Case msoAnimDirectionOutCenter: MsoAnimDirectionToString = "msoAnimDirectionOutCenter"
        Case msoAnimDirectionOutBottom: MsoAnimDirectionToString = "msoAnimDirectionOutBottom"
        Case msoAnimDirectionFontBold: MsoAnimDirectionToString = "msoAnimDirectionFontBold"
        Case msoAnimDirectionFontItalic: MsoAnimDirectionToString = "msoAnimDirectionFontItalic"
        Case msoAnimDirectionFontUnderline: MsoAnimDirectionToString = "msoAnimDirectionFontUnderline"
        Case msoAnimDirectionFontStrikethrough: MsoAnimDirectionToString = "msoAnimDirectionFontStrikethrough"
        Case msoAnimDirectionFontShadow: MsoAnimDirectionToString = "msoAnimDirectionFontShadow"
        Case msoAnimDirectionFontAllCaps: MsoAnimDirectionToString = "msoAnimDirectionFontAllCaps"
        Case msoAnimDirectionInstant: MsoAnimDirectionToString = "msoAnimDirectionInstant"
        Case msoAnimDirectionGradual: MsoAnimDirectionToString = "msoAnimDirectionGradual"
        Case msoAnimDirectionCycleClockwise: MsoAnimDirectionToString = "msoAnimDirectionCycleClockwise"
        Case msoAnimDirectionCycleCounterclockwise: MsoAnimDirectionToString = "msoAnimDirectionCycleCounterclockwise"
    End Select
End Function
