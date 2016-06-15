Attribute VB_Name = "wMsoPresetTextEffectShape"
Function MsoPresetTextEffectShapeFromString(value As String) As MsoPresetTextEffectShape
    If IsNumeric(value) Then
        MsoPresetTextEffectShapeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoTextEffectShapePlainText": MsoPresetTextEffectShapeFromString = msoTextEffectShapePlainText
        Case "msoTextEffectShapeStop": MsoPresetTextEffectShapeFromString = msoTextEffectShapeStop
        Case "msoTextEffectShapeTriangleUp": MsoPresetTextEffectShapeFromString = msoTextEffectShapeTriangleUp
        Case "msoTextEffectShapeTriangleDown": MsoPresetTextEffectShapeFromString = msoTextEffectShapeTriangleDown
        Case "msoTextEffectShapeChevronUp": MsoPresetTextEffectShapeFromString = msoTextEffectShapeChevronUp
        Case "msoTextEffectShapeChevronDown": MsoPresetTextEffectShapeFromString = msoTextEffectShapeChevronDown
        Case "msoTextEffectShapeRingInside": MsoPresetTextEffectShapeFromString = msoTextEffectShapeRingInside
        Case "msoTextEffectShapeRingOutside": MsoPresetTextEffectShapeFromString = msoTextEffectShapeRingOutside
        Case "msoTextEffectShapeArchUpCurve": MsoPresetTextEffectShapeFromString = msoTextEffectShapeArchUpCurve
        Case "msoTextEffectShapeArchDownCurve": MsoPresetTextEffectShapeFromString = msoTextEffectShapeArchDownCurve
        Case "msoTextEffectShapeCircleCurve": MsoPresetTextEffectShapeFromString = msoTextEffectShapeCircleCurve
        Case "msoTextEffectShapeButtonCurve": MsoPresetTextEffectShapeFromString = msoTextEffectShapeButtonCurve
        Case "msoTextEffectShapeArchUpPour": MsoPresetTextEffectShapeFromString = msoTextEffectShapeArchUpPour
        Case "msoTextEffectShapeArchDownPour": MsoPresetTextEffectShapeFromString = msoTextEffectShapeArchDownPour
        Case "msoTextEffectShapeCirclePour": MsoPresetTextEffectShapeFromString = msoTextEffectShapeCirclePour
        Case "msoTextEffectShapeButtonPour": MsoPresetTextEffectShapeFromString = msoTextEffectShapeButtonPour
        Case "msoTextEffectShapeCurveUp": MsoPresetTextEffectShapeFromString = msoTextEffectShapeCurveUp
        Case "msoTextEffectShapeCurveDown": MsoPresetTextEffectShapeFromString = msoTextEffectShapeCurveDown
        Case "msoTextEffectShapeCanUp": MsoPresetTextEffectShapeFromString = msoTextEffectShapeCanUp
        Case "msoTextEffectShapeCanDown": MsoPresetTextEffectShapeFromString = msoTextEffectShapeCanDown
        Case "msoTextEffectShapeWave1": MsoPresetTextEffectShapeFromString = msoTextEffectShapeWave1
        Case "msoTextEffectShapeWave2": MsoPresetTextEffectShapeFromString = msoTextEffectShapeWave2
        Case "msoTextEffectShapeDoubleWave1": MsoPresetTextEffectShapeFromString = msoTextEffectShapeDoubleWave1
        Case "msoTextEffectShapeDoubleWave2": MsoPresetTextEffectShapeFromString = msoTextEffectShapeDoubleWave2
        Case "msoTextEffectShapeInflate": MsoPresetTextEffectShapeFromString = msoTextEffectShapeInflate
        Case "msoTextEffectShapeDeflate": MsoPresetTextEffectShapeFromString = msoTextEffectShapeDeflate
        Case "msoTextEffectShapeInflateBottom": MsoPresetTextEffectShapeFromString = msoTextEffectShapeInflateBottom
        Case "msoTextEffectShapeDeflateBottom": MsoPresetTextEffectShapeFromString = msoTextEffectShapeDeflateBottom
        Case "msoTextEffectShapeInflateTop": MsoPresetTextEffectShapeFromString = msoTextEffectShapeInflateTop
        Case "msoTextEffectShapeDeflateTop": MsoPresetTextEffectShapeFromString = msoTextEffectShapeDeflateTop
        Case "msoTextEffectShapeDeflateInflate": MsoPresetTextEffectShapeFromString = msoTextEffectShapeDeflateInflate
        Case "msoTextEffectShapeDeflateInflateDeflate": MsoPresetTextEffectShapeFromString = msoTextEffectShapeDeflateInflateDeflate
        Case "msoTextEffectShapeFadeRight": MsoPresetTextEffectShapeFromString = msoTextEffectShapeFadeRight
        Case "msoTextEffectShapeFadeLeft": MsoPresetTextEffectShapeFromString = msoTextEffectShapeFadeLeft
        Case "msoTextEffectShapeFadeUp": MsoPresetTextEffectShapeFromString = msoTextEffectShapeFadeUp
        Case "msoTextEffectShapeFadeDown": MsoPresetTextEffectShapeFromString = msoTextEffectShapeFadeDown
        Case "msoTextEffectShapeSlantUp": MsoPresetTextEffectShapeFromString = msoTextEffectShapeSlantUp
        Case "msoTextEffectShapeSlantDown": MsoPresetTextEffectShapeFromString = msoTextEffectShapeSlantDown
        Case "msoTextEffectShapeCascadeUp": MsoPresetTextEffectShapeFromString = msoTextEffectShapeCascadeUp
        Case "msoTextEffectShapeCascadeDown": MsoPresetTextEffectShapeFromString = msoTextEffectShapeCascadeDown
        Case "msoTextEffectShapeMixed": MsoPresetTextEffectShapeFromString = msoTextEffectShapeMixed
    End Select
End Function

Function MsoPresetTextEffectShapeToString(value As MsoPresetTextEffectShape) As String
    Select Case value
        Case msoTextEffectShapePlainText: MsoPresetTextEffectShapeToString = "msoTextEffectShapePlainText"
        Case msoTextEffectShapeStop: MsoPresetTextEffectShapeToString = "msoTextEffectShapeStop"
        Case msoTextEffectShapeTriangleUp: MsoPresetTextEffectShapeToString = "msoTextEffectShapeTriangleUp"
        Case msoTextEffectShapeTriangleDown: MsoPresetTextEffectShapeToString = "msoTextEffectShapeTriangleDown"
        Case msoTextEffectShapeChevronUp: MsoPresetTextEffectShapeToString = "msoTextEffectShapeChevronUp"
        Case msoTextEffectShapeChevronDown: MsoPresetTextEffectShapeToString = "msoTextEffectShapeChevronDown"
        Case msoTextEffectShapeRingInside: MsoPresetTextEffectShapeToString = "msoTextEffectShapeRingInside"
        Case msoTextEffectShapeRingOutside: MsoPresetTextEffectShapeToString = "msoTextEffectShapeRingOutside"
        Case msoTextEffectShapeArchUpCurve: MsoPresetTextEffectShapeToString = "msoTextEffectShapeArchUpCurve"
        Case msoTextEffectShapeArchDownCurve: MsoPresetTextEffectShapeToString = "msoTextEffectShapeArchDownCurve"
        Case msoTextEffectShapeCircleCurve: MsoPresetTextEffectShapeToString = "msoTextEffectShapeCircleCurve"
        Case msoTextEffectShapeButtonCurve: MsoPresetTextEffectShapeToString = "msoTextEffectShapeButtonCurve"
        Case msoTextEffectShapeArchUpPour: MsoPresetTextEffectShapeToString = "msoTextEffectShapeArchUpPour"
        Case msoTextEffectShapeArchDownPour: MsoPresetTextEffectShapeToString = "msoTextEffectShapeArchDownPour"
        Case msoTextEffectShapeCirclePour: MsoPresetTextEffectShapeToString = "msoTextEffectShapeCirclePour"
        Case msoTextEffectShapeButtonPour: MsoPresetTextEffectShapeToString = "msoTextEffectShapeButtonPour"
        Case msoTextEffectShapeCurveUp: MsoPresetTextEffectShapeToString = "msoTextEffectShapeCurveUp"
        Case msoTextEffectShapeCurveDown: MsoPresetTextEffectShapeToString = "msoTextEffectShapeCurveDown"
        Case msoTextEffectShapeCanUp: MsoPresetTextEffectShapeToString = "msoTextEffectShapeCanUp"
        Case msoTextEffectShapeCanDown: MsoPresetTextEffectShapeToString = "msoTextEffectShapeCanDown"
        Case msoTextEffectShapeWave1: MsoPresetTextEffectShapeToString = "msoTextEffectShapeWave1"
        Case msoTextEffectShapeWave2: MsoPresetTextEffectShapeToString = "msoTextEffectShapeWave2"
        Case msoTextEffectShapeDoubleWave1: MsoPresetTextEffectShapeToString = "msoTextEffectShapeDoubleWave1"
        Case msoTextEffectShapeDoubleWave2: MsoPresetTextEffectShapeToString = "msoTextEffectShapeDoubleWave2"
        Case msoTextEffectShapeInflate: MsoPresetTextEffectShapeToString = "msoTextEffectShapeInflate"
        Case msoTextEffectShapeDeflate: MsoPresetTextEffectShapeToString = "msoTextEffectShapeDeflate"
        Case msoTextEffectShapeInflateBottom: MsoPresetTextEffectShapeToString = "msoTextEffectShapeInflateBottom"
        Case msoTextEffectShapeDeflateBottom: MsoPresetTextEffectShapeToString = "msoTextEffectShapeDeflateBottom"
        Case msoTextEffectShapeInflateTop: MsoPresetTextEffectShapeToString = "msoTextEffectShapeInflateTop"
        Case msoTextEffectShapeDeflateTop: MsoPresetTextEffectShapeToString = "msoTextEffectShapeDeflateTop"
        Case msoTextEffectShapeDeflateInflate: MsoPresetTextEffectShapeToString = "msoTextEffectShapeDeflateInflate"
        Case msoTextEffectShapeDeflateInflateDeflate: MsoPresetTextEffectShapeToString = "msoTextEffectShapeDeflateInflateDeflate"
        Case msoTextEffectShapeFadeRight: MsoPresetTextEffectShapeToString = "msoTextEffectShapeFadeRight"
        Case msoTextEffectShapeFadeLeft: MsoPresetTextEffectShapeToString = "msoTextEffectShapeFadeLeft"
        Case msoTextEffectShapeFadeUp: MsoPresetTextEffectShapeToString = "msoTextEffectShapeFadeUp"
        Case msoTextEffectShapeFadeDown: MsoPresetTextEffectShapeToString = "msoTextEffectShapeFadeDown"
        Case msoTextEffectShapeSlantUp: MsoPresetTextEffectShapeToString = "msoTextEffectShapeSlantUp"
        Case msoTextEffectShapeSlantDown: MsoPresetTextEffectShapeToString = "msoTextEffectShapeSlantDown"
        Case msoTextEffectShapeCascadeUp: MsoPresetTextEffectShapeToString = "msoTextEffectShapeCascadeUp"
        Case msoTextEffectShapeCascadeDown: MsoPresetTextEffectShapeToString = "msoTextEffectShapeCascadeDown"
        Case msoTextEffectShapeMixed: MsoPresetTextEffectShapeToString = "msoTextEffectShapeMixed"
    End Select
End Function
