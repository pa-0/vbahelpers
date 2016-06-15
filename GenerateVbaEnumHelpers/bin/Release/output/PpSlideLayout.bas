Attribute VB_Name = "wPpSlideLayout"
Function PpSlideLayoutFromString(value As String) As PpSlideLayout
    If IsNumeric(value) Then
        PpSlideLayoutFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "ppLayoutTitle": PpSlideLayoutFromString = ppLayoutTitle
        Case "ppLayoutText": PpSlideLayoutFromString = ppLayoutText
        Case "ppLayoutTwoColumnText": PpSlideLayoutFromString = ppLayoutTwoColumnText
        Case "ppLayoutTable": PpSlideLayoutFromString = ppLayoutTable
        Case "ppLayoutTextAndChart": PpSlideLayoutFromString = ppLayoutTextAndChart
        Case "ppLayoutChartAndText": PpSlideLayoutFromString = ppLayoutChartAndText
        Case "ppLayoutOrgchart": PpSlideLayoutFromString = ppLayoutOrgchart
        Case "ppLayoutChart": PpSlideLayoutFromString = ppLayoutChart
        Case "ppLayoutTextAndClipart": PpSlideLayoutFromString = ppLayoutTextAndClipart
        Case "ppLayoutClipartAndText": PpSlideLayoutFromString = ppLayoutClipartAndText
        Case "ppLayoutTitleOnly": PpSlideLayoutFromString = ppLayoutTitleOnly
        Case "ppLayoutBlank": PpSlideLayoutFromString = ppLayoutBlank
        Case "ppLayoutTextAndObject": PpSlideLayoutFromString = ppLayoutTextAndObject
        Case "ppLayoutObjectAndText": PpSlideLayoutFromString = ppLayoutObjectAndText
        Case "ppLayoutLargeObject": PpSlideLayoutFromString = ppLayoutLargeObject
        Case "ppLayoutObject": PpSlideLayoutFromString = ppLayoutObject
        Case "ppLayoutTextAndMediaClip": PpSlideLayoutFromString = ppLayoutTextAndMediaClip
        Case "ppLayoutMediaClipAndText": PpSlideLayoutFromString = ppLayoutMediaClipAndText
        Case "ppLayoutObjectOverText": PpSlideLayoutFromString = ppLayoutObjectOverText
        Case "ppLayoutTextOverObject": PpSlideLayoutFromString = ppLayoutTextOverObject
        Case "ppLayoutTextAndTwoObjects": PpSlideLayoutFromString = ppLayoutTextAndTwoObjects
        Case "ppLayoutTwoObjectsAndText": PpSlideLayoutFromString = ppLayoutTwoObjectsAndText
        Case "ppLayoutTwoObjectsOverText": PpSlideLayoutFromString = ppLayoutTwoObjectsOverText
        Case "ppLayoutFourObjects": PpSlideLayoutFromString = ppLayoutFourObjects
        Case "ppLayoutVerticalText": PpSlideLayoutFromString = ppLayoutVerticalText
        Case "ppLayoutClipArtAndVerticalText": PpSlideLayoutFromString = ppLayoutClipArtAndVerticalText
        Case "ppLayoutVerticalTitleAndText": PpSlideLayoutFromString = ppLayoutVerticalTitleAndText
        Case "ppLayoutVerticalTitleAndTextOverChart": PpSlideLayoutFromString = ppLayoutVerticalTitleAndTextOverChart
        Case "ppLayoutTwoObjects": PpSlideLayoutFromString = ppLayoutTwoObjects
        Case "ppLayoutObjectAndTwoObjects": PpSlideLayoutFromString = ppLayoutObjectAndTwoObjects
        Case "ppLayoutTwoObjectsAndObject": PpSlideLayoutFromString = ppLayoutTwoObjectsAndObject
        Case "ppLayoutCustom": PpSlideLayoutFromString = ppLayoutCustom
        Case "ppLayoutSectionHeader": PpSlideLayoutFromString = ppLayoutSectionHeader
        Case "ppLayoutComparison": PpSlideLayoutFromString = ppLayoutComparison
        Case "ppLayoutContentWithCaption": PpSlideLayoutFromString = ppLayoutContentWithCaption
        Case "ppLayoutPictureWithCaption": PpSlideLayoutFromString = ppLayoutPictureWithCaption
        Case "ppLayoutMixed": PpSlideLayoutFromString = ppLayoutMixed
    End Select
End Function

Function PpSlideLayoutToString(value As PpSlideLayout) As String
    Select Case value
        Case ppLayoutTitle: PpSlideLayoutToString = "ppLayoutTitle"
        Case ppLayoutText: PpSlideLayoutToString = "ppLayoutText"
        Case ppLayoutTwoColumnText: PpSlideLayoutToString = "ppLayoutTwoColumnText"
        Case ppLayoutTable: PpSlideLayoutToString = "ppLayoutTable"
        Case ppLayoutTextAndChart: PpSlideLayoutToString = "ppLayoutTextAndChart"
        Case ppLayoutChartAndText: PpSlideLayoutToString = "ppLayoutChartAndText"
        Case ppLayoutOrgchart: PpSlideLayoutToString = "ppLayoutOrgchart"
        Case ppLayoutChart: PpSlideLayoutToString = "ppLayoutChart"
        Case ppLayoutTextAndClipart: PpSlideLayoutToString = "ppLayoutTextAndClipart"
        Case ppLayoutClipartAndText: PpSlideLayoutToString = "ppLayoutClipartAndText"
        Case ppLayoutTitleOnly: PpSlideLayoutToString = "ppLayoutTitleOnly"
        Case ppLayoutBlank: PpSlideLayoutToString = "ppLayoutBlank"
        Case ppLayoutTextAndObject: PpSlideLayoutToString = "ppLayoutTextAndObject"
        Case ppLayoutObjectAndText: PpSlideLayoutToString = "ppLayoutObjectAndText"
        Case ppLayoutLargeObject: PpSlideLayoutToString = "ppLayoutLargeObject"
        Case ppLayoutObject: PpSlideLayoutToString = "ppLayoutObject"
        Case ppLayoutTextAndMediaClip: PpSlideLayoutToString = "ppLayoutTextAndMediaClip"
        Case ppLayoutMediaClipAndText: PpSlideLayoutToString = "ppLayoutMediaClipAndText"
        Case ppLayoutObjectOverText: PpSlideLayoutToString = "ppLayoutObjectOverText"
        Case ppLayoutTextOverObject: PpSlideLayoutToString = "ppLayoutTextOverObject"
        Case ppLayoutTextAndTwoObjects: PpSlideLayoutToString = "ppLayoutTextAndTwoObjects"
        Case ppLayoutTwoObjectsAndText: PpSlideLayoutToString = "ppLayoutTwoObjectsAndText"
        Case ppLayoutTwoObjectsOverText: PpSlideLayoutToString = "ppLayoutTwoObjectsOverText"
        Case ppLayoutFourObjects: PpSlideLayoutToString = "ppLayoutFourObjects"
        Case ppLayoutVerticalText: PpSlideLayoutToString = "ppLayoutVerticalText"
        Case ppLayoutClipArtAndVerticalText: PpSlideLayoutToString = "ppLayoutClipArtAndVerticalText"
        Case ppLayoutVerticalTitleAndText: PpSlideLayoutToString = "ppLayoutVerticalTitleAndText"
        Case ppLayoutVerticalTitleAndTextOverChart: PpSlideLayoutToString = "ppLayoutVerticalTitleAndTextOverChart"
        Case ppLayoutTwoObjects: PpSlideLayoutToString = "ppLayoutTwoObjects"
        Case ppLayoutObjectAndTwoObjects: PpSlideLayoutToString = "ppLayoutObjectAndTwoObjects"
        Case ppLayoutTwoObjectsAndObject: PpSlideLayoutToString = "ppLayoutTwoObjectsAndObject"
        Case ppLayoutCustom: PpSlideLayoutToString = "ppLayoutCustom"
        Case ppLayoutSectionHeader: PpSlideLayoutToString = "ppLayoutSectionHeader"
        Case ppLayoutComparison: PpSlideLayoutToString = "ppLayoutComparison"
        Case ppLayoutContentWithCaption: PpSlideLayoutToString = "ppLayoutContentWithCaption"
        Case ppLayoutPictureWithCaption: PpSlideLayoutToString = "ppLayoutPictureWithCaption"
        Case ppLayoutMixed: PpSlideLayoutToString = "ppLayoutMixed"
    End Select
End Function
