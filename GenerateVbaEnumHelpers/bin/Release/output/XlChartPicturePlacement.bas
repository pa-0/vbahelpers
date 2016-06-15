Attribute VB_Name = "wXlChartPicturePlacement"
Function XlChartPicturePlacementFromString(value As String) As XlChartPicturePlacement
    If IsNumeric(value) Then
        XlChartPicturePlacementFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlSides": XlChartPicturePlacementFromString = xlSides
        Case "xlEnd": XlChartPicturePlacementFromString = xlEnd
        Case "xlEndSides": XlChartPicturePlacementFromString = xlEndSides
        Case "xlFront": XlChartPicturePlacementFromString = xlFront
        Case "xlFrontSides": XlChartPicturePlacementFromString = xlFrontSides
        Case "xlFrontEnd": XlChartPicturePlacementFromString = xlFrontEnd
        Case "xlAllFaces": XlChartPicturePlacementFromString = xlAllFaces
    End Select
End Function

Function XlChartPicturePlacementToString(value As XlChartPicturePlacement) As String
    Select Case value
        Case xlSides: XlChartPicturePlacementToString = "xlSides"
        Case xlEnd: XlChartPicturePlacementToString = "xlEnd"
        Case xlEndSides: XlChartPicturePlacementToString = "xlEndSides"
        Case xlFront: XlChartPicturePlacementToString = "xlFront"
        Case xlFrontSides: XlChartPicturePlacementToString = "xlFrontSides"
        Case xlFrontEnd: XlChartPicturePlacementToString = "xlFrontEnd"
        Case xlAllFaces: XlChartPicturePlacementToString = "xlAllFaces"
    End Select
End Function
