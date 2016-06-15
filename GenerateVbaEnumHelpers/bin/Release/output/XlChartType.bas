Attribute VB_Name = "wXlChartType"
Function XlChartTypeFromString(value As String) As XlChartType
    If IsNumeric(value) Then
        XlChartTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlArea": XlChartTypeFromString = xlArea
        Case "xlLine": XlChartTypeFromString = xlLine
        Case "xlPie": XlChartTypeFromString = xlPie
        Case "xlBubble": XlChartTypeFromString = xlBubble
        Case "xlColumnClustered": XlChartTypeFromString = xlColumnClustered
        Case "xlColumnStacked": XlChartTypeFromString = xlColumnStacked
        Case "xlColumnStacked100": XlChartTypeFromString = xlColumnStacked100
        Case "xl3DColumnClustered": XlChartTypeFromString = xl3DColumnClustered
        Case "xl3DColumnStacked": XlChartTypeFromString = xl3DColumnStacked
        Case "xl3DColumnStacked100": XlChartTypeFromString = xl3DColumnStacked100
        Case "xlBarClustered": XlChartTypeFromString = xlBarClustered
        Case "xlBarStacked": XlChartTypeFromString = xlBarStacked
        Case "xlBarStacked100": XlChartTypeFromString = xlBarStacked100
        Case "xl3DBarClustered": XlChartTypeFromString = xl3DBarClustered
        Case "xl3DBarStacked": XlChartTypeFromString = xl3DBarStacked
        Case "xl3DBarStacked100": XlChartTypeFromString = xl3DBarStacked100
        Case "xlLineStacked": XlChartTypeFromString = xlLineStacked
        Case "xlLineStacked100": XlChartTypeFromString = xlLineStacked100
        Case "xlLineMarkers": XlChartTypeFromString = xlLineMarkers
        Case "xlLineMarkersStacked": XlChartTypeFromString = xlLineMarkersStacked
        Case "xlLineMarkersStacked100": XlChartTypeFromString = xlLineMarkersStacked100
        Case "xlPieOfPie": XlChartTypeFromString = xlPieOfPie
        Case "xlPieExploded": XlChartTypeFromString = xlPieExploded
        Case "xl3DPieExploded": XlChartTypeFromString = xl3DPieExploded
        Case "xlBarOfPie": XlChartTypeFromString = xlBarOfPie
        Case "xlXYScatterSmooth": XlChartTypeFromString = xlXYScatterSmooth
        Case "xlXYScatterSmoothNoMarkers": XlChartTypeFromString = xlXYScatterSmoothNoMarkers
        Case "xlXYScatterLines": XlChartTypeFromString = xlXYScatterLines
        Case "xlXYScatterLinesNoMarkers": XlChartTypeFromString = xlXYScatterLinesNoMarkers
        Case "xlAreaStacked": XlChartTypeFromString = xlAreaStacked
        Case "xlAreaStacked100": XlChartTypeFromString = xlAreaStacked100
        Case "xl3DAreaStacked": XlChartTypeFromString = xl3DAreaStacked
        Case "xl3DAreaStacked100": XlChartTypeFromString = xl3DAreaStacked100
        Case "xlDoughnutExploded": XlChartTypeFromString = xlDoughnutExploded
        Case "xlRadarMarkers": XlChartTypeFromString = xlRadarMarkers
        Case "xlRadarFilled": XlChartTypeFromString = xlRadarFilled
        Case "xlSurface": XlChartTypeFromString = xlSurface
        Case "xlSurfaceWireframe": XlChartTypeFromString = xlSurfaceWireframe
        Case "xlSurfaceTopView": XlChartTypeFromString = xlSurfaceTopView
        Case "xlSurfaceTopViewWireframe": XlChartTypeFromString = xlSurfaceTopViewWireframe
        Case "xlBubble3DEffect": XlChartTypeFromString = xlBubble3DEffect
        Case "xlStockHLC": XlChartTypeFromString = xlStockHLC
        Case "xlStockOHLC": XlChartTypeFromString = xlStockOHLC
        Case "xlStockVHLC": XlChartTypeFromString = xlStockVHLC
        Case "xlStockVOHLC": XlChartTypeFromString = xlStockVOHLC
        Case "xlCylinderColClustered": XlChartTypeFromString = xlCylinderColClustered
        Case "xlCylinderColStacked": XlChartTypeFromString = xlCylinderColStacked
        Case "xlCylinderColStacked100": XlChartTypeFromString = xlCylinderColStacked100
        Case "xlCylinderBarClustered": XlChartTypeFromString = xlCylinderBarClustered
        Case "xlCylinderBarStacked": XlChartTypeFromString = xlCylinderBarStacked
        Case "xlCylinderBarStacked100": XlChartTypeFromString = xlCylinderBarStacked100
        Case "xlCylinderCol": XlChartTypeFromString = xlCylinderCol
        Case "xlConeColClustered": XlChartTypeFromString = xlConeColClustered
        Case "xlConeColStacked": XlChartTypeFromString = xlConeColStacked
        Case "xlConeColStacked100": XlChartTypeFromString = xlConeColStacked100
        Case "xlConeBarClustered": XlChartTypeFromString = xlConeBarClustered
        Case "xlConeBarStacked": XlChartTypeFromString = xlConeBarStacked
        Case "xlConeBarStacked100": XlChartTypeFromString = xlConeBarStacked100
        Case "xlConeCol": XlChartTypeFromString = xlConeCol
        Case "xlPyramidColClustered": XlChartTypeFromString = xlPyramidColClustered
        Case "xlPyramidColStacked": XlChartTypeFromString = xlPyramidColStacked
        Case "xlPyramidColStacked100": XlChartTypeFromString = xlPyramidColStacked100
        Case "xlPyramidBarClustered": XlChartTypeFromString = xlPyramidBarClustered
        Case "xlPyramidBarStacked": XlChartTypeFromString = xlPyramidBarStacked
        Case "xlPyramidBarStacked100": XlChartTypeFromString = xlPyramidBarStacked100
        Case "xlPyramidCol": XlChartTypeFromString = xlPyramidCol
        Case "xlXYScatter": XlChartTypeFromString = xlXYScatter
        Case "xlRadar": XlChartTypeFromString = xlRadar
        Case "xlDoughnut": XlChartTypeFromString = xlDoughnut
        Case "xl3DPie": XlChartTypeFromString = xl3DPie
        Case "xl3DLine": XlChartTypeFromString = xl3DLine
        Case "xl3DColumn": XlChartTypeFromString = xl3DColumn
        Case "xl3DArea": XlChartTypeFromString = xl3DArea
    End Select
End Function

Function XlChartTypeToString(value As XlChartType) As String
    Select Case value
        Case xlArea: XlChartTypeToString = "xlArea"
        Case xlLine: XlChartTypeToString = "xlLine"
        Case xlPie: XlChartTypeToString = "xlPie"
        Case xlBubble: XlChartTypeToString = "xlBubble"
        Case xlColumnClustered: XlChartTypeToString = "xlColumnClustered"
        Case xlColumnStacked: XlChartTypeToString = "xlColumnStacked"
        Case xlColumnStacked100: XlChartTypeToString = "xlColumnStacked100"
        Case xl3DColumnClustered: XlChartTypeToString = "xl3DColumnClustered"
        Case xl3DColumnStacked: XlChartTypeToString = "xl3DColumnStacked"
        Case xl3DColumnStacked100: XlChartTypeToString = "xl3DColumnStacked100"
        Case xlBarClustered: XlChartTypeToString = "xlBarClustered"
        Case xlBarStacked: XlChartTypeToString = "xlBarStacked"
        Case xlBarStacked100: XlChartTypeToString = "xlBarStacked100"
        Case xl3DBarClustered: XlChartTypeToString = "xl3DBarClustered"
        Case xl3DBarStacked: XlChartTypeToString = "xl3DBarStacked"
        Case xl3DBarStacked100: XlChartTypeToString = "xl3DBarStacked100"
        Case xlLineStacked: XlChartTypeToString = "xlLineStacked"
        Case xlLineStacked100: XlChartTypeToString = "xlLineStacked100"
        Case xlLineMarkers: XlChartTypeToString = "xlLineMarkers"
        Case xlLineMarkersStacked: XlChartTypeToString = "xlLineMarkersStacked"
        Case xlLineMarkersStacked100: XlChartTypeToString = "xlLineMarkersStacked100"
        Case xlPieOfPie: XlChartTypeToString = "xlPieOfPie"
        Case xlPieExploded: XlChartTypeToString = "xlPieExploded"
        Case xl3DPieExploded: XlChartTypeToString = "xl3DPieExploded"
        Case xlBarOfPie: XlChartTypeToString = "xlBarOfPie"
        Case xlXYScatterSmooth: XlChartTypeToString = "xlXYScatterSmooth"
        Case xlXYScatterSmoothNoMarkers: XlChartTypeToString = "xlXYScatterSmoothNoMarkers"
        Case xlXYScatterLines: XlChartTypeToString = "xlXYScatterLines"
        Case xlXYScatterLinesNoMarkers: XlChartTypeToString = "xlXYScatterLinesNoMarkers"
        Case xlAreaStacked: XlChartTypeToString = "xlAreaStacked"
        Case xlAreaStacked100: XlChartTypeToString = "xlAreaStacked100"
        Case xl3DAreaStacked: XlChartTypeToString = "xl3DAreaStacked"
        Case xl3DAreaStacked100: XlChartTypeToString = "xl3DAreaStacked100"
        Case xlDoughnutExploded: XlChartTypeToString = "xlDoughnutExploded"
        Case xlRadarMarkers: XlChartTypeToString = "xlRadarMarkers"
        Case xlRadarFilled: XlChartTypeToString = "xlRadarFilled"
        Case xlSurface: XlChartTypeToString = "xlSurface"
        Case xlSurfaceWireframe: XlChartTypeToString = "xlSurfaceWireframe"
        Case xlSurfaceTopView: XlChartTypeToString = "xlSurfaceTopView"
        Case xlSurfaceTopViewWireframe: XlChartTypeToString = "xlSurfaceTopViewWireframe"
        Case xlBubble3DEffect: XlChartTypeToString = "xlBubble3DEffect"
        Case xlStockHLC: XlChartTypeToString = "xlStockHLC"
        Case xlStockOHLC: XlChartTypeToString = "xlStockOHLC"
        Case xlStockVHLC: XlChartTypeToString = "xlStockVHLC"
        Case xlStockVOHLC: XlChartTypeToString = "xlStockVOHLC"
        Case xlCylinderColClustered: XlChartTypeToString = "xlCylinderColClustered"
        Case xlCylinderColStacked: XlChartTypeToString = "xlCylinderColStacked"
        Case xlCylinderColStacked100: XlChartTypeToString = "xlCylinderColStacked100"
        Case xlCylinderBarClustered: XlChartTypeToString = "xlCylinderBarClustered"
        Case xlCylinderBarStacked: XlChartTypeToString = "xlCylinderBarStacked"
        Case xlCylinderBarStacked100: XlChartTypeToString = "xlCylinderBarStacked100"
        Case xlCylinderCol: XlChartTypeToString = "xlCylinderCol"
        Case xlConeColClustered: XlChartTypeToString = "xlConeColClustered"
        Case xlConeColStacked: XlChartTypeToString = "xlConeColStacked"
        Case xlConeColStacked100: XlChartTypeToString = "xlConeColStacked100"
        Case xlConeBarClustered: XlChartTypeToString = "xlConeBarClustered"
        Case xlConeBarStacked: XlChartTypeToString = "xlConeBarStacked"
        Case xlConeBarStacked100: XlChartTypeToString = "xlConeBarStacked100"
        Case xlConeCol: XlChartTypeToString = "xlConeCol"
        Case xlPyramidColClustered: XlChartTypeToString = "xlPyramidColClustered"
        Case xlPyramidColStacked: XlChartTypeToString = "xlPyramidColStacked"
        Case xlPyramidColStacked100: XlChartTypeToString = "xlPyramidColStacked100"
        Case xlPyramidBarClustered: XlChartTypeToString = "xlPyramidBarClustered"
        Case xlPyramidBarStacked: XlChartTypeToString = "xlPyramidBarStacked"
        Case xlPyramidBarStacked100: XlChartTypeToString = "xlPyramidBarStacked100"
        Case xlPyramidCol: XlChartTypeToString = "xlPyramidCol"
        Case xlXYScatter: XlChartTypeToString = "xlXYScatter"
        Case xlRadar: XlChartTypeToString = "xlRadar"
        Case xlDoughnut: XlChartTypeToString = "xlDoughnut"
        Case xl3DPie: XlChartTypeToString = "xl3DPie"
        Case xl3DLine: XlChartTypeToString = "xl3DLine"
        Case xl3DColumn: XlChartTypeToString = "xl3DColumn"
        Case xl3DArea: XlChartTypeToString = "xl3DArea"
    End Select
End Function