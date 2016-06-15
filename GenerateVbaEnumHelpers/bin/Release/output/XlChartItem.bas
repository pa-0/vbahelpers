Attribute VB_Name = "wXlChartItem"
Function XlChartItemFromString(value As String) As XlChartItem
    If IsNumeric(value) Then
        XlChartItemFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlDataLabel": XlChartItemFromString = xlDataLabel
        Case "xlChartArea": XlChartItemFromString = xlChartArea
        Case "xlSeries": XlChartItemFromString = xlSeries
        Case "xlChartTitle": XlChartItemFromString = xlChartTitle
        Case "xlWalls": XlChartItemFromString = xlWalls
        Case "xlCorners": XlChartItemFromString = xlCorners
        Case "xlDataTable": XlChartItemFromString = xlDataTable
        Case "xlTrendline": XlChartItemFromString = xlTrendline
        Case "xlErrorBars": XlChartItemFromString = xlErrorBars
        Case "xlXErrorBars": XlChartItemFromString = xlXErrorBars
        Case "xlYErrorBars": XlChartItemFromString = xlYErrorBars
        Case "xlLegendEntry": XlChartItemFromString = xlLegendEntry
        Case "xlLegendKey": XlChartItemFromString = xlLegendKey
        Case "xlShape": XlChartItemFromString = xlShape
        Case "xlMajorGridlines": XlChartItemFromString = xlMajorGridlines
        Case "xlMinorGridlines": XlChartItemFromString = xlMinorGridlines
        Case "xlAxisTitle": XlChartItemFromString = xlAxisTitle
        Case "xlUpBars": XlChartItemFromString = xlUpBars
        Case "xlPlotArea": XlChartItemFromString = xlPlotArea
        Case "xlDownBars": XlChartItemFromString = xlDownBars
        Case "xlAxis": XlChartItemFromString = xlAxis
        Case "xlSeriesLines": XlChartItemFromString = xlSeriesLines
        Case "xlFloor": XlChartItemFromString = xlFloor
        Case "xlLegend": XlChartItemFromString = xlLegend
        Case "xlHiLoLines": XlChartItemFromString = xlHiLoLines
        Case "xlDropLines": XlChartItemFromString = xlDropLines
        Case "xlRadarAxisLabels": XlChartItemFromString = xlRadarAxisLabels
        Case "xlNothing": XlChartItemFromString = xlNothing
        Case "xlLeaderLines": XlChartItemFromString = xlLeaderLines
        Case "xlDisplayUnitLabel": XlChartItemFromString = xlDisplayUnitLabel
        Case "xlPivotChartFieldButton": XlChartItemFromString = xlPivotChartFieldButton
        Case "xlPivotChartDropZone": XlChartItemFromString = xlPivotChartDropZone
    End Select
End Function

Function XlChartItemToString(value As XlChartItem) As String
    Select Case value
        Case xlDataLabel: XlChartItemToString = "xlDataLabel"
        Case xlChartArea: XlChartItemToString = "xlChartArea"
        Case xlSeries: XlChartItemToString = "xlSeries"
        Case xlChartTitle: XlChartItemToString = "xlChartTitle"
        Case xlWalls: XlChartItemToString = "xlWalls"
        Case xlCorners: XlChartItemToString = "xlCorners"
        Case xlDataTable: XlChartItemToString = "xlDataTable"
        Case xlTrendline: XlChartItemToString = "xlTrendline"
        Case xlErrorBars: XlChartItemToString = "xlErrorBars"
        Case xlXErrorBars: XlChartItemToString = "xlXErrorBars"
        Case xlYErrorBars: XlChartItemToString = "xlYErrorBars"
        Case xlLegendEntry: XlChartItemToString = "xlLegendEntry"
        Case xlLegendKey: XlChartItemToString = "xlLegendKey"
        Case xlShape: XlChartItemToString = "xlShape"
        Case xlMajorGridlines: XlChartItemToString = "xlMajorGridlines"
        Case xlMinorGridlines: XlChartItemToString = "xlMinorGridlines"
        Case xlAxisTitle: XlChartItemToString = "xlAxisTitle"
        Case xlUpBars: XlChartItemToString = "xlUpBars"
        Case xlPlotArea: XlChartItemToString = "xlPlotArea"
        Case xlDownBars: XlChartItemToString = "xlDownBars"
        Case xlAxis: XlChartItemToString = "xlAxis"
        Case xlSeriesLines: XlChartItemToString = "xlSeriesLines"
        Case xlFloor: XlChartItemToString = "xlFloor"
        Case xlLegend: XlChartItemToString = "xlLegend"
        Case xlHiLoLines: XlChartItemToString = "xlHiLoLines"
        Case xlDropLines: XlChartItemToString = "xlDropLines"
        Case xlRadarAxisLabels: XlChartItemToString = "xlRadarAxisLabels"
        Case xlNothing: XlChartItemToString = "xlNothing"
        Case xlLeaderLines: XlChartItemToString = "xlLeaderLines"
        Case xlDisplayUnitLabel: XlChartItemToString = "xlDisplayUnitLabel"
        Case xlPivotChartFieldButton: XlChartItemToString = "xlPivotChartFieldButton"
        Case xlPivotChartDropZone: XlChartItemToString = "xlPivotChartDropZone"
    End Select
End Function
