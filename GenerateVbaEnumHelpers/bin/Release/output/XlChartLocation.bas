Attribute VB_Name = "wXlChartLocation"
Function XlChartLocationFromString(value As String) As XlChartLocation
    If IsNumeric(value) Then
        XlChartLocationFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlLocationAsNewSheet": XlChartLocationFromString = xlLocationAsNewSheet
        Case "xlLocationAsObject": XlChartLocationFromString = xlLocationAsObject
        Case "xlLocationAutomatic": XlChartLocationFromString = xlLocationAutomatic
    End Select
End Function

Function XlChartLocationToString(value As XlChartLocation) As String
    Select Case value
        Case xlLocationAsNewSheet: XlChartLocationToString = "xlLocationAsNewSheet"
        Case xlLocationAsObject: XlChartLocationToString = "xlLocationAsObject"
        Case xlLocationAutomatic: XlChartLocationToString = "xlLocationAutomatic"
    End Select
End Function
