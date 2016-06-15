Attribute VB_Name = "wXlChartElementPosition"
Function XlChartElementPositionFromString(value As String) As XlChartElementPosition
    If IsNumeric(value) Then
        XlChartElementPositionFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlChartElementPositionCustom": XlChartElementPositionFromString = xlChartElementPositionCustom
        Case "xlChartElementPositionAutomatic": XlChartElementPositionFromString = xlChartElementPositionAutomatic
    End Select
End Function

Function XlChartElementPositionToString(value As XlChartElementPosition) As String
    Select Case value
        Case xlChartElementPositionCustom: XlChartElementPositionToString = "xlChartElementPositionCustom"
        Case xlChartElementPositionAutomatic: XlChartElementPositionToString = "xlChartElementPositionAutomatic"
    End Select
End Function
