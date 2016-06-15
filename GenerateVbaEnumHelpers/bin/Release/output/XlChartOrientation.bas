Attribute VB_Name = "wXlChartOrientation"
Function XlChartOrientationFromString(value As String) As XlChartOrientation
    If IsNumeric(value) Then
        XlChartOrientationFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlUpward": XlChartOrientationFromString = xlUpward
        Case "xlDownward": XlChartOrientationFromString = xlDownward
        Case "xlVertical": XlChartOrientationFromString = xlVertical
        Case "xlHorizontal": XlChartOrientationFromString = xlHorizontal
    End Select
End Function

Function XlChartOrientationToString(value As XlChartOrientation) As String
    Select Case value
        Case xlUpward: XlChartOrientationToString = "xlUpward"
        Case xlDownward: XlChartOrientationToString = "xlDownward"
        Case xlVertical: XlChartOrientationToString = "xlVertical"
        Case xlHorizontal: XlChartOrientationToString = "xlHorizontal"
    End Select
End Function
