Attribute VB_Name = "wXlErrorBarDirection"
Function XlErrorBarDirectionFromString(value As String) As XlErrorBarDirection
    If IsNumeric(value) Then
        XlErrorBarDirectionFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlChartY": XlErrorBarDirectionFromString = xlChartY
        Case "xlChartX": XlErrorBarDirectionFromString = xlChartX
    End Select
End Function

Function XlErrorBarDirectionToString(value As XlErrorBarDirection) As String
    Select Case value
        Case xlChartY: XlErrorBarDirectionToString = "xlChartY"
        Case xlChartX: XlErrorBarDirectionToString = "xlChartX"
    End Select
End Function
