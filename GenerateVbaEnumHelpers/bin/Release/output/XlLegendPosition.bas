Attribute VB_Name = "wXlLegendPosition"
Function XlLegendPositionFromString(value As String) As XlLegendPosition
    If IsNumeric(value) Then
        XlLegendPositionFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlLegendPositionCorner": XlLegendPositionFromString = xlLegendPositionCorner
        Case "xlLegendPositionCustom": XlLegendPositionFromString = xlLegendPositionCustom
        Case "xlLegendPositionTop": XlLegendPositionFromString = xlLegendPositionTop
        Case "xlLegendPositionRight": XlLegendPositionFromString = xlLegendPositionRight
        Case "xlLegendPositionLeft": XlLegendPositionFromString = xlLegendPositionLeft
        Case "xlLegendPositionBottom": XlLegendPositionFromString = xlLegendPositionBottom
    End Select
End Function

Function XlLegendPositionToString(value As XlLegendPosition) As String
    Select Case value
        Case xlLegendPositionCorner: XlLegendPositionToString = "xlLegendPositionCorner"
        Case xlLegendPositionCustom: XlLegendPositionToString = "xlLegendPositionCustom"
        Case xlLegendPositionTop: XlLegendPositionToString = "xlLegendPositionTop"
        Case xlLegendPositionRight: XlLegendPositionToString = "xlLegendPositionRight"
        Case xlLegendPositionLeft: XlLegendPositionToString = "xlLegendPositionLeft"
        Case xlLegendPositionBottom: XlLegendPositionToString = "xlLegendPositionBottom"
    End Select
End Function
