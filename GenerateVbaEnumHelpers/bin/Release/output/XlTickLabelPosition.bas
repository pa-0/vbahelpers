Attribute VB_Name = "wXlTickLabelPosition"
Function XlTickLabelPositionFromString(value As String) As XlTickLabelPosition
    If IsNumeric(value) Then
        XlTickLabelPositionFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlTickLabelPositionNextToAxis": XlTickLabelPositionFromString = xlTickLabelPositionNextToAxis
        Case "xlTickLabelPositionNone": XlTickLabelPositionFromString = xlTickLabelPositionNone
        Case "xlTickLabelPositionLow": XlTickLabelPositionFromString = xlTickLabelPositionLow
        Case "xlTickLabelPositionHigh": XlTickLabelPositionFromString = xlTickLabelPositionHigh
    End Select
End Function

Function XlTickLabelPositionToString(value As XlTickLabelPosition) As String
    Select Case value
        Case xlTickLabelPositionNextToAxis: XlTickLabelPositionToString = "xlTickLabelPositionNextToAxis"
        Case xlTickLabelPositionNone: XlTickLabelPositionToString = "xlTickLabelPositionNone"
        Case xlTickLabelPositionLow: XlTickLabelPositionToString = "xlTickLabelPositionLow"
        Case xlTickLabelPositionHigh: XlTickLabelPositionToString = "xlTickLabelPositionHigh"
    End Select
End Function
