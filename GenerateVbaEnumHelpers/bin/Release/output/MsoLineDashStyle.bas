Attribute VB_Name = "wMsoLineDashStyle"
Function MsoLineDashStyleFromString(value As String) As MsoLineDashStyle
    If IsNumeric(value) Then
        MsoLineDashStyleFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoLineSolid": MsoLineDashStyleFromString = msoLineSolid
        Case "msoLineSquareDot": MsoLineDashStyleFromString = msoLineSquareDot
        Case "msoLineRoundDot": MsoLineDashStyleFromString = msoLineRoundDot
        Case "msoLineDash": MsoLineDashStyleFromString = msoLineDash
        Case "msoLineDashDot": MsoLineDashStyleFromString = msoLineDashDot
        Case "msoLineDashDotDot": MsoLineDashStyleFromString = msoLineDashDotDot
        Case "msoLineLongDash": MsoLineDashStyleFromString = msoLineLongDash
        Case "msoLineLongDashDot": MsoLineDashStyleFromString = msoLineLongDashDot
        Case "msoLineLongDashDotDot": MsoLineDashStyleFromString = msoLineLongDashDotDot
        Case "msoLineSysDash": MsoLineDashStyleFromString = msoLineSysDash
        Case "msoLineSysDot": MsoLineDashStyleFromString = msoLineSysDot
        Case "msoLineSysDashDot": MsoLineDashStyleFromString = msoLineSysDashDot
        Case "msoLineDashStyleMixed": MsoLineDashStyleFromString = msoLineDashStyleMixed
    End Select
End Function

Function MsoLineDashStyleToString(value As MsoLineDashStyle) As String
    Select Case value
        Case msoLineSolid: MsoLineDashStyleToString = "msoLineSolid"
        Case msoLineSquareDot: MsoLineDashStyleToString = "msoLineSquareDot"
        Case msoLineRoundDot: MsoLineDashStyleToString = "msoLineRoundDot"
        Case msoLineDash: MsoLineDashStyleToString = "msoLineDash"
        Case msoLineDashDot: MsoLineDashStyleToString = "msoLineDashDot"
        Case msoLineDashDotDot: MsoLineDashStyleToString = "msoLineDashDotDot"
        Case msoLineLongDash: MsoLineDashStyleToString = "msoLineLongDash"
        Case msoLineLongDashDot: MsoLineDashStyleToString = "msoLineLongDashDot"
        Case msoLineLongDashDotDot: MsoLineDashStyleToString = "msoLineLongDashDotDot"
        Case msoLineSysDash: MsoLineDashStyleToString = "msoLineSysDash"
        Case msoLineSysDot: MsoLineDashStyleToString = "msoLineSysDot"
        Case msoLineSysDashDot: MsoLineDashStyleToString = "msoLineSysDashDot"
        Case msoLineDashStyleMixed: MsoLineDashStyleToString = "msoLineDashStyleMixed"
    End Select
End Function
