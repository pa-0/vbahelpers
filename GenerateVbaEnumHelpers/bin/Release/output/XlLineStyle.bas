Attribute VB_Name = "wXlLineStyle"
Function XlLineStyleFromString(value As String) As XlLineStyle
    If IsNumeric(value) Then
        XlLineStyleFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlContinuous": XlLineStyleFromString = xlContinuous
        Case "xlDashDot": XlLineStyleFromString = xlDashDot
        Case "xlDashDotDot": XlLineStyleFromString = xlDashDotDot
        Case "xlSlantDashDot": XlLineStyleFromString = xlSlantDashDot
        Case "xlLineStyleNone": XlLineStyleFromString = xlLineStyleNone
        Case "xlDouble": XlLineStyleFromString = xlDouble
        Case "xlDot": XlLineStyleFromString = xlDot
        Case "xlDash": XlLineStyleFromString = xlDash
    End Select
End Function

Function XlLineStyleToString(value As XlLineStyle) As String
    Select Case value
        Case xlContinuous: XlLineStyleToString = "xlContinuous"
        Case xlDashDot: XlLineStyleToString = "xlDashDot"
        Case xlDashDotDot: XlLineStyleToString = "xlDashDotDot"
        Case xlSlantDashDot: XlLineStyleToString = "xlSlantDashDot"
        Case xlLineStyleNone: XlLineStyleToString = "xlLineStyleNone"
        Case xlDouble: XlLineStyleToString = "xlDouble"
        Case xlDot: XlLineStyleToString = "xlDot"
        Case xlDash: XlLineStyleToString = "xlDash"
    End Select
End Function
