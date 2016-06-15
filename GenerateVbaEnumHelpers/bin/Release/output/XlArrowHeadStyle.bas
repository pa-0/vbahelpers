Attribute VB_Name = "wXlArrowHeadStyle"
Function XlArrowHeadStyleFromString(value As String) As XlArrowHeadStyle
    If IsNumeric(value) Then
        XlArrowHeadStyleFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlArrowHeadStyleOpen": XlArrowHeadStyleFromString = xlArrowHeadStyleOpen
        Case "xlArrowHeadStyleClosed": XlArrowHeadStyleFromString = xlArrowHeadStyleClosed
        Case "xlArrowHeadStyleDoubleOpen": XlArrowHeadStyleFromString = xlArrowHeadStyleDoubleOpen
        Case "xlArrowHeadStyleDoubleClosed": XlArrowHeadStyleFromString = xlArrowHeadStyleDoubleClosed
        Case "xlArrowHeadStyleNone": XlArrowHeadStyleFromString = xlArrowHeadStyleNone
    End Select
End Function

Function XlArrowHeadStyleToString(value As XlArrowHeadStyle) As String
    Select Case value
        Case xlArrowHeadStyleOpen: XlArrowHeadStyleToString = "xlArrowHeadStyleOpen"
        Case xlArrowHeadStyleClosed: XlArrowHeadStyleToString = "xlArrowHeadStyleClosed"
        Case xlArrowHeadStyleDoubleOpen: XlArrowHeadStyleToString = "xlArrowHeadStyleDoubleOpen"
        Case xlArrowHeadStyleDoubleClosed: XlArrowHeadStyleToString = "xlArrowHeadStyleDoubleClosed"
        Case xlArrowHeadStyleNone: XlArrowHeadStyleToString = "xlArrowHeadStyleNone"
    End Select
End Function
