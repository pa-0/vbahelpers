Attribute VB_Name = "wXlAboveBelow"
Function XlAboveBelowFromString(value As String) As XlAboveBelow
    If IsNumeric(value) Then
        XlAboveBelowFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlAboveAverage": XlAboveBelowFromString = xlAboveAverage
        Case "xlBelowAverage": XlAboveBelowFromString = xlBelowAverage
        Case "xlEqualAboveAverage": XlAboveBelowFromString = xlEqualAboveAverage
        Case "xlEqualBelowAverage": XlAboveBelowFromString = xlEqualBelowAverage
        Case "xlAboveStdDev": XlAboveBelowFromString = xlAboveStdDev
        Case "xlBelowStdDev": XlAboveBelowFromString = xlBelowStdDev
    End Select
End Function

Function XlAboveBelowToString(value As XlAboveBelow) As String
    Select Case value
        Case xlAboveAverage: XlAboveBelowToString = "xlAboveAverage"
        Case xlBelowAverage: XlAboveBelowToString = "xlBelowAverage"
        Case xlEqualAboveAverage: XlAboveBelowToString = "xlEqualAboveAverage"
        Case xlEqualBelowAverage: XlAboveBelowToString = "xlEqualBelowAverage"
        Case xlAboveStdDev: XlAboveBelowToString = "xlAboveStdDev"
        Case xlBelowStdDev: XlAboveBelowToString = "xlBelowStdDev"
    End Select
End Function
