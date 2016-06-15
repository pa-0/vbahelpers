Attribute VB_Name = "wXlTrendlineType"
Function XlTrendlineTypeFromString(value As String) As XlTrendlineType
    If IsNumeric(value) Then
        XlTrendlineTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlPolynomial": XlTrendlineTypeFromString = xlPolynomial
        Case "xlPower": XlTrendlineTypeFromString = xlPower
        Case "xlExponential": XlTrendlineTypeFromString = xlExponential
        Case "xlMovingAvg": XlTrendlineTypeFromString = xlMovingAvg
        Case "xlLogarithmic": XlTrendlineTypeFromString = xlLogarithmic
        Case "xlLinear": XlTrendlineTypeFromString = xlLinear
    End Select
End Function

Function XlTrendlineTypeToString(value As XlTrendlineType) As String
    Select Case value
        Case xlPolynomial: XlTrendlineTypeToString = "xlPolynomial"
        Case xlPower: XlTrendlineTypeToString = "xlPower"
        Case xlExponential: XlTrendlineTypeToString = "xlExponential"
        Case xlMovingAvg: XlTrendlineTypeToString = "xlMovingAvg"
        Case xlLogarithmic: XlTrendlineTypeToString = "xlLogarithmic"
        Case xlLinear: XlTrendlineTypeToString = "xlLinear"
    End Select
End Function
