Attribute VB_Name = "wXlTotalsCalculation"
Function XlTotalsCalculationFromString(value As String) As XlTotalsCalculation
    If IsNumeric(value) Then
        XlTotalsCalculationFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlTotalsCalculationNone": XlTotalsCalculationFromString = xlTotalsCalculationNone
        Case "xlTotalsCalculationSum": XlTotalsCalculationFromString = xlTotalsCalculationSum
        Case "xlTotalsCalculationAverage": XlTotalsCalculationFromString = xlTotalsCalculationAverage
        Case "xlTotalsCalculationCount": XlTotalsCalculationFromString = xlTotalsCalculationCount
        Case "xlTotalsCalculationCountNums": XlTotalsCalculationFromString = xlTotalsCalculationCountNums
        Case "xlTotalsCalculationMin": XlTotalsCalculationFromString = xlTotalsCalculationMin
        Case "xlTotalsCalculationMax": XlTotalsCalculationFromString = xlTotalsCalculationMax
        Case "xlTotalsCalculationStdDev": XlTotalsCalculationFromString = xlTotalsCalculationStdDev
        Case "xlTotalsCalculationVar": XlTotalsCalculationFromString = xlTotalsCalculationVar
        Case "xlTotalsCalculationCustom": XlTotalsCalculationFromString = xlTotalsCalculationCustom
    End Select
End Function

Function XlTotalsCalculationToString(value As XlTotalsCalculation) As String
    Select Case value
        Case xlTotalsCalculationNone: XlTotalsCalculationToString = "xlTotalsCalculationNone"
        Case xlTotalsCalculationSum: XlTotalsCalculationToString = "xlTotalsCalculationSum"
        Case xlTotalsCalculationAverage: XlTotalsCalculationToString = "xlTotalsCalculationAverage"
        Case xlTotalsCalculationCount: XlTotalsCalculationToString = "xlTotalsCalculationCount"
        Case xlTotalsCalculationCountNums: XlTotalsCalculationToString = "xlTotalsCalculationCountNums"
        Case xlTotalsCalculationMin: XlTotalsCalculationToString = "xlTotalsCalculationMin"
        Case xlTotalsCalculationMax: XlTotalsCalculationToString = "xlTotalsCalculationMax"
        Case xlTotalsCalculationStdDev: XlTotalsCalculationToString = "xlTotalsCalculationStdDev"
        Case xlTotalsCalculationVar: XlTotalsCalculationToString = "xlTotalsCalculationVar"
        Case xlTotalsCalculationCustom: XlTotalsCalculationToString = "xlTotalsCalculationCustom"
    End Select
End Function
