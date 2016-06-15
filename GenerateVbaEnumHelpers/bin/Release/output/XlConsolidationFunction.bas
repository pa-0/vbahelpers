Attribute VB_Name = "wXlConsolidationFunction"
Function XlConsolidationFunctionFromString(value As String) As XlConsolidationFunction
    If IsNumeric(value) Then
        XlConsolidationFunctionFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlUnknown": XlConsolidationFunctionFromString = xlUnknown
        Case "xlVarP": XlConsolidationFunctionFromString = xlVarP
        Case "xlVar": XlConsolidationFunctionFromString = xlVar
        Case "xlSum": XlConsolidationFunctionFromString = xlSum
        Case "xlStDevP": XlConsolidationFunctionFromString = xlStDevP
        Case "xlStDev": XlConsolidationFunctionFromString = xlStDev
        Case "xlProduct": XlConsolidationFunctionFromString = xlProduct
        Case "xlMin": XlConsolidationFunctionFromString = xlMin
        Case "xlMax": XlConsolidationFunctionFromString = xlMax
        Case "xlCountNums": XlConsolidationFunctionFromString = xlCountNums
        Case "xlCount": XlConsolidationFunctionFromString = xlCount
        Case "xlAverage": XlConsolidationFunctionFromString = xlAverage
    End Select
End Function

Function XlConsolidationFunctionToString(value As XlConsolidationFunction) As String
    Select Case value
        Case xlUnknown: XlConsolidationFunctionToString = "xlUnknown"
        Case xlVarP: XlConsolidationFunctionToString = "xlVarP"
        Case xlVar: XlConsolidationFunctionToString = "xlVar"
        Case xlSum: XlConsolidationFunctionToString = "xlSum"
        Case xlStDevP: XlConsolidationFunctionToString = "xlStDevP"
        Case xlStDev: XlConsolidationFunctionToString = "xlStDev"
        Case xlProduct: XlConsolidationFunctionToString = "xlProduct"
        Case xlMin: XlConsolidationFunctionToString = "xlMin"
        Case xlMax: XlConsolidationFunctionToString = "xlMax"
        Case xlCountNums: XlConsolidationFunctionToString = "xlCountNums"
        Case xlCount: XlConsolidationFunctionToString = "xlCount"
        Case xlAverage: XlConsolidationFunctionToString = "xlAverage"
    End Select
End Function
