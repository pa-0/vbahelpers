Attribute VB_Name = "wXlCalculation"
Function XlCalculationFromString(value As String) As XlCalculation
    If IsNumeric(value) Then
        XlCalculationFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlCalculationSemiautomatic": XlCalculationFromString = xlCalculationSemiautomatic
        Case "xlCalculationManual": XlCalculationFromString = xlCalculationManual
        Case "xlCalculationAutomatic": XlCalculationFromString = xlCalculationAutomatic
    End Select
End Function

Function XlCalculationToString(value As XlCalculation) As String
    Select Case value
        Case xlCalculationSemiautomatic: XlCalculationToString = "xlCalculationSemiautomatic"
        Case xlCalculationManual: XlCalculationToString = "xlCalculationManual"
        Case xlCalculationAutomatic: XlCalculationToString = "xlCalculationAutomatic"
    End Select
End Function
