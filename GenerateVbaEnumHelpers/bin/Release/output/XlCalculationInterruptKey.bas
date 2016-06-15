Attribute VB_Name = "wXlCalculationInterruptKey"
Function XlCalculationInterruptKeyFromString(value As String) As XlCalculationInterruptKey
    If IsNumeric(value) Then
        XlCalculationInterruptKeyFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlNoKey": XlCalculationInterruptKeyFromString = xlNoKey
        Case "xlEscKey": XlCalculationInterruptKeyFromString = xlEscKey
        Case "xlAnyKey": XlCalculationInterruptKeyFromString = xlAnyKey
    End Select
End Function

Function XlCalculationInterruptKeyToString(value As XlCalculationInterruptKey) As String
    Select Case value
        Case xlNoKey: XlCalculationInterruptKeyToString = "xlNoKey"
        Case xlEscKey: XlCalculationInterruptKeyToString = "xlEscKey"
        Case xlAnyKey: XlCalculationInterruptKeyToString = "xlAnyKey"
    End Select
End Function
