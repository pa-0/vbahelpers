Attribute VB_Name = "wXlErrorBarType"
Function XlErrorBarTypeFromString(value As String) As XlErrorBarType
    If IsNumeric(value) Then
        XlErrorBarTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlErrorBarTypeFixedValue": XlErrorBarTypeFromString = xlErrorBarTypeFixedValue
        Case "xlErrorBarTypePercent": XlErrorBarTypeFromString = xlErrorBarTypePercent
        Case "xlErrorBarTypeStError": XlErrorBarTypeFromString = xlErrorBarTypeStError
        Case "xlErrorBarTypeStDev": XlErrorBarTypeFromString = xlErrorBarTypeStDev
        Case "xlErrorBarTypeCustom": XlErrorBarTypeFromString = xlErrorBarTypeCustom
    End Select
End Function

Function XlErrorBarTypeToString(value As XlErrorBarType) As String
    Select Case value
        Case xlErrorBarTypeFixedValue: XlErrorBarTypeToString = "xlErrorBarTypeFixedValue"
        Case xlErrorBarTypePercent: XlErrorBarTypeToString = "xlErrorBarTypePercent"
        Case xlErrorBarTypeStError: XlErrorBarTypeToString = "xlErrorBarTypeStError"
        Case xlErrorBarTypeStDev: XlErrorBarTypeToString = "xlErrorBarTypeStDev"
        Case xlErrorBarTypeCustom: XlErrorBarTypeToString = "xlErrorBarTypeCustom"
    End Select
End Function
