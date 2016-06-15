Attribute VB_Name = "wOlFormatCurrency"
Function OlFormatCurrencyFromString(value As String) As OlFormatCurrency
    If IsNumeric(value) Then
        OlFormatCurrencyFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "olFormatCurrencyDecimal": OlFormatCurrencyFromString = olFormatCurrencyDecimal
        Case "olFormatCurrencyNonDecimal": OlFormatCurrencyFromString = olFormatCurrencyNonDecimal
    End Select
End Function

Function OlFormatCurrencyToString(value As OlFormatCurrency) As String
    Select Case value
        Case olFormatCurrencyDecimal: OlFormatCurrencyToString = "olFormatCurrencyDecimal"
        Case olFormatCurrencyNonDecimal: OlFormatCurrencyToString = "olFormatCurrencyNonDecimal"
    End Select
End Function
