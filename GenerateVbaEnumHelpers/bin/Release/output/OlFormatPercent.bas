Attribute VB_Name = "wOlFormatPercent"
Function OlFormatPercentFromString(value As String) As OlFormatPercent
    If IsNumeric(value) Then
        OlFormatPercentFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "olFormatPercentRounded": OlFormatPercentFromString = olFormatPercentRounded
        Case "olFormatPercent1Decimal": OlFormatPercentFromString = olFormatPercent1Decimal
        Case "olFormatPercent2Decimal": OlFormatPercentFromString = olFormatPercent2Decimal
        Case "olFormatPercentAllDigits": OlFormatPercentFromString = olFormatPercentAllDigits
    End Select
End Function

Function OlFormatPercentToString(value As OlFormatPercent) As String
    Select Case value
        Case olFormatPercentRounded: OlFormatPercentToString = "olFormatPercentRounded"
        Case olFormatPercent1Decimal: OlFormatPercentToString = "olFormatPercent1Decimal"
        Case olFormatPercent2Decimal: OlFormatPercentToString = "olFormatPercent2Decimal"
        Case olFormatPercentAllDigits: OlFormatPercentToString = "olFormatPercentAllDigits"
    End Select
End Function
