Attribute VB_Name = "wOlFormatNumber"
Function OlFormatNumberFromString(value As String) As OlFormatNumber
    If IsNumeric(value) Then
        OlFormatNumberFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "olFormatNumberAllDigits": OlFormatNumberFromString = olFormatNumberAllDigits
        Case "olFormatNumberTruncated": OlFormatNumberFromString = olFormatNumberTruncated
        Case "olFormatNumber1Decimal": OlFormatNumberFromString = olFormatNumber1Decimal
        Case "olFormatNumber2Decimal": OlFormatNumberFromString = olFormatNumber2Decimal
        Case "olFormatNumberScientific": OlFormatNumberFromString = olFormatNumberScientific
        Case "olFormatNumberComputer1": OlFormatNumberFromString = olFormatNumberComputer1
        Case "olFormatNumberComputer2": OlFormatNumberFromString = olFormatNumberComputer2
        Case "olFormatNumberComputer3": OlFormatNumberFromString = olFormatNumberComputer3
        Case "olFormatNumberRaw": OlFormatNumberFromString = olFormatNumberRaw
    End Select
End Function

Function OlFormatNumberToString(value As OlFormatNumber) As String
    Select Case value
        Case olFormatNumberAllDigits: OlFormatNumberToString = "olFormatNumberAllDigits"
        Case olFormatNumberTruncated: OlFormatNumberToString = "olFormatNumberTruncated"
        Case olFormatNumber1Decimal: OlFormatNumberToString = "olFormatNumber1Decimal"
        Case olFormatNumber2Decimal: OlFormatNumberToString = "olFormatNumber2Decimal"
        Case olFormatNumberScientific: OlFormatNumberToString = "olFormatNumberScientific"
        Case olFormatNumberComputer1: OlFormatNumberToString = "olFormatNumberComputer1"
        Case olFormatNumberComputer2: OlFormatNumberToString = "olFormatNumberComputer2"
        Case olFormatNumberComputer3: OlFormatNumberToString = "olFormatNumberComputer3"
        Case olFormatNumberRaw: OlFormatNumberToString = "olFormatNumberRaw"
    End Select
End Function
