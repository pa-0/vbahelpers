Attribute VB_Name = "wXlApplicationInternational"
Function XlApplicationInternationalFromString(value As String) As XlApplicationInternational
    If IsNumeric(value) Then
        XlApplicationInternationalFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlCountryCode": XlApplicationInternationalFromString = xlCountryCode
        Case "xlCountrySetting": XlApplicationInternationalFromString = xlCountrySetting
        Case "xlDecimalSeparator": XlApplicationInternationalFromString = xlDecimalSeparator
        Case "xlThousandsSeparator": XlApplicationInternationalFromString = xlThousandsSeparator
        Case "xlListSeparator": XlApplicationInternationalFromString = xlListSeparator
        Case "xlUpperCaseRowLetter": XlApplicationInternationalFromString = xlUpperCaseRowLetter
        Case "xlUpperCaseColumnLetter": XlApplicationInternationalFromString = xlUpperCaseColumnLetter
        Case "xlLowerCaseRowLetter": XlApplicationInternationalFromString = xlLowerCaseRowLetter
        Case "xlLowerCaseColumnLetter": XlApplicationInternationalFromString = xlLowerCaseColumnLetter
        Case "xlLeftBracket": XlApplicationInternationalFromString = xlLeftBracket
        Case "xlRightBracket": XlApplicationInternationalFromString = xlRightBracket
        Case "xlLeftBrace": XlApplicationInternationalFromString = xlLeftBrace
        Case "xlRightBrace": XlApplicationInternationalFromString = xlRightBrace
        Case "xlColumnSeparator": XlApplicationInternationalFromString = xlColumnSeparator
        Case "xlRowSeparator": XlApplicationInternationalFromString = xlRowSeparator
        Case "xlAlternateArraySeparator": XlApplicationInternationalFromString = xlAlternateArraySeparator
        Case "xlDateSeparator": XlApplicationInternationalFromString = xlDateSeparator
        Case "xlTimeSeparator": XlApplicationInternationalFromString = xlTimeSeparator
        Case "xlYearCode": XlApplicationInternationalFromString = xlYearCode
        Case "xlMonthCode": XlApplicationInternationalFromString = xlMonthCode
        Case "xlDayCode": XlApplicationInternationalFromString = xlDayCode
        Case "xlHourCode": XlApplicationInternationalFromString = xlHourCode
        Case "xlMinuteCode": XlApplicationInternationalFromString = xlMinuteCode
        Case "xlSecondCode": XlApplicationInternationalFromString = xlSecondCode
        Case "xlCurrencyCode": XlApplicationInternationalFromString = xlCurrencyCode
        Case "xlGeneralFormatName": XlApplicationInternationalFromString = xlGeneralFormatName
        Case "xlCurrencyDigits": XlApplicationInternationalFromString = xlCurrencyDigits
        Case "xlCurrencyNegative": XlApplicationInternationalFromString = xlCurrencyNegative
        Case "xlNoncurrencyDigits": XlApplicationInternationalFromString = xlNoncurrencyDigits
        Case "xlMonthNameChars": XlApplicationInternationalFromString = xlMonthNameChars
        Case "xlWeekdayNameChars": XlApplicationInternationalFromString = xlWeekdayNameChars
        Case "xlDateOrder": XlApplicationInternationalFromString = xlDateOrder
        Case "xl24HourClock": XlApplicationInternationalFromString = xl24HourClock
        Case "xlNonEnglishFunctions": XlApplicationInternationalFromString = xlNonEnglishFunctions
        Case "xlMetric": XlApplicationInternationalFromString = xlMetric
        Case "xlCurrencySpaceBefore": XlApplicationInternationalFromString = xlCurrencySpaceBefore
        Case "xlCurrencyBefore": XlApplicationInternationalFromString = xlCurrencyBefore
        Case "xlCurrencyMinusSign": XlApplicationInternationalFromString = xlCurrencyMinusSign
        Case "xlCurrencyTrailingZeros": XlApplicationInternationalFromString = xlCurrencyTrailingZeros
        Case "xlCurrencyLeadingZeros": XlApplicationInternationalFromString = xlCurrencyLeadingZeros
        Case "xlMonthLeadingZero": XlApplicationInternationalFromString = xlMonthLeadingZero
        Case "xlDayLeadingZero": XlApplicationInternationalFromString = xlDayLeadingZero
        Case "xl4DigitYears": XlApplicationInternationalFromString = xl4DigitYears
        Case "xlMDY": XlApplicationInternationalFromString = xlMDY
        Case "xlTimeLeadingZero": XlApplicationInternationalFromString = xlTimeLeadingZero
    End Select
End Function

Function XlApplicationInternationalToString(value As XlApplicationInternational) As String
    Select Case value
        Case xlCountryCode: XlApplicationInternationalToString = "xlCountryCode"
        Case xlCountrySetting: XlApplicationInternationalToString = "xlCountrySetting"
        Case xlDecimalSeparator: XlApplicationInternationalToString = "xlDecimalSeparator"
        Case xlThousandsSeparator: XlApplicationInternationalToString = "xlThousandsSeparator"
        Case xlListSeparator: XlApplicationInternationalToString = "xlListSeparator"
        Case xlUpperCaseRowLetter: XlApplicationInternationalToString = "xlUpperCaseRowLetter"
        Case xlUpperCaseColumnLetter: XlApplicationInternationalToString = "xlUpperCaseColumnLetter"
        Case xlLowerCaseRowLetter: XlApplicationInternationalToString = "xlLowerCaseRowLetter"
        Case xlLowerCaseColumnLetter: XlApplicationInternationalToString = "xlLowerCaseColumnLetter"
        Case xlLeftBracket: XlApplicationInternationalToString = "xlLeftBracket"
        Case xlRightBracket: XlApplicationInternationalToString = "xlRightBracket"
        Case xlLeftBrace: XlApplicationInternationalToString = "xlLeftBrace"
        Case xlRightBrace: XlApplicationInternationalToString = "xlRightBrace"
        Case xlColumnSeparator: XlApplicationInternationalToString = "xlColumnSeparator"
        Case xlRowSeparator: XlApplicationInternationalToString = "xlRowSeparator"
        Case xlAlternateArraySeparator: XlApplicationInternationalToString = "xlAlternateArraySeparator"
        Case xlDateSeparator: XlApplicationInternationalToString = "xlDateSeparator"
        Case xlTimeSeparator: XlApplicationInternationalToString = "xlTimeSeparator"
        Case xlYearCode: XlApplicationInternationalToString = "xlYearCode"
        Case xlMonthCode: XlApplicationInternationalToString = "xlMonthCode"
        Case xlDayCode: XlApplicationInternationalToString = "xlDayCode"
        Case xlHourCode: XlApplicationInternationalToString = "xlHourCode"
        Case xlMinuteCode: XlApplicationInternationalToString = "xlMinuteCode"
        Case xlSecondCode: XlApplicationInternationalToString = "xlSecondCode"
        Case xlCurrencyCode: XlApplicationInternationalToString = "xlCurrencyCode"
        Case xlGeneralFormatName: XlApplicationInternationalToString = "xlGeneralFormatName"
        Case xlCurrencyDigits: XlApplicationInternationalToString = "xlCurrencyDigits"
        Case xlCurrencyNegative: XlApplicationInternationalToString = "xlCurrencyNegative"
        Case xlNoncurrencyDigits: XlApplicationInternationalToString = "xlNoncurrencyDigits"
        Case xlMonthNameChars: XlApplicationInternationalToString = "xlMonthNameChars"
        Case xlWeekdayNameChars: XlApplicationInternationalToString = "xlWeekdayNameChars"
        Case xlDateOrder: XlApplicationInternationalToString = "xlDateOrder"
        Case xl24HourClock: XlApplicationInternationalToString = "xl24HourClock"
        Case xlNonEnglishFunctions: XlApplicationInternationalToString = "xlNonEnglishFunctions"
        Case xlMetric: XlApplicationInternationalToString = "xlMetric"
        Case xlCurrencySpaceBefore: XlApplicationInternationalToString = "xlCurrencySpaceBefore"
        Case xlCurrencyBefore: XlApplicationInternationalToString = "xlCurrencyBefore"
        Case xlCurrencyMinusSign: XlApplicationInternationalToString = "xlCurrencyMinusSign"
        Case xlCurrencyTrailingZeros: XlApplicationInternationalToString = "xlCurrencyTrailingZeros"
        Case xlCurrencyLeadingZeros: XlApplicationInternationalToString = "xlCurrencyLeadingZeros"
        Case xlMonthLeadingZero: XlApplicationInternationalToString = "xlMonthLeadingZero"
        Case xlDayLeadingZero: XlApplicationInternationalToString = "xlDayLeadingZero"
        Case xl4DigitYears: XlApplicationInternationalToString = "xl4DigitYears"
        Case xlMDY: XlApplicationInternationalToString = "xlMDY"
        Case xlTimeLeadingZero: XlApplicationInternationalToString = "xlTimeLeadingZero"
    End Select
End Function
