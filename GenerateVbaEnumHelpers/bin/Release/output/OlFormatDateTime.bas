Attribute VB_Name = "wOlFormatDateTime"
Function OlFormatDateTimeFromString(value As String) As OlFormatDateTime
    If IsNumeric(value) Then
        OlFormatDateTimeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "olFormatDateTimeLongDayDateTime": OlFormatDateTimeFromString = olFormatDateTimeLongDayDateTime
        Case "olFormatDateTimeShortDateTime": OlFormatDateTimeFromString = olFormatDateTimeShortDateTime
        Case "olFormatDateTimeShortDayDateTime": OlFormatDateTimeFromString = olFormatDateTimeShortDayDateTime
        Case "olFormatDateTimeShortDayMonthDateTime": OlFormatDateTimeFromString = olFormatDateTimeShortDayMonthDateTime
        Case "OlFormatDateTimeLongDayDate": OlFormatDateTimeFromString = OlFormatDateTimeLongDayDate
        Case "olFormatDateTimeLongDate": OlFormatDateTimeFromString = olFormatDateTimeLongDate
        Case "olFormatDateTimeLongDateReversed": OlFormatDateTimeFromString = olFormatDateTimeLongDateReversed
        Case "olFormatDateTimeShortDate": OlFormatDateTimeFromString = olFormatDateTimeShortDate
        Case "olFormatDateTimeShortDateNumOnly": OlFormatDateTimeFromString = olFormatDateTimeShortDateNumOnly
        Case "olFormatDateTimeShortDayMonth": OlFormatDateTimeFromString = olFormatDateTimeShortDayMonth
        Case "olFormatDateTimeShortMonthYear": OlFormatDateTimeFromString = olFormatDateTimeShortMonthYear
        Case "olFormatDateTimeShortMonthYearNumOnly": OlFormatDateTimeFromString = olFormatDateTimeShortMonthYearNumOnly
        Case "olFormatDateTimeShortDayDate": OlFormatDateTimeFromString = olFormatDateTimeShortDayDate
        Case "olFormatDateTimeLongTime": OlFormatDateTimeFromString = olFormatDateTimeLongTime
        Case "olFormatDateTimeShortTime": OlFormatDateTimeFromString = olFormatDateTimeShortTime
        Case "olFormatDateTimeBestFit": OlFormatDateTimeFromString = olFormatDateTimeBestFit
    End Select
End Function

Function OlFormatDateTimeToString(value As OlFormatDateTime) As String
    Select Case value
        Case olFormatDateTimeLongDayDateTime: OlFormatDateTimeToString = "olFormatDateTimeLongDayDateTime"
        Case olFormatDateTimeShortDateTime: OlFormatDateTimeToString = "olFormatDateTimeShortDateTime"
        Case olFormatDateTimeShortDayDateTime: OlFormatDateTimeToString = "olFormatDateTimeShortDayDateTime"
        Case olFormatDateTimeShortDayMonthDateTime: OlFormatDateTimeToString = "olFormatDateTimeShortDayMonthDateTime"
        Case OlFormatDateTimeLongDayDate: OlFormatDateTimeToString = "OlFormatDateTimeLongDayDate"
        Case olFormatDateTimeLongDate: OlFormatDateTimeToString = "olFormatDateTimeLongDate"
        Case olFormatDateTimeLongDateReversed: OlFormatDateTimeToString = "olFormatDateTimeLongDateReversed"
        Case olFormatDateTimeShortDate: OlFormatDateTimeToString = "olFormatDateTimeShortDate"
        Case olFormatDateTimeShortDateNumOnly: OlFormatDateTimeToString = "olFormatDateTimeShortDateNumOnly"
        Case olFormatDateTimeShortDayMonth: OlFormatDateTimeToString = "olFormatDateTimeShortDayMonth"
        Case olFormatDateTimeShortMonthYear: OlFormatDateTimeToString = "olFormatDateTimeShortMonthYear"
        Case olFormatDateTimeShortMonthYearNumOnly: OlFormatDateTimeToString = "olFormatDateTimeShortMonthYearNumOnly"
        Case olFormatDateTimeShortDayDate: OlFormatDateTimeToString = "olFormatDateTimeShortDayDate"
        Case olFormatDateTimeLongTime: OlFormatDateTimeToString = "olFormatDateTimeLongTime"
        Case olFormatDateTimeShortTime: OlFormatDateTimeToString = "olFormatDateTimeShortTime"
        Case olFormatDateTimeBestFit: OlFormatDateTimeToString = "olFormatDateTimeBestFit"
    End Select
End Function
