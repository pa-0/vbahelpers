Attribute VB_Name = "wXlFilterAllDatesInPeriod"
Function XlFilterAllDatesInPeriodFromString(value As String) As XlFilterAllDatesInPeriod
    If IsNumeric(value) Then
        XlFilterAllDatesInPeriodFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlFilterAllDatesInPeriodYear": XlFilterAllDatesInPeriodFromString = xlFilterAllDatesInPeriodYear
        Case "xlFilterAllDatesInPeriodMonth": XlFilterAllDatesInPeriodFromString = xlFilterAllDatesInPeriodMonth
        Case "xlFilterAllDatesInPeriodDay": XlFilterAllDatesInPeriodFromString = xlFilterAllDatesInPeriodDay
        Case "xlFilterAllDatesInPeriodHour": XlFilterAllDatesInPeriodFromString = xlFilterAllDatesInPeriodHour
        Case "xlFilterAllDatesInPeriodMinute": XlFilterAllDatesInPeriodFromString = xlFilterAllDatesInPeriodMinute
        Case "xlFilterAllDatesInPeriodSecond": XlFilterAllDatesInPeriodFromString = xlFilterAllDatesInPeriodSecond
    End Select
End Function

Function XlFilterAllDatesInPeriodToString(value As XlFilterAllDatesInPeriod) As String
    Select Case value
        Case xlFilterAllDatesInPeriodYear: XlFilterAllDatesInPeriodToString = "xlFilterAllDatesInPeriodYear"
        Case xlFilterAllDatesInPeriodMonth: XlFilterAllDatesInPeriodToString = "xlFilterAllDatesInPeriodMonth"
        Case xlFilterAllDatesInPeriodDay: XlFilterAllDatesInPeriodToString = "xlFilterAllDatesInPeriodDay"
        Case xlFilterAllDatesInPeriodHour: XlFilterAllDatesInPeriodToString = "xlFilterAllDatesInPeriodHour"
        Case xlFilterAllDatesInPeriodMinute: XlFilterAllDatesInPeriodToString = "xlFilterAllDatesInPeriodMinute"
        Case xlFilterAllDatesInPeriodSecond: XlFilterAllDatesInPeriodToString = "xlFilterAllDatesInPeriodSecond"
    End Select
End Function
