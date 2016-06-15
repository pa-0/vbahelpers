Attribute VB_Name = "wXlTimePeriods"
Function XlTimePeriodsFromString(value As String) As XlTimePeriods
    If IsNumeric(value) Then
        XlTimePeriodsFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlToday": XlTimePeriodsFromString = xlToday
        Case "xlYesterday": XlTimePeriodsFromString = xlYesterday
        Case "xlLast7Days": XlTimePeriodsFromString = xlLast7Days
        Case "xlThisWeek": XlTimePeriodsFromString = xlThisWeek
        Case "xlLastWeek": XlTimePeriodsFromString = xlLastWeek
        Case "xlLastMonth": XlTimePeriodsFromString = xlLastMonth
        Case "xlTomorrow": XlTimePeriodsFromString = xlTomorrow
        Case "xlNextWeek": XlTimePeriodsFromString = xlNextWeek
        Case "xlNextMonth": XlTimePeriodsFromString = xlNextMonth
        Case "xlThisMonth": XlTimePeriodsFromString = xlThisMonth
    End Select
End Function

Function XlTimePeriodsToString(value As XlTimePeriods) As String
    Select Case value
        Case xlToday: XlTimePeriodsToString = "xlToday"
        Case xlYesterday: XlTimePeriodsToString = "xlYesterday"
        Case xlLast7Days: XlTimePeriodsToString = "xlLast7Days"
        Case xlThisWeek: XlTimePeriodsToString = "xlThisWeek"
        Case xlLastWeek: XlTimePeriodsToString = "xlLastWeek"
        Case xlLastMonth: XlTimePeriodsToString = "xlLastMonth"
        Case xlTomorrow: XlTimePeriodsToString = "xlTomorrow"
        Case xlNextWeek: XlTimePeriodsToString = "xlNextWeek"
        Case xlNextMonth: XlTimePeriodsToString = "xlNextMonth"
        Case xlThisMonth: XlTimePeriodsToString = "xlThisMonth"
    End Select
End Function
