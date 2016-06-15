Attribute VB_Name = "wOlCalendarViewMode"
Function OlCalendarViewModeFromString(value As String) As OlCalendarViewMode
    If IsNumeric(value) Then
        OlCalendarViewModeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "olCalendarViewDay": OlCalendarViewModeFromString = olCalendarViewDay
        Case "olCalendarViewWeek": OlCalendarViewModeFromString = olCalendarViewWeek
        Case "olCalendarViewMonth": OlCalendarViewModeFromString = olCalendarViewMonth
        Case "olCalendarViewMultiDay": OlCalendarViewModeFromString = olCalendarViewMultiDay
        Case "olCalendarView5DayWeek": OlCalendarViewModeFromString = olCalendarView5DayWeek
    End Select
End Function

Function OlCalendarViewModeToString(value As OlCalendarViewMode) As String
    Select Case value
        Case olCalendarViewDay: OlCalendarViewModeToString = "olCalendarViewDay"
        Case olCalendarViewWeek: OlCalendarViewModeToString = "olCalendarViewWeek"
        Case olCalendarViewMonth: OlCalendarViewModeToString = "olCalendarViewMonth"
        Case olCalendarViewMultiDay: OlCalendarViewModeToString = "olCalendarViewMultiDay"
        Case olCalendarView5DayWeek: OlCalendarViewModeToString = "olCalendarView5DayWeek"
    End Select
End Function
