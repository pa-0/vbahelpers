Attribute VB_Name = "wOlCalendarMailFormat"
Function OlCalendarMailFormatFromString(value As String) As OlCalendarMailFormat
    If IsNumeric(value) Then
        OlCalendarMailFormatFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "olCalendarMailFormatDailySchedule": OlCalendarMailFormatFromString = olCalendarMailFormatDailySchedule
        Case "olCalendarMailFormatEventList": OlCalendarMailFormatFromString = olCalendarMailFormatEventList
    End Select
End Function

Function OlCalendarMailFormatToString(value As OlCalendarMailFormat) As String
    Select Case value
        Case olCalendarMailFormatDailySchedule: OlCalendarMailFormatToString = "olCalendarMailFormatDailySchedule"
        Case olCalendarMailFormatEventList: OlCalendarMailFormatToString = "olCalendarMailFormatEventList"
    End Select
End Function
