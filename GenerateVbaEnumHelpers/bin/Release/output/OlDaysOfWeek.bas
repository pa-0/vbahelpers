Attribute VB_Name = "wOlDaysOfWeek"
Function OlDaysOfWeekFromString(value As String) As OlDaysOfWeek
    If IsNumeric(value) Then
        OlDaysOfWeekFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "olSunday": OlDaysOfWeekFromString = olSunday
        Case "olMonday": OlDaysOfWeekFromString = olMonday
        Case "olTuesday": OlDaysOfWeekFromString = olTuesday
        Case "olWednesday": OlDaysOfWeekFromString = olWednesday
        Case "olThursday": OlDaysOfWeekFromString = olThursday
        Case "olFriday": OlDaysOfWeekFromString = olFriday
        Case "olSaturday": OlDaysOfWeekFromString = olSaturday
    End Select
End Function

Function OlDaysOfWeekToString(value As OlDaysOfWeek) As String
    Select Case value
        Case olSunday: OlDaysOfWeekToString = "olSunday"
        Case olMonday: OlDaysOfWeekToString = "olMonday"
        Case olTuesday: OlDaysOfWeekToString = "olTuesday"
        Case olWednesday: OlDaysOfWeekToString = "olWednesday"
        Case olThursday: OlDaysOfWeekToString = "olThursday"
        Case olFriday: OlDaysOfWeekToString = "olFriday"
        Case olSaturday: OlDaysOfWeekToString = "olSaturday"
    End Select
End Function
