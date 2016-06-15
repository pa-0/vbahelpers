Attribute VB_Name = "wOlMarkInterval"
Function OlMarkIntervalFromString(value As String) As OlMarkInterval
    If IsNumeric(value) Then
        OlMarkIntervalFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "olMarkToday": OlMarkIntervalFromString = olMarkToday
        Case "olMarkTomorrow": OlMarkIntervalFromString = olMarkTomorrow
        Case "olMarkThisWeek": OlMarkIntervalFromString = olMarkThisWeek
        Case "olMarkNextWeek": OlMarkIntervalFromString = olMarkNextWeek
        Case "olMarkNoDate": OlMarkIntervalFromString = olMarkNoDate
        Case "olMarkComplete": OlMarkIntervalFromString = olMarkComplete
    End Select
End Function

Function OlMarkIntervalToString(value As OlMarkInterval) As String
    Select Case value
        Case olMarkToday: OlMarkIntervalToString = "olMarkToday"
        Case olMarkTomorrow: OlMarkIntervalToString = "olMarkTomorrow"
        Case olMarkThisWeek: OlMarkIntervalToString = "olMarkThisWeek"
        Case olMarkNextWeek: OlMarkIntervalToString = "olMarkNextWeek"
        Case olMarkNoDate: OlMarkIntervalToString = "olMarkNoDate"
        Case olMarkComplete: OlMarkIntervalToString = "olMarkComplete"
    End Select
End Function
