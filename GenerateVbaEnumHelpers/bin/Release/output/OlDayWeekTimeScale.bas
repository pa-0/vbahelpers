Attribute VB_Name = "wOlDayWeekTimeScale"
Function OlDayWeekTimeScaleFromString(value As String) As OlDayWeekTimeScale
    If IsNumeric(value) Then
        OlDayWeekTimeScaleFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "olTimeScale5Minutes": OlDayWeekTimeScaleFromString = olTimeScale5Minutes
        Case "olTimeScale6Minutes": OlDayWeekTimeScaleFromString = olTimeScale6Minutes
        Case "olTimeScale10Minutes": OlDayWeekTimeScaleFromString = olTimeScale10Minutes
        Case "olTimeScale15Minutes": OlDayWeekTimeScaleFromString = olTimeScale15Minutes
        Case "olTimeScale30Minutes": OlDayWeekTimeScaleFromString = olTimeScale30Minutes
        Case "olTimeScale60Minutes": OlDayWeekTimeScaleFromString = olTimeScale60Minutes
    End Select
End Function

Function OlDayWeekTimeScaleToString(value As OlDayWeekTimeScale) As String
    Select Case value
        Case olTimeScale5Minutes: OlDayWeekTimeScaleToString = "olTimeScale5Minutes"
        Case olTimeScale6Minutes: OlDayWeekTimeScaleToString = "olTimeScale6Minutes"
        Case olTimeScale10Minutes: OlDayWeekTimeScaleToString = "olTimeScale10Minutes"
        Case olTimeScale15Minutes: OlDayWeekTimeScaleToString = "olTimeScale15Minutes"
        Case olTimeScale30Minutes: OlDayWeekTimeScaleToString = "olTimeScale30Minutes"
        Case olTimeScale60Minutes: OlDayWeekTimeScaleToString = "olTimeScale60Minutes"
    End Select
End Function
