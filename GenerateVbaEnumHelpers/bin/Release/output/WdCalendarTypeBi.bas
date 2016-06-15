Attribute VB_Name = "wWdCalendarTypeBi"
Function WdCalendarTypeBiFromString(value As String) As WdCalendarTypeBi
    If IsNumeric(value) Then
        WdCalendarTypeBiFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdCalendarTypeBidi": WdCalendarTypeBiFromString = wdCalendarTypeBidi
        Case "wdCalendarTypeGregorian": WdCalendarTypeBiFromString = wdCalendarTypeGregorian
    End Select
End Function

Function WdCalendarTypeBiToString(value As WdCalendarTypeBi) As String
    Select Case value
        Case wdCalendarTypeBidi: WdCalendarTypeBiToString = "wdCalendarTypeBidi"
        Case wdCalendarTypeGregorian: WdCalendarTypeBiToString = "wdCalendarTypeGregorian"
    End Select
End Function
