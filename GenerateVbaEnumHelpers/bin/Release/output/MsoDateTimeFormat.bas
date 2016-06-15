Attribute VB_Name = "wMsoDateTimeFormat"
Function MsoDateTimeFormatFromString(value As String) As MsoDateTimeFormat
    If IsNumeric(value) Then
        MsoDateTimeFormatFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoDateTimeMdyy": MsoDateTimeFormatFromString = msoDateTimeMdyy
        Case "msoDateTimeddddMMMMddyyyy": MsoDateTimeFormatFromString = msoDateTimeddddMMMMddyyyy
        Case "msoDateTimedMMMMyyyy": MsoDateTimeFormatFromString = msoDateTimedMMMMyyyy
        Case "msoDateTimeMMMMdyyyy": MsoDateTimeFormatFromString = msoDateTimeMMMMdyyyy
        Case "msoDateTimedMMMyy": MsoDateTimeFormatFromString = msoDateTimedMMMyy
        Case "msoDateTimeMMMMyy": MsoDateTimeFormatFromString = msoDateTimeMMMMyy
        Case "msoDateTimeMMyy": MsoDateTimeFormatFromString = msoDateTimeMMyy
        Case "msoDateTimeMMddyyHmm": MsoDateTimeFormatFromString = msoDateTimeMMddyyHmm
        Case "msoDateTimeMMddyyhmmAMPM": MsoDateTimeFormatFromString = msoDateTimeMMddyyhmmAMPM
        Case "msoDateTimeHmm": MsoDateTimeFormatFromString = msoDateTimeHmm
        Case "msoDateTimeHmmss": MsoDateTimeFormatFromString = msoDateTimeHmmss
        Case "msoDateTimehmmAMPM": MsoDateTimeFormatFromString = msoDateTimehmmAMPM
        Case "msoDateTimehmmssAMPM": MsoDateTimeFormatFromString = msoDateTimehmmssAMPM
        Case "msoDateTimeFigureOut": MsoDateTimeFormatFromString = msoDateTimeFigureOut
        Case "msoDateTimeFormatMixed": MsoDateTimeFormatFromString = msoDateTimeFormatMixed
    End Select
End Function

Function MsoDateTimeFormatToString(value As MsoDateTimeFormat) As String
    Select Case value
        Case msoDateTimeMdyy: MsoDateTimeFormatToString = "msoDateTimeMdyy"
        Case msoDateTimeddddMMMMddyyyy: MsoDateTimeFormatToString = "msoDateTimeddddMMMMddyyyy"
        Case msoDateTimedMMMMyyyy: MsoDateTimeFormatToString = "msoDateTimedMMMMyyyy"
        Case msoDateTimeMMMMdyyyy: MsoDateTimeFormatToString = "msoDateTimeMMMMdyyyy"
        Case msoDateTimedMMMyy: MsoDateTimeFormatToString = "msoDateTimedMMMyy"
        Case msoDateTimeMMMMyy: MsoDateTimeFormatToString = "msoDateTimeMMMMyy"
        Case msoDateTimeMMyy: MsoDateTimeFormatToString = "msoDateTimeMMyy"
        Case msoDateTimeMMddyyHmm: MsoDateTimeFormatToString = "msoDateTimeMMddyyHmm"
        Case msoDateTimeMMddyyhmmAMPM: MsoDateTimeFormatToString = "msoDateTimeMMddyyhmmAMPM"
        Case msoDateTimeHmm: MsoDateTimeFormatToString = "msoDateTimeHmm"
        Case msoDateTimeHmmss: MsoDateTimeFormatToString = "msoDateTimeHmmss"
        Case msoDateTimehmmAMPM: MsoDateTimeFormatToString = "msoDateTimehmmAMPM"
        Case msoDateTimehmmssAMPM: MsoDateTimeFormatToString = "msoDateTimehmmssAMPM"
        Case msoDateTimeFigureOut: MsoDateTimeFormatToString = "msoDateTimeFigureOut"
        Case msoDateTimeFormatMixed: MsoDateTimeFormatToString = "msoDateTimeFormatMixed"
    End Select
End Function
