Attribute VB_Name = "wPpDateTimeFormat"
Function PpDateTimeFormatFromString(value As String) As PpDateTimeFormat
    If IsNumeric(value) Then
        PpDateTimeFormatFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "ppDateTimeMdyy": PpDateTimeFormatFromString = ppDateTimeMdyy
        Case "ppDateTimeddddMMMMddyyyy": PpDateTimeFormatFromString = ppDateTimeddddMMMMddyyyy
        Case "ppDateTimedMMMMyyyy": PpDateTimeFormatFromString = ppDateTimedMMMMyyyy
        Case "ppDateTimeMMMMdyyyy": PpDateTimeFormatFromString = ppDateTimeMMMMdyyyy
        Case "ppDateTimedMMMyy": PpDateTimeFormatFromString = ppDateTimedMMMyy
        Case "ppDateTimeMMMMyy": PpDateTimeFormatFromString = ppDateTimeMMMMyy
        Case "ppDateTimeMMyy": PpDateTimeFormatFromString = ppDateTimeMMyy
        Case "ppDateTimeMMddyyHmm": PpDateTimeFormatFromString = ppDateTimeMMddyyHmm
        Case "ppDateTimeMMddyyhmmAMPM": PpDateTimeFormatFromString = ppDateTimeMMddyyhmmAMPM
        Case "ppDateTimeHmm": PpDateTimeFormatFromString = ppDateTimeHmm
        Case "ppDateTimeHmmss": PpDateTimeFormatFromString = ppDateTimeHmmss
        Case "ppDateTimehmmAMPM": PpDateTimeFormatFromString = ppDateTimehmmAMPM
        Case "ppDateTimehmmssAMPM": PpDateTimeFormatFromString = ppDateTimehmmssAMPM
        Case "ppDateTimeFigureOut": PpDateTimeFormatFromString = ppDateTimeFigureOut
        Case "ppDateTimeFormatMixed": PpDateTimeFormatFromString = ppDateTimeFormatMixed
    End Select
End Function

Function PpDateTimeFormatToString(value As PpDateTimeFormat) As String
    Select Case value
        Case ppDateTimeMdyy: PpDateTimeFormatToString = "ppDateTimeMdyy"
        Case ppDateTimeddddMMMMddyyyy: PpDateTimeFormatToString = "ppDateTimeddddMMMMddyyyy"
        Case ppDateTimedMMMMyyyy: PpDateTimeFormatToString = "ppDateTimedMMMMyyyy"
        Case ppDateTimeMMMMdyyyy: PpDateTimeFormatToString = "ppDateTimeMMMMdyyyy"
        Case ppDateTimedMMMyy: PpDateTimeFormatToString = "ppDateTimedMMMyy"
        Case ppDateTimeMMMMyy: PpDateTimeFormatToString = "ppDateTimeMMMMyy"
        Case ppDateTimeMMyy: PpDateTimeFormatToString = "ppDateTimeMMyy"
        Case ppDateTimeMMddyyHmm: PpDateTimeFormatToString = "ppDateTimeMMddyyHmm"
        Case ppDateTimeMMddyyhmmAMPM: PpDateTimeFormatToString = "ppDateTimeMMddyyhmmAMPM"
        Case ppDateTimeHmm: PpDateTimeFormatToString = "ppDateTimeHmm"
        Case ppDateTimeHmmss: PpDateTimeFormatToString = "ppDateTimeHmmss"
        Case ppDateTimehmmAMPM: PpDateTimeFormatToString = "ppDateTimehmmAMPM"
        Case ppDateTimehmmssAMPM: PpDateTimeFormatToString = "ppDateTimehmmssAMPM"
        Case ppDateTimeFigureOut: PpDateTimeFormatToString = "ppDateTimeFigureOut"
        Case ppDateTimeFormatMixed: PpDateTimeFormatToString = "ppDateTimeFormatMixed"
    End Select
End Function
