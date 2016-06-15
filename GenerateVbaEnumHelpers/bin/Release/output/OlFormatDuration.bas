Attribute VB_Name = "wOlFormatDuration"
Function OlFormatDurationFromString(value As String) As OlFormatDuration
    If IsNumeric(value) Then
        OlFormatDurationFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "olFormatDurationShort": OlFormatDurationFromString = olFormatDurationShort
        Case "olFormatDurationLong": OlFormatDurationFromString = olFormatDurationLong
        Case "olFormatDurationShortBusiness": OlFormatDurationFromString = olFormatDurationShortBusiness
        Case "olFormatDurationLongBusiness": OlFormatDurationFromString = olFormatDurationLongBusiness
    End Select
End Function

Function OlFormatDurationToString(value As OlFormatDuration) As String
    Select Case value
        Case olFormatDurationShort: OlFormatDurationToString = "olFormatDurationShort"
        Case olFormatDurationLong: OlFormatDurationToString = "olFormatDurationLong"
        Case olFormatDurationShortBusiness: OlFormatDurationToString = "olFormatDurationShortBusiness"
        Case olFormatDurationLongBusiness: OlFormatDurationToString = "olFormatDurationLongBusiness"
    End Select
End Function
