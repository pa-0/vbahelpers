Attribute VB_Name = "wOlCalendarDetail"
Function OlCalendarDetailFromString(value As String) As OlCalendarDetail
    If IsNumeric(value) Then
        OlCalendarDetailFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "olFreeBusyOnly": OlCalendarDetailFromString = olFreeBusyOnly
        Case "olFreeBusyAndSubject": OlCalendarDetailFromString = olFreeBusyAndSubject
        Case "olFullDetails": OlCalendarDetailFromString = olFullDetails
    End Select
End Function

Function OlCalendarDetailToString(value As OlCalendarDetail) As String
    Select Case value
        Case olFreeBusyOnly: OlCalendarDetailToString = "olFreeBusyOnly"
        Case olFreeBusyAndSubject: OlCalendarDetailToString = "olFreeBusyAndSubject"
        Case olFullDetails: OlCalendarDetailToString = "olFullDetails"
    End Select
End Function
