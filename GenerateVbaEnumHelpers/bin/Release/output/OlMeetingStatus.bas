Attribute VB_Name = "wOlMeetingStatus"
Function OlMeetingStatusFromString(value As String) As OlMeetingStatus
    If IsNumeric(value) Then
        OlMeetingStatusFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "olNonMeeting": OlMeetingStatusFromString = olNonMeeting
        Case "olMeeting": OlMeetingStatusFromString = olMeeting
        Case "olMeetingReceived": OlMeetingStatusFromString = olMeetingReceived
        Case "olMeetingCanceled": OlMeetingStatusFromString = olMeetingCanceled
        Case "olMeetingReceivedAndCanceled": OlMeetingStatusFromString = olMeetingReceivedAndCanceled
    End Select
End Function

Function OlMeetingStatusToString(value As OlMeetingStatus) As String
    Select Case value
        Case olNonMeeting: OlMeetingStatusToString = "olNonMeeting"
        Case olMeeting: OlMeetingStatusToString = "olMeeting"
        Case olMeetingReceived: OlMeetingStatusToString = "olMeetingReceived"
        Case olMeetingCanceled: OlMeetingStatusToString = "olMeetingCanceled"
        Case olMeetingReceivedAndCanceled: OlMeetingStatusToString = "olMeetingReceivedAndCanceled"
    End Select
End Function
