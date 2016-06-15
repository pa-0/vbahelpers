Attribute VB_Name = "wOlMeetingResponse"
Function OlMeetingResponseFromString(value As String) As OlMeetingResponse
    If IsNumeric(value) Then
        OlMeetingResponseFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "olMeetingTentative": OlMeetingResponseFromString = olMeetingTentative
        Case "olMeetingAccepted": OlMeetingResponseFromString = olMeetingAccepted
        Case "olMeetingDeclined": OlMeetingResponseFromString = olMeetingDeclined
    End Select
End Function

Function OlMeetingResponseToString(value As OlMeetingResponse) As String
    Select Case value
        Case olMeetingTentative: OlMeetingResponseToString = "olMeetingTentative"
        Case olMeetingAccepted: OlMeetingResponseToString = "olMeetingAccepted"
        Case olMeetingDeclined: OlMeetingResponseToString = "olMeetingDeclined"
    End Select
End Function
