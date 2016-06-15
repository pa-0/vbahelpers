Attribute VB_Name = "wOlNetMeetingType"
Function OlNetMeetingTypeFromString(value As String) As OlNetMeetingType
    If IsNumeric(value) Then
        OlNetMeetingTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "olNetMeeting": OlNetMeetingTypeFromString = olNetMeeting
        Case "olNetShow": OlNetMeetingTypeFromString = olNetShow
        Case "olExchangeConferencing": OlNetMeetingTypeFromString = olExchangeConferencing
    End Select
End Function

Function OlNetMeetingTypeToString(value As OlNetMeetingType) As String
    Select Case value
        Case olNetMeeting: OlNetMeetingTypeToString = "olNetMeeting"
        Case olNetShow: OlNetMeetingTypeToString = "olNetShow"
        Case olExchangeConferencing: OlNetMeetingTypeToString = "olExchangeConferencing"
    End Select
End Function
