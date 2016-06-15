Attribute VB_Name = "wOlMeetingRecipientType"
Function OlMeetingRecipientTypeFromString(value As String) As OlMeetingRecipientType
    If IsNumeric(value) Then
        OlMeetingRecipientTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "olOrganizer": OlMeetingRecipientTypeFromString = olOrganizer
        Case "olRequired": OlMeetingRecipientTypeFromString = olRequired
        Case "olOptional": OlMeetingRecipientTypeFromString = olOptional
        Case "olResource": OlMeetingRecipientTypeFromString = olResource
    End Select
End Function

Function OlMeetingRecipientTypeToString(value As OlMeetingRecipientType) As String
    Select Case value
        Case olOrganizer: OlMeetingRecipientTypeToString = "olOrganizer"
        Case olRequired: OlMeetingRecipientTypeToString = "olRequired"
        Case olOptional: OlMeetingRecipientTypeToString = "olOptional"
        Case olResource: OlMeetingRecipientTypeToString = "olResource"
    End Select
End Function
