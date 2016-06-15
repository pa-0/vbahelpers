Attribute VB_Name = "wOlDefaultSelectNamesDisplayMode"
Function OlDefaultSelectNamesDisplayModeFromString(value As String) As OlDefaultSelectNamesDisplayMode
    If IsNumeric(value) Then
        OlDefaultSelectNamesDisplayModeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "olDefaultMail": OlDefaultSelectNamesDisplayModeFromString = olDefaultMail
        Case "olDefaultMeeting": OlDefaultSelectNamesDisplayModeFromString = olDefaultMeeting
        Case "olDefaultTask": OlDefaultSelectNamesDisplayModeFromString = olDefaultTask
        Case "olDefaultSharingRequest": OlDefaultSelectNamesDisplayModeFromString = olDefaultSharingRequest
        Case "olDefaultMembers": OlDefaultSelectNamesDisplayModeFromString = olDefaultMembers
        Case "olDefaultDelegates": OlDefaultSelectNamesDisplayModeFromString = olDefaultDelegates
        Case "olDefaultSingleName": OlDefaultSelectNamesDisplayModeFromString = olDefaultSingleName
        Case "olDefaultPickRooms": OlDefaultSelectNamesDisplayModeFromString = olDefaultPickRooms
    End Select
End Function

Function OlDefaultSelectNamesDisplayModeToString(value As OlDefaultSelectNamesDisplayMode) As String
    Select Case value
        Case olDefaultMail: OlDefaultSelectNamesDisplayModeToString = "olDefaultMail"
        Case olDefaultMeeting: OlDefaultSelectNamesDisplayModeToString = "olDefaultMeeting"
        Case olDefaultTask: OlDefaultSelectNamesDisplayModeToString = "olDefaultTask"
        Case olDefaultSharingRequest: OlDefaultSelectNamesDisplayModeToString = "olDefaultSharingRequest"
        Case olDefaultMembers: OlDefaultSelectNamesDisplayModeToString = "olDefaultMembers"
        Case olDefaultDelegates: OlDefaultSelectNamesDisplayModeToString = "olDefaultDelegates"
        Case olDefaultSingleName: OlDefaultSelectNamesDisplayModeToString = "olDefaultSingleName"
        Case olDefaultPickRooms: OlDefaultSelectNamesDisplayModeToString = "olDefaultPickRooms"
    End Select
End Function
