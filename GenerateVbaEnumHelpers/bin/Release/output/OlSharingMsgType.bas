Attribute VB_Name = "wOlSharingMsgType"
Function OlSharingMsgTypeFromString(value As String) As OlSharingMsgType
    If IsNumeric(value) Then
        OlSharingMsgTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "olSharingMsgTypeUnknown": OlSharingMsgTypeFromString = olSharingMsgTypeUnknown
        Case "olSharingMsgTypeRequest": OlSharingMsgTypeFromString = olSharingMsgTypeRequest
        Case "olSharingMsgTypeInvite": OlSharingMsgTypeFromString = olSharingMsgTypeInvite
        Case "olSharingMsgTypeInviteAndRequest": OlSharingMsgTypeFromString = olSharingMsgTypeInviteAndRequest
        Case "olSharingMsgTypeResponseAllow": OlSharingMsgTypeFromString = olSharingMsgTypeResponseAllow
        Case "olSharingMsgTypeResponseDeny": OlSharingMsgTypeFromString = olSharingMsgTypeResponseDeny
    End Select
End Function

Function OlSharingMsgTypeToString(value As OlSharingMsgType) As String
    Select Case value
        Case olSharingMsgTypeUnknown: OlSharingMsgTypeToString = "olSharingMsgTypeUnknown"
        Case olSharingMsgTypeRequest: OlSharingMsgTypeToString = "olSharingMsgTypeRequest"
        Case olSharingMsgTypeInvite: OlSharingMsgTypeToString = "olSharingMsgTypeInvite"
        Case olSharingMsgTypeInviteAndRequest: OlSharingMsgTypeToString = "olSharingMsgTypeInviteAndRequest"
        Case olSharingMsgTypeResponseAllow: OlSharingMsgTypeToString = "olSharingMsgTypeResponseAllow"
        Case olSharingMsgTypeResponseDeny: OlSharingMsgTypeToString = "olSharingMsgTypeResponseDeny"
    End Select
End Function
