Attribute VB_Name = "wOlAlwaysDeleteConversation"
Function OlAlwaysDeleteConversationFromString(value As String) As OlAlwaysDeleteConversation
    If IsNumeric(value) Then
        OlAlwaysDeleteConversationFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "olDoNotDelete": OlAlwaysDeleteConversationFromString = olDoNotDelete
        Case "olAlwaysDelete": OlAlwaysDeleteConversationFromString = olAlwaysDelete
        Case "olAlwaysDeleteUnsupported": OlAlwaysDeleteConversationFromString = olAlwaysDeleteUnsupported
    End Select
End Function

Function OlAlwaysDeleteConversationToString(value As OlAlwaysDeleteConversation) As String
    Select Case value
        Case olDoNotDelete: OlAlwaysDeleteConversationToString = "olDoNotDelete"
        Case olAlwaysDelete: OlAlwaysDeleteConversationToString = "olAlwaysDelete"
        Case olAlwaysDeleteUnsupported: OlAlwaysDeleteConversationToString = "olAlwaysDeleteUnsupported"
    End Select
End Function
