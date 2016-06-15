Attribute VB_Name = "wOlActionCopyLike"
Function OlActionCopyLikeFromString(value As String) As OlActionCopyLike
    If IsNumeric(value) Then
        OlActionCopyLikeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "olReply": OlActionCopyLikeFromString = olReply
        Case "olReplyAll": OlActionCopyLikeFromString = olReplyAll
        Case "olForward": OlActionCopyLikeFromString = olForward
        Case "olReplyFolder": OlActionCopyLikeFromString = olReplyFolder
        Case "olRespond": OlActionCopyLikeFromString = olRespond
    End Select
End Function

Function OlActionCopyLikeToString(value As OlActionCopyLike) As String
    Select Case value
        Case olReply: OlActionCopyLikeToString = "olReply"
        Case olReplyAll: OlActionCopyLikeToString = "olReplyAll"
        Case olForward: OlActionCopyLikeToString = "olForward"
        Case olReplyFolder: OlActionCopyLikeToString = "olReplyFolder"
        Case olRespond: OlActionCopyLikeToString = "olRespond"
    End Select
End Function
