Attribute VB_Name = "wOlSelectionContents"
Function OlSelectionContentsFromString(value As String) As OlSelectionContents
    If IsNumeric(value) Then
        OlSelectionContentsFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "olConversationHeaders": OlSelectionContentsFromString = olConversationHeaders
    End Select
End Function

Function OlSelectionContentsToString(value As OlSelectionContents) As String
    Select Case value
        Case olConversationHeaders: OlSelectionContentsToString = "olConversationHeaders"
    End Select
End Function
