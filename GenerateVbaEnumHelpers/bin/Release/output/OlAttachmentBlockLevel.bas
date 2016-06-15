Attribute VB_Name = "wOlAttachmentBlockLevel"
Function OlAttachmentBlockLevelFromString(value As String) As OlAttachmentBlockLevel
    If IsNumeric(value) Then
        OlAttachmentBlockLevelFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "olAttachmentBlockLevelNone": OlAttachmentBlockLevelFromString = olAttachmentBlockLevelNone
        Case "olAttachmentBlockLevelOpen": OlAttachmentBlockLevelFromString = olAttachmentBlockLevelOpen
    End Select
End Function

Function OlAttachmentBlockLevelToString(value As OlAttachmentBlockLevel) As String
    Select Case value
        Case olAttachmentBlockLevelNone: OlAttachmentBlockLevelToString = "olAttachmentBlockLevelNone"
        Case olAttachmentBlockLevelOpen: OlAttachmentBlockLevelToString = "olAttachmentBlockLevelOpen"
    End Select
End Function
