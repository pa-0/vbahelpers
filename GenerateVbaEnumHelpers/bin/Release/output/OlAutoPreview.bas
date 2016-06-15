Attribute VB_Name = "wOlAutoPreview"
Function OlAutoPreviewFromString(value As String) As OlAutoPreview
    If IsNumeric(value) Then
        OlAutoPreviewFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "olAutoPreviewAll": OlAutoPreviewFromString = olAutoPreviewAll
        Case "olAutoPreviewUnread": OlAutoPreviewFromString = olAutoPreviewUnread
        Case "olAutoPreviewNone": OlAutoPreviewFromString = olAutoPreviewNone
    End Select
End Function

Function OlAutoPreviewToString(value As OlAutoPreview) As String
    Select Case value
        Case olAutoPreviewAll: OlAutoPreviewToString = "olAutoPreviewAll"
        Case olAutoPreviewUnread: OlAutoPreviewToString = "olAutoPreviewUnread"
        Case olAutoPreviewNone: OlAutoPreviewToString = "olAutoPreviewNone"
    End Select
End Function
