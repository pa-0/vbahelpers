Attribute VB_Name = "wOlActionReplyStyle"
Function OlActionReplyStyleFromString(value As String) As OlActionReplyStyle
    If IsNumeric(value) Then
        OlActionReplyStyleFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "olOmitOriginalText": OlActionReplyStyleFromString = olOmitOriginalText
        Case "olEmbedOriginalItem": OlActionReplyStyleFromString = olEmbedOriginalItem
        Case "olIncludeOriginalText": OlActionReplyStyleFromString = olIncludeOriginalText
        Case "olIndentOriginalText": OlActionReplyStyleFromString = olIndentOriginalText
        Case "olLinkOriginalItem": OlActionReplyStyleFromString = olLinkOriginalItem
        Case "olUserPreference": OlActionReplyStyleFromString = olUserPreference
        Case "olReplyTickOriginalText": OlActionReplyStyleFromString = olReplyTickOriginalText
    End Select
End Function

Function OlActionReplyStyleToString(value As OlActionReplyStyle) As String
    Select Case value
        Case olOmitOriginalText: OlActionReplyStyleToString = "olOmitOriginalText"
        Case olEmbedOriginalItem: OlActionReplyStyleToString = "olEmbedOriginalItem"
        Case olIncludeOriginalText: OlActionReplyStyleToString = "olIncludeOriginalText"
        Case olIndentOriginalText: OlActionReplyStyleToString = "olIndentOriginalText"
        Case olLinkOriginalItem: OlActionReplyStyleToString = "olLinkOriginalItem"
        Case olUserPreference: OlActionReplyStyleToString = "olUserPreference"
        Case olReplyTickOriginalText: OlActionReplyStyleToString = "olReplyTickOriginalText"
    End Select
End Function
