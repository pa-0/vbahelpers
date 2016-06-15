Attribute VB_Name = "wOlFormRegionIcon"
Function OlFormRegionIconFromString(value As String) As OlFormRegionIcon
    If IsNumeric(value) Then
        OlFormRegionIconFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "olFormRegionIconDefault": OlFormRegionIconFromString = olFormRegionIconDefault
        Case "olFormRegionIconUnread": OlFormRegionIconFromString = olFormRegionIconUnread
        Case "olFormRegionIconRead": OlFormRegionIconFromString = olFormRegionIconRead
        Case "olFormRegionIconReplied": OlFormRegionIconFromString = olFormRegionIconReplied
        Case "olFormRegionIconForwarded": OlFormRegionIconFromString = olFormRegionIconForwarded
        Case "olFormRegionIconUnsent": OlFormRegionIconFromString = olFormRegionIconUnsent
        Case "olFormRegionIconSubmitted": OlFormRegionIconFromString = olFormRegionIconSubmitted
        Case "olFormRegionIconSigned": OlFormRegionIconFromString = olFormRegionIconSigned
        Case "olFormRegionIconEncrypted": OlFormRegionIconFromString = olFormRegionIconEncrypted
        Case "olFormRegionIconWindow": OlFormRegionIconFromString = olFormRegionIconWindow
        Case "olFormRegionIconPage": OlFormRegionIconFromString = olFormRegionIconPage
        Case "olFormRegionIconRecurring": OlFormRegionIconFromString = olFormRegionIconRecurring
    End Select
End Function

Function OlFormRegionIconToString(value As OlFormRegionIcon) As String
    Select Case value
        Case olFormRegionIconDefault: OlFormRegionIconToString = "olFormRegionIconDefault"
        Case olFormRegionIconUnread: OlFormRegionIconToString = "olFormRegionIconUnread"
        Case olFormRegionIconRead: OlFormRegionIconToString = "olFormRegionIconRead"
        Case olFormRegionIconReplied: OlFormRegionIconToString = "olFormRegionIconReplied"
        Case olFormRegionIconForwarded: OlFormRegionIconToString = "olFormRegionIconForwarded"
        Case olFormRegionIconUnsent: OlFormRegionIconToString = "olFormRegionIconUnsent"
        Case olFormRegionIconSubmitted: OlFormRegionIconToString = "olFormRegionIconSubmitted"
        Case olFormRegionIconSigned: OlFormRegionIconToString = "olFormRegionIconSigned"
        Case olFormRegionIconEncrypted: OlFormRegionIconToString = "olFormRegionIconEncrypted"
        Case olFormRegionIconWindow: OlFormRegionIconToString = "olFormRegionIconWindow"
        Case olFormRegionIconPage: OlFormRegionIconToString = "olFormRegionIconPage"
        Case olFormRegionIconRecurring: OlFormRegionIconToString = "olFormRegionIconRecurring"
    End Select
End Function
