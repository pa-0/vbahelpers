Attribute VB_Name = "wOlJournalRecipientType"
Function OlJournalRecipientTypeFromString(value As String) As OlJournalRecipientType
    If IsNumeric(value) Then
        OlJournalRecipientTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "olAssociatedContact": OlJournalRecipientTypeFromString = olAssociatedContact
    End Select
End Function

Function OlJournalRecipientTypeToString(value As OlJournalRecipientType) As String
    Select Case value
        Case olAssociatedContact: OlJournalRecipientTypeToString = "olAssociatedContact"
    End Select
End Function
