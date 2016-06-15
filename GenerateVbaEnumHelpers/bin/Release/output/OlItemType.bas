Attribute VB_Name = "wOlItemType"
Function OlItemTypeFromString(value As String) As OlItemType
    If IsNumeric(value) Then
        OlItemTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "olMailItem": OlItemTypeFromString = olMailItem
        Case "olAppointmentItem": OlItemTypeFromString = olAppointmentItem
        Case "olContactItem": OlItemTypeFromString = olContactItem
        Case "olTaskItem": OlItemTypeFromString = olTaskItem
        Case "olJournalItem": OlItemTypeFromString = olJournalItem
        Case "olNoteItem": OlItemTypeFromString = olNoteItem
        Case "olPostItem": OlItemTypeFromString = olPostItem
        Case "olDistributionListItem": OlItemTypeFromString = olDistributionListItem
        Case "olMobileItemSMS": OlItemTypeFromString = olMobileItemSMS
        Case "olMobileItemMMS": OlItemTypeFromString = olMobileItemMMS
    End Select
End Function

Function OlItemTypeToString(value As OlItemType) As String
    Select Case value
        Case olMailItem: OlItemTypeToString = "olMailItem"
        Case olAppointmentItem: OlItemTypeToString = "olAppointmentItem"
        Case olContactItem: OlItemTypeToString = "olContactItem"
        Case olTaskItem: OlItemTypeToString = "olTaskItem"
        Case olJournalItem: OlItemTypeToString = "olJournalItem"
        Case olNoteItem: OlItemTypeToString = "olNoteItem"
        Case olPostItem: OlItemTypeToString = "olPostItem"
        Case olDistributionListItem: OlItemTypeToString = "olDistributionListItem"
        Case olMobileItemSMS: OlItemTypeToString = "olMobileItemSMS"
        Case olMobileItemMMS: OlItemTypeToString = "olMobileItemMMS"
    End Select
End Function
