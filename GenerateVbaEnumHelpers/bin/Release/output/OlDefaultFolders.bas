Attribute VB_Name = "wOlDefaultFolders"
Function OlDefaultFoldersFromString(value As String) As OlDefaultFolders
    If IsNumeric(value) Then
        OlDefaultFoldersFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "olFolderDeletedItems": OlDefaultFoldersFromString = olFolderDeletedItems
        Case "olFolderOutbox": OlDefaultFoldersFromString = olFolderOutbox
        Case "olFolderSentMail": OlDefaultFoldersFromString = olFolderSentMail
        Case "olFolderInbox": OlDefaultFoldersFromString = olFolderInbox
        Case "olFolderCalendar": OlDefaultFoldersFromString = olFolderCalendar
        Case "olFolderContacts": OlDefaultFoldersFromString = olFolderContacts
        Case "olFolderJournal": OlDefaultFoldersFromString = olFolderJournal
        Case "olFolderNotes": OlDefaultFoldersFromString = olFolderNotes
        Case "olFolderTasks": OlDefaultFoldersFromString = olFolderTasks
        Case "olFolderDrafts": OlDefaultFoldersFromString = olFolderDrafts
        Case "olPublicFoldersAllPublicFolders": OlDefaultFoldersFromString = olPublicFoldersAllPublicFolders
        Case "olFolderConflicts": OlDefaultFoldersFromString = olFolderConflicts
        Case "olFolderSyncIssues": OlDefaultFoldersFromString = olFolderSyncIssues
        Case "olFolderLocalFailures": OlDefaultFoldersFromString = olFolderLocalFailures
        Case "olFolderServerFailures": OlDefaultFoldersFromString = olFolderServerFailures
        Case "olFolderJunk": OlDefaultFoldersFromString = olFolderJunk
        Case "olFolderRssFeeds": OlDefaultFoldersFromString = olFolderRssFeeds
        Case "olFolderToDo": OlDefaultFoldersFromString = olFolderToDo
        Case "olFolderManagedEmail": OlDefaultFoldersFromString = olFolderManagedEmail
        Case "olFolderSuggestedContacts": OlDefaultFoldersFromString = olFolderSuggestedContacts
    End Select
End Function

Function OlDefaultFoldersToString(value As OlDefaultFolders) As String
    Select Case value
        Case olFolderDeletedItems: OlDefaultFoldersToString = "olFolderDeletedItems"
        Case olFolderOutbox: OlDefaultFoldersToString = "olFolderOutbox"
        Case olFolderSentMail: OlDefaultFoldersToString = "olFolderSentMail"
        Case olFolderInbox: OlDefaultFoldersToString = "olFolderInbox"
        Case olFolderCalendar: OlDefaultFoldersToString = "olFolderCalendar"
        Case olFolderContacts: OlDefaultFoldersToString = "olFolderContacts"
        Case olFolderJournal: OlDefaultFoldersToString = "olFolderJournal"
        Case olFolderNotes: OlDefaultFoldersToString = "olFolderNotes"
        Case olFolderTasks: OlDefaultFoldersToString = "olFolderTasks"
        Case olFolderDrafts: OlDefaultFoldersToString = "olFolderDrafts"
        Case olPublicFoldersAllPublicFolders: OlDefaultFoldersToString = "olPublicFoldersAllPublicFolders"
        Case olFolderConflicts: OlDefaultFoldersToString = "olFolderConflicts"
        Case olFolderSyncIssues: OlDefaultFoldersToString = "olFolderSyncIssues"
        Case olFolderLocalFailures: OlDefaultFoldersToString = "olFolderLocalFailures"
        Case olFolderServerFailures: OlDefaultFoldersToString = "olFolderServerFailures"
        Case olFolderJunk: OlDefaultFoldersToString = "olFolderJunk"
        Case olFolderRssFeeds: OlDefaultFoldersToString = "olFolderRssFeeds"
        Case olFolderToDo: OlDefaultFoldersToString = "olFolderToDo"
        Case olFolderManagedEmail: OlDefaultFoldersToString = "olFolderManagedEmail"
        Case olFolderSuggestedContacts: OlDefaultFoldersToString = "olFolderSuggestedContacts"
    End Select
End Function
