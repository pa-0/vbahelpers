Attribute VB_Name = "wOlSpecialFolders"
Function OlSpecialFoldersFromString(value As String) As OlSpecialFolders
    If IsNumeric(value) Then
        OlSpecialFoldersFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "olSpecialFolderAllTasks": OlSpecialFoldersFromString = olSpecialFolderAllTasks
        Case "olSpecialFolderReminders": OlSpecialFoldersFromString = olSpecialFolderReminders
    End Select
End Function

Function OlSpecialFoldersToString(value As OlSpecialFolders) As String
    Select Case value
        Case olSpecialFolderAllTasks: OlSpecialFoldersToString = "olSpecialFolderAllTasks"
        Case olSpecialFolderReminders: OlSpecialFoldersToString = "olSpecialFolderReminders"
    End Select
End Function
