Attribute VB_Name = "wOlNavigationModuleType"
Function OlNavigationModuleTypeFromString(value As String) As OlNavigationModuleType
    If IsNumeric(value) Then
        OlNavigationModuleTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "olModuleMail": OlNavigationModuleTypeFromString = olModuleMail
        Case "olModuleCalendar": OlNavigationModuleTypeFromString = olModuleCalendar
        Case "olModuleContacts": OlNavigationModuleTypeFromString = olModuleContacts
        Case "olModuleTasks": OlNavigationModuleTypeFromString = olModuleTasks
        Case "olModuleJournal": OlNavigationModuleTypeFromString = olModuleJournal
        Case "olModuleNotes": OlNavigationModuleTypeFromString = olModuleNotes
        Case "olModuleFolderList": OlNavigationModuleTypeFromString = olModuleFolderList
        Case "olModuleShortcuts": OlNavigationModuleTypeFromString = olModuleShortcuts
        Case "olModuleSolutions": OlNavigationModuleTypeFromString = olModuleSolutions
    End Select
End Function

Function OlNavigationModuleTypeToString(value As OlNavigationModuleType) As String
    Select Case value
        Case olModuleMail: OlNavigationModuleTypeToString = "olModuleMail"
        Case olModuleCalendar: OlNavigationModuleTypeToString = "olModuleCalendar"
        Case olModuleContacts: OlNavigationModuleTypeToString = "olModuleContacts"
        Case olModuleTasks: OlNavigationModuleTypeToString = "olModuleTasks"
        Case olModuleJournal: OlNavigationModuleTypeToString = "olModuleJournal"
        Case olModuleNotes: OlNavigationModuleTypeToString = "olModuleNotes"
        Case olModuleFolderList: OlNavigationModuleTypeToString = "olModuleFolderList"
        Case olModuleShortcuts: OlNavigationModuleTypeToString = "olModuleShortcuts"
        Case olModuleSolutions: OlNavigationModuleTypeToString = "olModuleSolutions"
    End Select
End Function
