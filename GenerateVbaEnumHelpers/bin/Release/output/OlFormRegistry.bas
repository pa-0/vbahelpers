Attribute VB_Name = "wOlFormRegistry"
Function OlFormRegistryFromString(value As String) As OlFormRegistry
    If IsNumeric(value) Then
        OlFormRegistryFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "olDefaultRegistry": OlFormRegistryFromString = olDefaultRegistry
        Case "olPersonalRegistry": OlFormRegistryFromString = olPersonalRegistry
        Case "olFolderRegistry": OlFormRegistryFromString = olFolderRegistry
        Case "olOrganizationRegistry": OlFormRegistryFromString = olOrganizationRegistry
    End Select
End Function

Function OlFormRegistryToString(value As OlFormRegistry) As String
    Select Case value
        Case olDefaultRegistry: OlFormRegistryToString = "olDefaultRegistry"
        Case olPersonalRegistry: OlFormRegistryToString = "olPersonalRegistry"
        Case olFolderRegistry: OlFormRegistryToString = "olFolderRegistry"
        Case olOrganizationRegistry: OlFormRegistryToString = "olOrganizationRegistry"
    End Select
End Function
