Attribute VB_Name = "wOlAddressEntryUserType"
Function OlAddressEntryUserTypeFromString(value As String) As OlAddressEntryUserType
    If IsNumeric(value) Then
        OlAddressEntryUserTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "olExchangeUserAddressEntry": OlAddressEntryUserTypeFromString = olExchangeUserAddressEntry
        Case "olExchangeDistributionListAddressEntry": OlAddressEntryUserTypeFromString = olExchangeDistributionListAddressEntry
        Case "olExchangePublicFolderAddressEntry": OlAddressEntryUserTypeFromString = olExchangePublicFolderAddressEntry
        Case "olExchangeAgentAddressEntry": OlAddressEntryUserTypeFromString = olExchangeAgentAddressEntry
        Case "olExchangeOrganizationAddressEntry": OlAddressEntryUserTypeFromString = olExchangeOrganizationAddressEntry
        Case "olExchangeRemoteUserAddressEntry": OlAddressEntryUserTypeFromString = olExchangeRemoteUserAddressEntry
        Case "olOutlookContactAddressEntry": OlAddressEntryUserTypeFromString = olOutlookContactAddressEntry
        Case "olOutlookDistributionListAddressEntry": OlAddressEntryUserTypeFromString = olOutlookDistributionListAddressEntry
        Case "olLdapAddressEntry": OlAddressEntryUserTypeFromString = olLdapAddressEntry
        Case "olSmtpAddressEntry": OlAddressEntryUserTypeFromString = olSmtpAddressEntry
        Case "olOtherAddressEntry": OlAddressEntryUserTypeFromString = olOtherAddressEntry
    End Select
End Function

Function OlAddressEntryUserTypeToString(value As OlAddressEntryUserType) As String
    Select Case value
        Case olExchangeUserAddressEntry: OlAddressEntryUserTypeToString = "olExchangeUserAddressEntry"
        Case olExchangeDistributionListAddressEntry: OlAddressEntryUserTypeToString = "olExchangeDistributionListAddressEntry"
        Case olExchangePublicFolderAddressEntry: OlAddressEntryUserTypeToString = "olExchangePublicFolderAddressEntry"
        Case olExchangeAgentAddressEntry: OlAddressEntryUserTypeToString = "olExchangeAgentAddressEntry"
        Case olExchangeOrganizationAddressEntry: OlAddressEntryUserTypeToString = "olExchangeOrganizationAddressEntry"
        Case olExchangeRemoteUserAddressEntry: OlAddressEntryUserTypeToString = "olExchangeRemoteUserAddressEntry"
        Case olOutlookContactAddressEntry: OlAddressEntryUserTypeToString = "olOutlookContactAddressEntry"
        Case olOutlookDistributionListAddressEntry: OlAddressEntryUserTypeToString = "olOutlookDistributionListAddressEntry"
        Case olLdapAddressEntry: OlAddressEntryUserTypeToString = "olLdapAddressEntry"
        Case olSmtpAddressEntry: OlAddressEntryUserTypeToString = "olSmtpAddressEntry"
        Case olOtherAddressEntry: OlAddressEntryUserTypeToString = "olOtherAddressEntry"
    End Select
End Function
