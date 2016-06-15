Attribute VB_Name = "wOlAddressListType"
Function OlAddressListTypeFromString(value As String) As OlAddressListType
    If IsNumeric(value) Then
        OlAddressListTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "olExchangeGlobalAddressList": OlAddressListTypeFromString = olExchangeGlobalAddressList
        Case "olExchangeContainer": OlAddressListTypeFromString = olExchangeContainer
        Case "olOutlookAddressList": OlAddressListTypeFromString = olOutlookAddressList
        Case "olOutlookLdapAddressList": OlAddressListTypeFromString = olOutlookLdapAddressList
        Case "olCustomAddressList": OlAddressListTypeFromString = olCustomAddressList
    End Select
End Function

Function OlAddressListTypeToString(value As OlAddressListType) As String
    Select Case value
        Case olExchangeGlobalAddressList: OlAddressListTypeToString = "olExchangeGlobalAddressList"
        Case olExchangeContainer: OlAddressListTypeToString = "olExchangeContainer"
        Case olOutlookAddressList: OlAddressListTypeToString = "olOutlookAddressList"
        Case olOutlookLdapAddressList: OlAddressListTypeToString = "olOutlookLdapAddressList"
        Case olCustomAddressList: OlAddressListTypeToString = "olCustomAddressList"
    End Select
End Function
