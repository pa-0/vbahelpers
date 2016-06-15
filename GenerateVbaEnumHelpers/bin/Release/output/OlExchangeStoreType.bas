Attribute VB_Name = "wOlExchangeStoreType"
Function OlExchangeStoreTypeFromString(value As String) As OlExchangeStoreType
    If IsNumeric(value) Then
        OlExchangeStoreTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "olPrimaryExchangeMailbox": OlExchangeStoreTypeFromString = olPrimaryExchangeMailbox
        Case "olExchangeMailbox": OlExchangeStoreTypeFromString = olExchangeMailbox
        Case "olExchangePublicFolder": OlExchangeStoreTypeFromString = olExchangePublicFolder
        Case "olNotExchange": OlExchangeStoreTypeFromString = olNotExchange
        Case "olAdditionalExchangeMailbox": OlExchangeStoreTypeFromString = olAdditionalExchangeMailbox
    End Select
End Function

Function OlExchangeStoreTypeToString(value As OlExchangeStoreType) As String
    Select Case value
        Case olPrimaryExchangeMailbox: OlExchangeStoreTypeToString = "olPrimaryExchangeMailbox"
        Case olExchangeMailbox: OlExchangeStoreTypeToString = "olExchangeMailbox"
        Case olExchangePublicFolder: OlExchangeStoreTypeToString = "olExchangePublicFolder"
        Case olNotExchange: OlExchangeStoreTypeToString = "olNotExchange"
        Case olAdditionalExchangeMailbox: OlExchangeStoreTypeToString = "olAdditionalExchangeMailbox"
    End Select
End Function
