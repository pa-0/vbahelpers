Attribute VB_Name = "wOlExchangeConnectionMode"
Function OlExchangeConnectionModeFromString(value As String) As OlExchangeConnectionMode
    If IsNumeric(value) Then
        OlExchangeConnectionModeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "olNoExchange": OlExchangeConnectionModeFromString = olNoExchange
        Case "olOffline": OlExchangeConnectionModeFromString = olOffline
        Case "olCachedOffline": OlExchangeConnectionModeFromString = olCachedOffline
        Case "olDisconnected": OlExchangeConnectionModeFromString = olDisconnected
        Case "olCachedDisconnected": OlExchangeConnectionModeFromString = olCachedDisconnected
        Case "olCachedConnectedHeaders": OlExchangeConnectionModeFromString = olCachedConnectedHeaders
        Case "olCachedConnectedDrizzle": OlExchangeConnectionModeFromString = olCachedConnectedDrizzle
        Case "olCachedConnectedFull": OlExchangeConnectionModeFromString = olCachedConnectedFull
        Case "olOnline": OlExchangeConnectionModeFromString = olOnline
    End Select
End Function

Function OlExchangeConnectionModeToString(value As OlExchangeConnectionMode) As String
    Select Case value
        Case olNoExchange: OlExchangeConnectionModeToString = "olNoExchange"
        Case olOffline: OlExchangeConnectionModeToString = "olOffline"
        Case olCachedOffline: OlExchangeConnectionModeToString = "olCachedOffline"
        Case olDisconnected: OlExchangeConnectionModeToString = "olDisconnected"
        Case olCachedDisconnected: OlExchangeConnectionModeToString = "olCachedDisconnected"
        Case olCachedConnectedHeaders: OlExchangeConnectionModeToString = "olCachedConnectedHeaders"
        Case olCachedConnectedDrizzle: OlExchangeConnectionModeToString = "olCachedConnectedDrizzle"
        Case olCachedConnectedFull: OlExchangeConnectionModeToString = "olCachedConnectedFull"
        Case olOnline: OlExchangeConnectionModeToString = "olOnline"
    End Select
End Function
