Attribute VB_Name = "wOlAutoDiscoverConnectionMode"
Function OlAutoDiscoverConnectionModeFromString(value As String) As OlAutoDiscoverConnectionMode
    If IsNumeric(value) Then
        OlAutoDiscoverConnectionModeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "olAutoDiscoverConnectionUnknown": OlAutoDiscoverConnectionModeFromString = olAutoDiscoverConnectionUnknown
        Case "olAutoDiscoverConnectionExternal": OlAutoDiscoverConnectionModeFromString = olAutoDiscoverConnectionExternal
        Case "olAutoDiscoverConnectionInternal": OlAutoDiscoverConnectionModeFromString = olAutoDiscoverConnectionInternal
        Case "olAutoDiscoverConnectionInternalDomain": OlAutoDiscoverConnectionModeFromString = olAutoDiscoverConnectionInternalDomain
    End Select
End Function

Function OlAutoDiscoverConnectionModeToString(value As OlAutoDiscoverConnectionMode) As String
    Select Case value
        Case olAutoDiscoverConnectionUnknown: OlAutoDiscoverConnectionModeToString = "olAutoDiscoverConnectionUnknown"
        Case olAutoDiscoverConnectionExternal: OlAutoDiscoverConnectionModeToString = "olAutoDiscoverConnectionExternal"
        Case olAutoDiscoverConnectionInternal: OlAutoDiscoverConnectionModeToString = "olAutoDiscoverConnectionInternal"
        Case olAutoDiscoverConnectionInternalDomain: OlAutoDiscoverConnectionModeToString = "olAutoDiscoverConnectionInternalDomain"
    End Select
End Function
