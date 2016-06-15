Attribute VB_Name = "wOlSyncState"
Function OlSyncStateFromString(value As String) As OlSyncState
    If IsNumeric(value) Then
        OlSyncStateFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "olSyncStopped": OlSyncStateFromString = olSyncStopped
        Case "olSyncStarted": OlSyncStateFromString = olSyncStarted
    End Select
End Function

Function OlSyncStateToString(value As OlSyncState) As String
    Select Case value
        Case olSyncStopped: OlSyncStateToString = "olSyncStopped"
        Case olSyncStarted: OlSyncStateToString = "olSyncStarted"
    End Select
End Function
