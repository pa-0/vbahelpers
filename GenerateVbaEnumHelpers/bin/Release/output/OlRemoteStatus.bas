Attribute VB_Name = "wOlRemoteStatus"
Function OlRemoteStatusFromString(value As String) As OlRemoteStatus
    If IsNumeric(value) Then
        OlRemoteStatusFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "olRemoteStatusNone": OlRemoteStatusFromString = olRemoteStatusNone
        Case "olUnMarked": OlRemoteStatusFromString = olUnMarked
        Case "olMarkedForDownload": OlRemoteStatusFromString = olMarkedForDownload
        Case "olMarkedForCopy": OlRemoteStatusFromString = olMarkedForCopy
        Case "olMarkedForDelete": OlRemoteStatusFromString = olMarkedForDelete
    End Select
End Function

Function OlRemoteStatusToString(value As OlRemoteStatus) As String
    Select Case value
        Case olRemoteStatusNone: OlRemoteStatusToString = "olRemoteStatusNone"
        Case olUnMarked: OlRemoteStatusToString = "olUnMarked"
        Case olMarkedForDownload: OlRemoteStatusToString = "olMarkedForDownload"
        Case olMarkedForCopy: OlRemoteStatusToString = "olMarkedForCopy"
        Case olMarkedForDelete: OlRemoteStatusToString = "olMarkedForDelete"
    End Select
End Function
