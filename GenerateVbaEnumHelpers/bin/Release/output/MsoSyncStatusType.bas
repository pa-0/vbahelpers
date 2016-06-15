Attribute VB_Name = "wMsoSyncStatusType"
Function MsoSyncStatusTypeFromString(value As String) As MsoSyncStatusType
    If IsNumeric(value) Then
        MsoSyncStatusTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoSyncStatusNoSharedWorkspace": MsoSyncStatusTypeFromString = msoSyncStatusNoSharedWorkspace
        Case "msoSyncStatusNotRoaming": MsoSyncStatusTypeFromString = msoSyncStatusNotRoaming
        Case "msoSyncStatusLatest": MsoSyncStatusTypeFromString = msoSyncStatusLatest
        Case "msoSyncStatusNewerAvailable": MsoSyncStatusTypeFromString = msoSyncStatusNewerAvailable
        Case "msoSyncStatusLocalChanges": MsoSyncStatusTypeFromString = msoSyncStatusLocalChanges
        Case "msoSyncStatusConflict": MsoSyncStatusTypeFromString = msoSyncStatusConflict
        Case "msoSyncStatusSuspended": MsoSyncStatusTypeFromString = msoSyncStatusSuspended
        Case "msoSyncStatusError": MsoSyncStatusTypeFromString = msoSyncStatusError
    End Select
End Function

Function MsoSyncStatusTypeToString(value As MsoSyncStatusType) As String
    Select Case value
        Case msoSyncStatusNoSharedWorkspace: MsoSyncStatusTypeToString = "msoSyncStatusNoSharedWorkspace"
        Case msoSyncStatusNotRoaming: MsoSyncStatusTypeToString = "msoSyncStatusNotRoaming"
        Case msoSyncStatusLatest: MsoSyncStatusTypeToString = "msoSyncStatusLatest"
        Case msoSyncStatusNewerAvailable: MsoSyncStatusTypeToString = "msoSyncStatusNewerAvailable"
        Case msoSyncStatusLocalChanges: MsoSyncStatusTypeToString = "msoSyncStatusLocalChanges"
        Case msoSyncStatusConflict: MsoSyncStatusTypeToString = "msoSyncStatusConflict"
        Case msoSyncStatusSuspended: MsoSyncStatusTypeToString = "msoSyncStatusSuspended"
        Case msoSyncStatusError: MsoSyncStatusTypeToString = "msoSyncStatusError"
    End Select
End Function
