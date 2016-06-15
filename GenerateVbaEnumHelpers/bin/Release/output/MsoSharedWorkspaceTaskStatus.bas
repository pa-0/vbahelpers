Attribute VB_Name = "wMsoSharedWorkspaceTaskStatus"
Function MsoSharedWorkspaceTaskStatusFromString(value As String) As MsoSharedWorkspaceTaskStatus
    If IsNumeric(value) Then
        MsoSharedWorkspaceTaskStatusFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoSharedWorkspaceTaskStatusNotStarted": MsoSharedWorkspaceTaskStatusFromString = msoSharedWorkspaceTaskStatusNotStarted
        Case "msoSharedWorkspaceTaskStatusInProgress": MsoSharedWorkspaceTaskStatusFromString = msoSharedWorkspaceTaskStatusInProgress
        Case "msoSharedWorkspaceTaskStatusCompleted": MsoSharedWorkspaceTaskStatusFromString = msoSharedWorkspaceTaskStatusCompleted
        Case "msoSharedWorkspaceTaskStatusDeferred": MsoSharedWorkspaceTaskStatusFromString = msoSharedWorkspaceTaskStatusDeferred
        Case "msoSharedWorkspaceTaskStatusWaiting": MsoSharedWorkspaceTaskStatusFromString = msoSharedWorkspaceTaskStatusWaiting
    End Select
End Function

Function MsoSharedWorkspaceTaskStatusToString(value As MsoSharedWorkspaceTaskStatus) As String
    Select Case value
        Case msoSharedWorkspaceTaskStatusNotStarted: MsoSharedWorkspaceTaskStatusToString = "msoSharedWorkspaceTaskStatusNotStarted"
        Case msoSharedWorkspaceTaskStatusInProgress: MsoSharedWorkspaceTaskStatusToString = "msoSharedWorkspaceTaskStatusInProgress"
        Case msoSharedWorkspaceTaskStatusCompleted: MsoSharedWorkspaceTaskStatusToString = "msoSharedWorkspaceTaskStatusCompleted"
        Case msoSharedWorkspaceTaskStatusDeferred: MsoSharedWorkspaceTaskStatusToString = "msoSharedWorkspaceTaskStatusDeferred"
        Case msoSharedWorkspaceTaskStatusWaiting: MsoSharedWorkspaceTaskStatusToString = "msoSharedWorkspaceTaskStatusWaiting"
    End Select
End Function
