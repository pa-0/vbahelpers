Attribute VB_Name = "wOlTaskStatus"
Function OlTaskStatusFromString(value As String) As OlTaskStatus
    If IsNumeric(value) Then
        OlTaskStatusFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "olTaskNotStarted": OlTaskStatusFromString = olTaskNotStarted
        Case "olTaskInProgress": OlTaskStatusFromString = olTaskInProgress
        Case "olTaskComplete": OlTaskStatusFromString = olTaskComplete
        Case "olTaskWaiting": OlTaskStatusFromString = olTaskWaiting
        Case "olTaskDeferred": OlTaskStatusFromString = olTaskDeferred
    End Select
End Function

Function OlTaskStatusToString(value As OlTaskStatus) As String
    Select Case value
        Case olTaskNotStarted: OlTaskStatusToString = "olTaskNotStarted"
        Case olTaskInProgress: OlTaskStatusToString = "olTaskInProgress"
        Case olTaskComplete: OlTaskStatusToString = "olTaskComplete"
        Case olTaskWaiting: OlTaskStatusToString = "olTaskWaiting"
        Case olTaskDeferred: OlTaskStatusToString = "olTaskDeferred"
    End Select
End Function
