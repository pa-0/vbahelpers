Attribute VB_Name = "wPpMediaTaskStatus"
Function PpMediaTaskStatusFromString(value As String) As PpMediaTaskStatus
    If IsNumeric(value) Then
        PpMediaTaskStatusFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "ppMediaTaskStatusNone": PpMediaTaskStatusFromString = ppMediaTaskStatusNone
        Case "ppMediaTaskStatusInProgress": PpMediaTaskStatusFromString = ppMediaTaskStatusInProgress
        Case "ppMediaTaskStatusQueued": PpMediaTaskStatusFromString = ppMediaTaskStatusQueued
        Case "ppMediaTaskStatusDone": PpMediaTaskStatusFromString = ppMediaTaskStatusDone
        Case "ppMediaTaskStatusFailed": PpMediaTaskStatusFromString = ppMediaTaskStatusFailed
    End Select
End Function

Function PpMediaTaskStatusToString(value As PpMediaTaskStatus) As String
    Select Case value
        Case ppMediaTaskStatusNone: PpMediaTaskStatusToString = "ppMediaTaskStatusNone"
        Case ppMediaTaskStatusInProgress: PpMediaTaskStatusToString = "ppMediaTaskStatusInProgress"
        Case ppMediaTaskStatusQueued: PpMediaTaskStatusToString = "ppMediaTaskStatusQueued"
        Case ppMediaTaskStatusDone: PpMediaTaskStatusToString = "ppMediaTaskStatusDone"
        Case ppMediaTaskStatusFailed: PpMediaTaskStatusToString = "ppMediaTaskStatusFailed"
    End Select
End Function
