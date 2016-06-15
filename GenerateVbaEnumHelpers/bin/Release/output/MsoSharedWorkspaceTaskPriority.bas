Attribute VB_Name = "wMsoSharedWorkspaceTaskPriority"
Function MsoSharedWorkspaceTaskPriorityFromString(value As String) As MsoSharedWorkspaceTaskPriority
    If IsNumeric(value) Then
        MsoSharedWorkspaceTaskPriorityFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoSharedWorkspaceTaskPriorityHigh": MsoSharedWorkspaceTaskPriorityFromString = msoSharedWorkspaceTaskPriorityHigh
        Case "msoSharedWorkspaceTaskPriorityNormal": MsoSharedWorkspaceTaskPriorityFromString = msoSharedWorkspaceTaskPriorityNormal
        Case "msoSharedWorkspaceTaskPriorityLow": MsoSharedWorkspaceTaskPriorityFromString = msoSharedWorkspaceTaskPriorityLow
    End Select
End Function

Function MsoSharedWorkspaceTaskPriorityToString(value As MsoSharedWorkspaceTaskPriority) As String
    Select Case value
        Case msoSharedWorkspaceTaskPriorityHigh: MsoSharedWorkspaceTaskPriorityToString = "msoSharedWorkspaceTaskPriorityHigh"
        Case msoSharedWorkspaceTaskPriorityNormal: MsoSharedWorkspaceTaskPriorityToString = "msoSharedWorkspaceTaskPriorityNormal"
        Case msoSharedWorkspaceTaskPriorityLow: MsoSharedWorkspaceTaskPriorityToString = "msoSharedWorkspaceTaskPriorityLow"
    End Select
End Function
