Attribute VB_Name = "wMsoSyncConflictResolutionType"
Function MsoSyncConflictResolutionTypeFromString(value As String) As MsoSyncConflictResolutionType
    If IsNumeric(value) Then
        MsoSyncConflictResolutionTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoSyncConflictClientWins": MsoSyncConflictResolutionTypeFromString = msoSyncConflictClientWins
        Case "msoSyncConflictServerWins": MsoSyncConflictResolutionTypeFromString = msoSyncConflictServerWins
        Case "msoSyncConflictMerge": MsoSyncConflictResolutionTypeFromString = msoSyncConflictMerge
    End Select
End Function

Function MsoSyncConflictResolutionTypeToString(value As MsoSyncConflictResolutionType) As String
    Select Case value
        Case msoSyncConflictClientWins: MsoSyncConflictResolutionTypeToString = "msoSyncConflictClientWins"
        Case msoSyncConflictServerWins: MsoSyncConflictResolutionTypeToString = "msoSyncConflictServerWins"
        Case msoSyncConflictMerge: MsoSyncConflictResolutionTypeToString = "msoSyncConflictMerge"
    End Select
End Function
