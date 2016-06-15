Attribute VB_Name = "wXlSaveConflictResolution"
Function XlSaveConflictResolutionFromString(value As String) As XlSaveConflictResolution
    If IsNumeric(value) Then
        XlSaveConflictResolutionFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlUserResolution": XlSaveConflictResolutionFromString = xlUserResolution
        Case "xlLocalSessionChanges": XlSaveConflictResolutionFromString = xlLocalSessionChanges
        Case "xlOtherSessionChanges": XlSaveConflictResolutionFromString = xlOtherSessionChanges
    End Select
End Function

Function XlSaveConflictResolutionToString(value As XlSaveConflictResolution) As String
    Select Case value
        Case xlUserResolution: XlSaveConflictResolutionToString = "xlUserResolution"
        Case xlLocalSessionChanges: XlSaveConflictResolutionToString = "xlLocalSessionChanges"
        Case xlOtherSessionChanges: XlSaveConflictResolutionToString = "xlOtherSessionChanges"
    End Select
End Function
