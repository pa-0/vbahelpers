Attribute VB_Name = "wOlTaskOwnership"
Function OlTaskOwnershipFromString(value As String) As OlTaskOwnership
    If IsNumeric(value) Then
        OlTaskOwnershipFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "olNewTask": OlTaskOwnershipFromString = olNewTask
        Case "olDelegatedTask": OlTaskOwnershipFromString = olDelegatedTask
        Case "olOwnTask": OlTaskOwnershipFromString = olOwnTask
    End Select
End Function

Function OlTaskOwnershipToString(value As OlTaskOwnership) As String
    Select Case value
        Case olNewTask: OlTaskOwnershipToString = "olNewTask"
        Case olDelegatedTask: OlTaskOwnershipToString = "olDelegatedTask"
        Case olOwnTask: OlTaskOwnershipToString = "olOwnTask"
    End Select
End Function
