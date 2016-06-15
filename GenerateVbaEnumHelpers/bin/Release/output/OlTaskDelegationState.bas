Attribute VB_Name = "wOlTaskDelegationState"
Function OlTaskDelegationStateFromString(value As String) As OlTaskDelegationState
    If IsNumeric(value) Then
        OlTaskDelegationStateFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "olTaskNotDelegated": OlTaskDelegationStateFromString = olTaskNotDelegated
        Case "olTaskDelegationUnknown": OlTaskDelegationStateFromString = olTaskDelegationUnknown
        Case "olTaskDelegationAccepted": OlTaskDelegationStateFromString = olTaskDelegationAccepted
        Case "olTaskDelegationDeclined": OlTaskDelegationStateFromString = olTaskDelegationDeclined
    End Select
End Function

Function OlTaskDelegationStateToString(value As OlTaskDelegationState) As String
    Select Case value
        Case olTaskNotDelegated: OlTaskDelegationStateToString = "olTaskNotDelegated"
        Case olTaskDelegationUnknown: OlTaskDelegationStateToString = "olTaskDelegationUnknown"
        Case olTaskDelegationAccepted: OlTaskDelegationStateToString = "olTaskDelegationAccepted"
        Case olTaskDelegationDeclined: OlTaskDelegationStateToString = "olTaskDelegationDeclined"
    End Select
End Function
