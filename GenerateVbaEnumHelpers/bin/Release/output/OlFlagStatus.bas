Attribute VB_Name = "wOlFlagStatus"
Function OlFlagStatusFromString(value As String) As OlFlagStatus
    If IsNumeric(value) Then
        OlFlagStatusFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "olNoFlag": OlFlagStatusFromString = olNoFlag
        Case "olFlagComplete": OlFlagStatusFromString = olFlagComplete
        Case "olFlagMarked": OlFlagStatusFromString = olFlagMarked
    End Select
End Function

Function OlFlagStatusToString(value As OlFlagStatus) As String
    Select Case value
        Case olNoFlag: OlFlagStatusToString = "olNoFlag"
        Case olFlagComplete: OlFlagStatusToString = "olFlagComplete"
        Case olFlagMarked: OlFlagStatusToString = "olFlagMarked"
    End Select
End Function
