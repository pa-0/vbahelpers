Attribute VB_Name = "wWdCursorMovement"
Function WdCursorMovementFromString(value As String) As WdCursorMovement
    If IsNumeric(value) Then
        WdCursorMovementFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdCursorMovementLogical": WdCursorMovementFromString = wdCursorMovementLogical
        Case "wdCursorMovementVisual": WdCursorMovementFromString = wdCursorMovementVisual
    End Select
End Function

Function WdCursorMovementToString(value As WdCursorMovement) As String
    Select Case value
        Case wdCursorMovementLogical: WdCursorMovementToString = "wdCursorMovementLogical"
        Case wdCursorMovementVisual: WdCursorMovementToString = "wdCursorMovementVisual"
    End Select
End Function
