Attribute VB_Name = "wWdMovementType"
Function WdMovementTypeFromString(value As String) As WdMovementType
    If IsNumeric(value) Then
        WdMovementTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdMove": WdMovementTypeFromString = wdMove
        Case "wdExtend": WdMovementTypeFromString = wdExtend
    End Select
End Function

Function WdMovementTypeToString(value As WdMovementType) As String
    Select Case value
        Case wdMove: WdMovementTypeToString = "wdMove"
        Case wdExtend: WdMovementTypeToString = "wdExtend"
    End Select
End Function
