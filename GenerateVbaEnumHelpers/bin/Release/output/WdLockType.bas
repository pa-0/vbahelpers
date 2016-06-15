Attribute VB_Name = "wWdLockType"
Function WdLockTypeFromString(value As String) As WdLockType
    If IsNumeric(value) Then
        WdLockTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdLockNone": WdLockTypeFromString = wdLockNone
        Case "wdLockReservation": WdLockTypeFromString = wdLockReservation
        Case "wdLockEphemeral": WdLockTypeFromString = wdLockEphemeral
        Case "wdLockChanged": WdLockTypeFromString = wdLockChanged
    End Select
End Function

Function WdLockTypeToString(value As WdLockType) As String
    Select Case value
        Case wdLockNone: WdLockTypeToString = "wdLockNone"
        Case wdLockReservation: WdLockTypeToString = "wdLockReservation"
        Case wdLockEphemeral: WdLockTypeToString = "wdLockEphemeral"
        Case wdLockChanged: WdLockTypeToString = "wdLockChanged"
    End Select
End Function
