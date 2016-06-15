Attribute VB_Name = "wOlShiftState"
Function OlShiftStateFromString(value As String) As OlShiftState
    If IsNumeric(value) Then
        OlShiftStateFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "olShiftStateShiftMask": OlShiftStateFromString = olShiftStateShiftMask
        Case "olShiftStateCtrlMask": OlShiftStateFromString = olShiftStateCtrlMask
        Case "olShiftStateAltMask": OlShiftStateFromString = olShiftStateAltMask
    End Select
End Function

Function OlShiftStateToString(value As OlShiftState) As String
    Select Case value
        Case olShiftStateShiftMask: OlShiftStateToString = "olShiftStateShiftMask"
        Case olShiftStateCtrlMask: OlShiftStateToString = "olShiftStateCtrlMask"
        Case olShiftStateAltMask: OlShiftStateToString = "olShiftStateAltMask"
    End Select
End Function
