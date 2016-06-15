Attribute VB_Name = "wMsoButtonState"
Function MsoButtonStateFromString(value As String) As MsoButtonState
    If IsNumeric(value) Then
        MsoButtonStateFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoButtonUp": MsoButtonStateFromString = msoButtonUp
        Case "msoButtonMixed": MsoButtonStateFromString = msoButtonMixed
        Case "msoButtonDown": MsoButtonStateFromString = msoButtonDown
    End Select
End Function

Function MsoButtonStateToString(value As MsoButtonState) As String
    Select Case value
        Case msoButtonUp: MsoButtonStateToString = "msoButtonUp"
        Case msoButtonMixed: MsoButtonStateToString = "msoButtonMixed"
        Case msoButtonDown: MsoButtonStateToString = "msoButtonDown"
    End Select
End Function
