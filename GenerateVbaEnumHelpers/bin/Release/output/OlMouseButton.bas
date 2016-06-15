Attribute VB_Name = "wOlMouseButton"
Function OlMouseButtonFromString(value As String) As OlMouseButton
    If IsNumeric(value) Then
        OlMouseButtonFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "olMouseButtonLeft": OlMouseButtonFromString = olMouseButtonLeft
        Case "olMouseButtonRight": OlMouseButtonFromString = olMouseButtonRight
        Case "olMouseButtonMiddle": OlMouseButtonFromString = olMouseButtonMiddle
    End Select
End Function

Function OlMouseButtonToString(value As OlMouseButton) As String
    Select Case value
        Case olMouseButtonLeft: OlMouseButtonToString = "olMouseButtonLeft"
        Case olMouseButtonRight: OlMouseButtonToString = "olMouseButtonRight"
        Case olMouseButtonMiddle: OlMouseButtonToString = "olMouseButtonMiddle"
    End Select
End Function
