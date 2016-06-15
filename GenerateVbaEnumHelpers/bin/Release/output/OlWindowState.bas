Attribute VB_Name = "wOlWindowState"
Function OlWindowStateFromString(value As String) As OlWindowState
    If IsNumeric(value) Then
        OlWindowStateFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "olMaximized": OlWindowStateFromString = olMaximized
        Case "olMinimized": OlWindowStateFromString = olMinimized
        Case "olNormalWindow": OlWindowStateFromString = olNormalWindow
    End Select
End Function

Function OlWindowStateToString(value As OlWindowState) As String
    Select Case value
        Case olMaximized: OlWindowStateToString = "olMaximized"
        Case olMinimized: OlWindowStateToString = "olMinimized"
        Case olNormalWindow: OlWindowStateToString = "olNormalWindow"
    End Select
End Function
