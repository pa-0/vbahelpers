Attribute VB_Name = "wOlActionResponseStyle"
Function OlActionResponseStyleFromString(value As String) As OlActionResponseStyle
    If IsNumeric(value) Then
        OlActionResponseStyleFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "olOpen": OlActionResponseStyleFromString = olOpen
        Case "olSend": OlActionResponseStyleFromString = olSend
        Case "olPrompt": OlActionResponseStyleFromString = olPrompt
    End Select
End Function

Function OlActionResponseStyleToString(value As OlActionResponseStyle) As String
    Select Case value
        Case olOpen: OlActionResponseStyleToString = "olOpen"
        Case olSend: OlActionResponseStyleToString = "olSend"
        Case olPrompt: OlActionResponseStyleToString = "olPrompt"
    End Select
End Function
