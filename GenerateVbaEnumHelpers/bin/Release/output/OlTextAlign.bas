Attribute VB_Name = "wOlTextAlign"
Function OlTextAlignFromString(value As String) As OlTextAlign
    If IsNumeric(value) Then
        OlTextAlignFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "olTextAlignLeft": OlTextAlignFromString = olTextAlignLeft
        Case "olTextAlignCenter": OlTextAlignFromString = olTextAlignCenter
        Case "olTextAlignRight": OlTextAlignFromString = olTextAlignRight
    End Select
End Function

Function OlTextAlignToString(value As OlTextAlign) As String
    Select Case value
        Case olTextAlignLeft: OlTextAlignToString = "olTextAlignLeft"
        Case olTextAlignCenter: OlTextAlignToString = "olTextAlignCenter"
        Case olTextAlignRight: OlTextAlignToString = "olTextAlignRight"
    End Select
End Function
