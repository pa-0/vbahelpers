Attribute VB_Name = "wOlAlign"
Function OlAlignFromString(value As String) As OlAlign
    If IsNumeric(value) Then
        OlAlignFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "olAlignLeft": OlAlignFromString = olAlignLeft
        Case "olAlignCenter": OlAlignFromString = olAlignCenter
        Case "olAlignRight": OlAlignFromString = olAlignRight
    End Select
End Function

Function OlAlignToString(value As OlAlign) As String
    Select Case value
        Case olAlignLeft: OlAlignToString = "olAlignLeft"
        Case olAlignCenter: OlAlignToString = "olAlignCenter"
        Case olAlignRight: OlAlignToString = "olAlignRight"
    End Select
End Function
