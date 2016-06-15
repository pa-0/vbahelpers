Attribute VB_Name = "wOlActionShowOn"
Function OlActionShowOnFromString(value As String) As OlActionShowOn
    If IsNumeric(value) Then
        OlActionShowOnFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "olDontShow": OlActionShowOnFromString = olDontShow
        Case "olMenu": OlActionShowOnFromString = olMenu
        Case "olMenuAndToolbar": OlActionShowOnFromString = olMenuAndToolbar
    End Select
End Function

Function OlActionShowOnToString(value As OlActionShowOn) As String
    Select Case value
        Case olDontShow: OlActionShowOnToString = "olDontShow"
        Case olMenu: OlActionShowOnToString = "olMenu"
        Case olMenuAndToolbar: OlActionShowOnToString = "olMenuAndToolbar"
    End Select
End Function
