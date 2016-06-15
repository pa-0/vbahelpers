Attribute VB_Name = "wOlInspectorClose"
Function OlInspectorCloseFromString(value As String) As OlInspectorClose
    If IsNumeric(value) Then
        OlInspectorCloseFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "olSave": OlInspectorCloseFromString = olSave
        Case "olDiscard": OlInspectorCloseFromString = olDiscard
        Case "olPromptForSave": OlInspectorCloseFromString = olPromptForSave
    End Select
End Function

Function OlInspectorCloseToString(value As OlInspectorClose) As String
    Select Case value
        Case olSave: OlInspectorCloseToString = "olSave"
        Case olDiscard: OlInspectorCloseToString = "olDiscard"
        Case olPromptForSave: OlInspectorCloseToString = "olPromptForSave"
    End Select
End Function
