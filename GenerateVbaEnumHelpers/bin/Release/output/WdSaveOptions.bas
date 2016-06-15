Attribute VB_Name = "wWdSaveOptions"
Function WdSaveOptionsFromString(value As String) As WdSaveOptions
    If IsNumeric(value) Then
        WdSaveOptionsFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdDoNotSaveChanges": WdSaveOptionsFromString = wdDoNotSaveChanges
        Case "wdPromptToSaveChanges": WdSaveOptionsFromString = wdPromptToSaveChanges
        Case "wdSaveChanges": WdSaveOptionsFromString = wdSaveChanges
    End Select
End Function

Function WdSaveOptionsToString(value As WdSaveOptions) As String
    Select Case value
        Case wdDoNotSaveChanges: WdSaveOptionsToString = "wdDoNotSaveChanges"
        Case wdPromptToSaveChanges: WdSaveOptionsToString = "wdPromptToSaveChanges"
        Case wdSaveChanges: WdSaveOptionsToString = "wdSaveChanges"
    End Select
End Function
