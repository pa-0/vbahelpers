Attribute VB_Name = "wPbSaveOptions"
Function PbSaveOptionsFromString(value As String) As PbSaveOptions
    If IsNumeric(value) Then
        PbSaveOptionsFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "pbPromptToSaveChanges": PbSaveOptionsFromString = pbPromptToSaveChanges
        Case "pbSaveChanges": PbSaveOptionsFromString = pbSaveChanges
        Case "pbDoNotSaveChanges": PbSaveOptionsFromString = pbDoNotSaveChanges
    End Select
End Function

Function PbSaveOptionsToString(value As PbSaveOptions) As String
    Select Case value
        Case pbPromptToSaveChanges: PbSaveOptionsToString = "pbPromptToSaveChanges"
        Case pbSaveChanges: PbSaveOptionsToString = "pbSaveChanges"
        Case pbDoNotSaveChanges: PbSaveOptionsToString = "pbDoNotSaveChanges"
    End Select
End Function
