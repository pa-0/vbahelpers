Attribute VB_Name = "wWdTrailingCharacter"
Function WdTrailingCharacterFromString(value As String) As WdTrailingCharacter
    If IsNumeric(value) Then
        WdTrailingCharacterFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdTrailingTab": WdTrailingCharacterFromString = wdTrailingTab
        Case "wdTrailingSpace": WdTrailingCharacterFromString = wdTrailingSpace
        Case "wdTrailingNone": WdTrailingCharacterFromString = wdTrailingNone
    End Select
End Function

Function WdTrailingCharacterToString(value As WdTrailingCharacter) As String
    Select Case value
        Case wdTrailingTab: WdTrailingCharacterToString = "wdTrailingTab"
        Case wdTrailingSpace: WdTrailingCharacterToString = "wdTrailingSpace"
        Case wdTrailingNone: WdTrailingCharacterToString = "wdTrailingNone"
    End Select
End Function
