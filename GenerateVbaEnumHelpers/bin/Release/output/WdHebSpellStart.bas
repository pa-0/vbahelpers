Attribute VB_Name = "wWdHebSpellStart"
Function WdHebSpellStartFromString(value As String) As WdHebSpellStart
    If IsNumeric(value) Then
        WdHebSpellStartFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdFullScript": WdHebSpellStartFromString = wdFullScript
        Case "wdPartialScript": WdHebSpellStartFromString = wdPartialScript
        Case "wdMixedScript": WdHebSpellStartFromString = wdMixedScript
        Case "wdMixedAuthorizedScript": WdHebSpellStartFromString = wdMixedAuthorizedScript
    End Select
End Function

Function WdHebSpellStartToString(value As WdHebSpellStart) As String
    Select Case value
        Case wdFullScript: WdHebSpellStartToString = "wdFullScript"
        Case wdPartialScript: WdHebSpellStartToString = "wdPartialScript"
        Case wdMixedScript: WdHebSpellStartToString = "wdMixedScript"
        Case wdMixedAuthorizedScript: WdHebSpellStartToString = "wdMixedAuthorizedScript"
    End Select
End Function
