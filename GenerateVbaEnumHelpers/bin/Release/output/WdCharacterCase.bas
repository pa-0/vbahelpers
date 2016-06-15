Attribute VB_Name = "wWdCharacterCase"
Function WdCharacterCaseFromString(value As String) As WdCharacterCase
    If IsNumeric(value) Then
        WdCharacterCaseFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdLowerCase": WdCharacterCaseFromString = wdLowerCase
        Case "wdUpperCase": WdCharacterCaseFromString = wdUpperCase
        Case "wdTitleWord": WdCharacterCaseFromString = wdTitleWord
        Case "wdTitleSentence": WdCharacterCaseFromString = wdTitleSentence
        Case "wdToggleCase": WdCharacterCaseFromString = wdToggleCase
        Case "wdHalfWidth": WdCharacterCaseFromString = wdHalfWidth
        Case "wdFullWidth": WdCharacterCaseFromString = wdFullWidth
        Case "wdKatakana": WdCharacterCaseFromString = wdKatakana
        Case "wdHiragana": WdCharacterCaseFromString = wdHiragana
        Case "wdNextCase": WdCharacterCaseFromString = wdNextCase
    End Select
End Function

Function WdCharacterCaseToString(value As WdCharacterCase) As String
    Select Case value
        Case wdLowerCase: WdCharacterCaseToString = "wdLowerCase"
        Case wdUpperCase: WdCharacterCaseToString = "wdUpperCase"
        Case wdTitleWord: WdCharacterCaseToString = "wdTitleWord"
        Case wdTitleSentence: WdCharacterCaseToString = "wdTitleSentence"
        Case wdToggleCase: WdCharacterCaseToString = "wdToggleCase"
        Case wdHalfWidth: WdCharacterCaseToString = "wdHalfWidth"
        Case wdFullWidth: WdCharacterCaseToString = "wdFullWidth"
        Case wdKatakana: WdCharacterCaseToString = "wdKatakana"
        Case wdHiragana: WdCharacterCaseToString = "wdHiragana"
        Case wdNextCase: WdCharacterCaseToString = "wdNextCase"
    End Select
End Function
