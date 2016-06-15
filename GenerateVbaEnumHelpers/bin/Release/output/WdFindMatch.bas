Attribute VB_Name = "wWdFindMatch"
Function WdFindMatchFromString(value As String) As WdFindMatch
    If IsNumeric(value) Then
        WdFindMatchFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdMatchGraphic": WdFindMatchFromString = wdMatchGraphic
        Case "wdMatchCommentMark": WdFindMatchFromString = wdMatchCommentMark
        Case "wdMatchTabCharacter": WdFindMatchFromString = wdMatchTabCharacter
        Case "wdMatchCaretCharacter": WdFindMatchFromString = wdMatchCaretCharacter
        Case "wdMatchColumnBreak": WdFindMatchFromString = wdMatchColumnBreak
        Case "wdMatchField": WdFindMatchFromString = wdMatchField
        Case "wdMatchNonbreakingHyphen": WdFindMatchFromString = wdMatchNonbreakingHyphen
        Case "wdMatchOptionalHyphen": WdFindMatchFromString = wdMatchOptionalHyphen
        Case "wdMatchNonbreakingSpace": WdFindMatchFromString = wdMatchNonbreakingSpace
        Case "wdMatchEnDash": WdFindMatchFromString = wdMatchEnDash
        Case "wdMatchEmDash": WdFindMatchFromString = wdMatchEmDash
        Case "wdMatchManualLineBreak": WdFindMatchFromString = wdMatchManualLineBreak
        Case "wdMatchParagraphMark": WdFindMatchFromString = wdMatchParagraphMark
        Case "wdMatchFootnoteMark": WdFindMatchFromString = wdMatchFootnoteMark
        Case "wdMatchEndnoteMark": WdFindMatchFromString = wdMatchEndnoteMark
        Case "wdMatchManualPageBreak": WdFindMatchFromString = wdMatchManualPageBreak
        Case "wdMatchAnyDigit": WdFindMatchFromString = wdMatchAnyDigit
        Case "wdMatchSectionBreak": WdFindMatchFromString = wdMatchSectionBreak
        Case "wdMatchAnyLetter": WdFindMatchFromString = wdMatchAnyLetter
        Case "wdMatchAnyCharacter": WdFindMatchFromString = wdMatchAnyCharacter
        Case "wdMatchWhiteSpace": WdFindMatchFromString = wdMatchWhiteSpace
    End Select
End Function

Function WdFindMatchToString(value As WdFindMatch) As String
    Select Case value
        Case wdMatchGraphic: WdFindMatchToString = "wdMatchGraphic"
        Case wdMatchCommentMark: WdFindMatchToString = "wdMatchCommentMark"
        Case wdMatchTabCharacter: WdFindMatchToString = "wdMatchTabCharacter"
        Case wdMatchCaretCharacter: WdFindMatchToString = "wdMatchCaretCharacter"
        Case wdMatchColumnBreak: WdFindMatchToString = "wdMatchColumnBreak"
        Case wdMatchField: WdFindMatchToString = "wdMatchField"
        Case wdMatchNonbreakingHyphen: WdFindMatchToString = "wdMatchNonbreakingHyphen"
        Case wdMatchOptionalHyphen: WdFindMatchToString = "wdMatchOptionalHyphen"
        Case wdMatchNonbreakingSpace: WdFindMatchToString = "wdMatchNonbreakingSpace"
        Case wdMatchEnDash: WdFindMatchToString = "wdMatchEnDash"
        Case wdMatchEmDash: WdFindMatchToString = "wdMatchEmDash"
        Case wdMatchManualLineBreak: WdFindMatchToString = "wdMatchManualLineBreak"
        Case wdMatchParagraphMark: WdFindMatchToString = "wdMatchParagraphMark"
        Case wdMatchFootnoteMark: WdFindMatchToString = "wdMatchFootnoteMark"
        Case wdMatchEndnoteMark: WdFindMatchToString = "wdMatchEndnoteMark"
        Case wdMatchManualPageBreak: WdFindMatchToString = "wdMatchManualPageBreak"
        Case wdMatchAnyDigit: WdFindMatchToString = "wdMatchAnyDigit"
        Case wdMatchSectionBreak: WdFindMatchToString = "wdMatchSectionBreak"
        Case wdMatchAnyLetter: WdFindMatchToString = "wdMatchAnyLetter"
        Case wdMatchAnyCharacter: WdFindMatchToString = "wdMatchAnyCharacter"
        Case wdMatchWhiteSpace: WdFindMatchToString = "wdMatchWhiteSpace"
    End Select
End Function
