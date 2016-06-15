Attribute VB_Name = "wWdHeadingSeparator"
Function WdHeadingSeparatorFromString(value As String) As WdHeadingSeparator
    If IsNumeric(value) Then
        WdHeadingSeparatorFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdHeadingSeparatorNone": WdHeadingSeparatorFromString = wdHeadingSeparatorNone
        Case "wdHeadingSeparatorBlankLine": WdHeadingSeparatorFromString = wdHeadingSeparatorBlankLine
        Case "wdHeadingSeparatorLetter": WdHeadingSeparatorFromString = wdHeadingSeparatorLetter
        Case "wdHeadingSeparatorLetterLow": WdHeadingSeparatorFromString = wdHeadingSeparatorLetterLow
        Case "wdHeadingSeparatorLetterFull": WdHeadingSeparatorFromString = wdHeadingSeparatorLetterFull
    End Select
End Function

Function WdHeadingSeparatorToString(value As WdHeadingSeparator) As String
    Select Case value
        Case wdHeadingSeparatorNone: WdHeadingSeparatorToString = "wdHeadingSeparatorNone"
        Case wdHeadingSeparatorBlankLine: WdHeadingSeparatorToString = "wdHeadingSeparatorBlankLine"
        Case wdHeadingSeparatorLetter: WdHeadingSeparatorToString = "wdHeadingSeparatorLetter"
        Case wdHeadingSeparatorLetterLow: WdHeadingSeparatorToString = "wdHeadingSeparatorLetterLow"
        Case wdHeadingSeparatorLetterFull: WdHeadingSeparatorToString = "wdHeadingSeparatorLetterFull"
    End Select
End Function
