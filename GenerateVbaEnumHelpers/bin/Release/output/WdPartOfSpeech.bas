Attribute VB_Name = "wWdPartOfSpeech"
Function WdPartOfSpeechFromString(value As String) As WdPartOfSpeech
    If IsNumeric(value) Then
        WdPartOfSpeechFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdAdjective": WdPartOfSpeechFromString = wdAdjective
        Case "wdNoun": WdPartOfSpeechFromString = wdNoun
        Case "wdAdverb": WdPartOfSpeechFromString = wdAdverb
        Case "wdVerb": WdPartOfSpeechFromString = wdVerb
        Case "wdPronoun": WdPartOfSpeechFromString = wdPronoun
        Case "wdConjunction": WdPartOfSpeechFromString = wdConjunction
        Case "wdPreposition": WdPartOfSpeechFromString = wdPreposition
        Case "wdInterjection": WdPartOfSpeechFromString = wdInterjection
        Case "wdIdiom": WdPartOfSpeechFromString = wdIdiom
        Case "wdOther": WdPartOfSpeechFromString = wdOther
    End Select
End Function

Function WdPartOfSpeechToString(value As WdPartOfSpeech) As String
    Select Case value
        Case wdAdjective: WdPartOfSpeechToString = "wdAdjective"
        Case wdNoun: WdPartOfSpeechToString = "wdNoun"
        Case wdAdverb: WdPartOfSpeechToString = "wdAdverb"
        Case wdVerb: WdPartOfSpeechToString = "wdVerb"
        Case wdPronoun: WdPartOfSpeechToString = "wdPronoun"
        Case wdConjunction: WdPartOfSpeechToString = "wdConjunction"
        Case wdPreposition: WdPartOfSpeechToString = "wdPreposition"
        Case wdInterjection: WdPartOfSpeechToString = "wdInterjection"
        Case wdIdiom: WdPartOfSpeechToString = "wdIdiom"
        Case wdOther: WdPartOfSpeechToString = "wdOther"
    End Select
End Function
