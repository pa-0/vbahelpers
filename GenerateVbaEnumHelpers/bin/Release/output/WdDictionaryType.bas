Attribute VB_Name = "wWdDictionaryType"
Function WdDictionaryTypeFromString(value As String) As WdDictionaryType
    If IsNumeric(value) Then
        WdDictionaryTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdSpelling": WdDictionaryTypeFromString = wdSpelling
        Case "wdGrammar": WdDictionaryTypeFromString = wdGrammar
        Case "wdThesaurus": WdDictionaryTypeFromString = wdThesaurus
        Case "wdHyphenation": WdDictionaryTypeFromString = wdHyphenation
        Case "wdSpellingComplete": WdDictionaryTypeFromString = wdSpellingComplete
        Case "wdSpellingCustom": WdDictionaryTypeFromString = wdSpellingCustom
        Case "wdSpellingLegal": WdDictionaryTypeFromString = wdSpellingLegal
        Case "wdSpellingMedical": WdDictionaryTypeFromString = wdSpellingMedical
        Case "wdHangulHanjaConversion": WdDictionaryTypeFromString = wdHangulHanjaConversion
        Case "wdHangulHanjaConversionCustom": WdDictionaryTypeFromString = wdHangulHanjaConversionCustom
    End Select
End Function

Function WdDictionaryTypeToString(value As WdDictionaryType) As String
    Select Case value
        Case wdSpelling: WdDictionaryTypeToString = "wdSpelling"
        Case wdGrammar: WdDictionaryTypeToString = "wdGrammar"
        Case wdThesaurus: WdDictionaryTypeToString = "wdThesaurus"
        Case wdHyphenation: WdDictionaryTypeToString = "wdHyphenation"
        Case wdSpellingComplete: WdDictionaryTypeToString = "wdSpellingComplete"
        Case wdSpellingCustom: WdDictionaryTypeToString = "wdSpellingCustom"
        Case wdSpellingLegal: WdDictionaryTypeToString = "wdSpellingLegal"
        Case wdSpellingMedical: WdDictionaryTypeToString = "wdSpellingMedical"
        Case wdHangulHanjaConversion: WdDictionaryTypeToString = "wdHangulHanjaConversion"
        Case wdHangulHanjaConversionCustom: WdDictionaryTypeToString = "wdHangulHanjaConversionCustom"
    End Select
End Function
