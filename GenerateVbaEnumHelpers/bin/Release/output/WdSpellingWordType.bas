Attribute VB_Name = "wWdSpellingWordType"
Function WdSpellingWordTypeFromString(value As String) As WdSpellingWordType
    If IsNumeric(value) Then
        WdSpellingWordTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdSpellword": WdSpellingWordTypeFromString = wdSpellword
        Case "wdWildcard": WdSpellingWordTypeFromString = wdWildcard
        Case "wdAnagram": WdSpellingWordTypeFromString = wdAnagram
    End Select
End Function

Function WdSpellingWordTypeToString(value As WdSpellingWordType) As String
    Select Case value
        Case wdSpellword: WdSpellingWordTypeToString = "wdSpellword"
        Case wdWildcard: WdSpellingWordTypeToString = "wdWildcard"
        Case wdAnagram: WdSpellingWordTypeToString = "wdAnagram"
    End Select
End Function
