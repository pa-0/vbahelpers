Attribute VB_Name = "wWdDateLanguage"
Function WdDateLanguageFromString(value As String) As WdDateLanguage
    If IsNumeric(value) Then
        WdDateLanguageFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdDateLanguageBidi": WdDateLanguageFromString = wdDateLanguageBidi
        Case "wdDateLanguageLatin": WdDateLanguageFromString = wdDateLanguageLatin
    End Select
End Function

Function WdDateLanguageToString(value As WdDateLanguage) As String
    Select Case value
        Case wdDateLanguageBidi: WdDateLanguageToString = "wdDateLanguageBidi"
        Case wdDateLanguageLatin: WdDateLanguageToString = "wdDateLanguageLatin"
    End Select
End Function
