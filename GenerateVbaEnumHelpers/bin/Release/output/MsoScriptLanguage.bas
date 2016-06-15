Attribute VB_Name = "wMsoScriptLanguage"
Function MsoScriptLanguageFromString(value As String) As MsoScriptLanguage
    If IsNumeric(value) Then
        MsoScriptLanguageFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoScriptLanguageJava": MsoScriptLanguageFromString = msoScriptLanguageJava
        Case "msoScriptLanguageVisualBasic": MsoScriptLanguageFromString = msoScriptLanguageVisualBasic
        Case "msoScriptLanguageASP": MsoScriptLanguageFromString = msoScriptLanguageASP
        Case "msoScriptLanguageOther": MsoScriptLanguageFromString = msoScriptLanguageOther
    End Select
End Function

Function MsoScriptLanguageToString(value As MsoScriptLanguage) As String
    Select Case value
        Case msoScriptLanguageJava: MsoScriptLanguageToString = "msoScriptLanguageJava"
        Case msoScriptLanguageVisualBasic: MsoScriptLanguageToString = "msoScriptLanguageVisualBasic"
        Case msoScriptLanguageASP: MsoScriptLanguageToString = "msoScriptLanguageASP"
        Case msoScriptLanguageOther: MsoScriptLanguageToString = "msoScriptLanguageOther"
    End Select
End Function
