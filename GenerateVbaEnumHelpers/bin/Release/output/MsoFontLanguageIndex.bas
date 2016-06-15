Attribute VB_Name = "wMsoFontLanguageIndex"
Function MsoFontLanguageIndexFromString(value As String) As MsoFontLanguageIndex
    If IsNumeric(value) Then
        MsoFontLanguageIndexFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoThemeLatin": MsoFontLanguageIndexFromString = msoThemeLatin
        Case "msoThemeComplexScript": MsoFontLanguageIndexFromString = msoThemeComplexScript
        Case "msoThemeEastAsian": MsoFontLanguageIndexFromString = msoThemeEastAsian
    End Select
End Function

Function MsoFontLanguageIndexToString(value As MsoFontLanguageIndex) As String
    Select Case value
        Case msoThemeLatin: MsoFontLanguageIndexToString = "msoThemeLatin"
        Case msoThemeComplexScript: MsoFontLanguageIndexToString = "msoThemeComplexScript"
        Case msoThemeEastAsian: MsoFontLanguageIndexToString = "msoThemeEastAsian"
    End Select
End Function
