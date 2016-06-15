Attribute VB_Name = "wMsoCharacterSet"
Function MsoCharacterSetFromString(value As String) As MsoCharacterSet
    If IsNumeric(value) Then
        MsoCharacterSetFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoCharacterSetArabic": MsoCharacterSetFromString = msoCharacterSetArabic
        Case "msoCharacterSetCyrillic": MsoCharacterSetFromString = msoCharacterSetCyrillic
        Case "msoCharacterSetEnglishWesternEuropeanOtherLatinScript": MsoCharacterSetFromString = msoCharacterSetEnglishWesternEuropeanOtherLatinScript
        Case "msoCharacterSetGreek": MsoCharacterSetFromString = msoCharacterSetGreek
        Case "msoCharacterSetHebrew": MsoCharacterSetFromString = msoCharacterSetHebrew
        Case "msoCharacterSetJapanese": MsoCharacterSetFromString = msoCharacterSetJapanese
        Case "msoCharacterSetKorean": MsoCharacterSetFromString = msoCharacterSetKorean
        Case "msoCharacterSetMultilingualUnicode": MsoCharacterSetFromString = msoCharacterSetMultilingualUnicode
        Case "msoCharacterSetSimplifiedChinese": MsoCharacterSetFromString = msoCharacterSetSimplifiedChinese
        Case "msoCharacterSetThai": MsoCharacterSetFromString = msoCharacterSetThai
        Case "msoCharacterSetTraditionalChinese": MsoCharacterSetFromString = msoCharacterSetTraditionalChinese
        Case "msoCharacterSetVietnamese": MsoCharacterSetFromString = msoCharacterSetVietnamese
    End Select
End Function

Function MsoCharacterSetToString(value As MsoCharacterSet) As String
    Select Case value
        Case msoCharacterSetArabic: MsoCharacterSetToString = "msoCharacterSetArabic"
        Case msoCharacterSetCyrillic: MsoCharacterSetToString = "msoCharacterSetCyrillic"
        Case msoCharacterSetEnglishWesternEuropeanOtherLatinScript: MsoCharacterSetToString = "msoCharacterSetEnglishWesternEuropeanOtherLatinScript"
        Case msoCharacterSetGreek: MsoCharacterSetToString = "msoCharacterSetGreek"
        Case msoCharacterSetHebrew: MsoCharacterSetToString = "msoCharacterSetHebrew"
        Case msoCharacterSetJapanese: MsoCharacterSetToString = "msoCharacterSetJapanese"
        Case msoCharacterSetKorean: MsoCharacterSetToString = "msoCharacterSetKorean"
        Case msoCharacterSetMultilingualUnicode: MsoCharacterSetToString = "msoCharacterSetMultilingualUnicode"
        Case msoCharacterSetSimplifiedChinese: MsoCharacterSetToString = "msoCharacterSetSimplifiedChinese"
        Case msoCharacterSetThai: MsoCharacterSetToString = "msoCharacterSetThai"
        Case msoCharacterSetTraditionalChinese: MsoCharacterSetToString = "msoCharacterSetTraditionalChinese"
        Case msoCharacterSetVietnamese: MsoCharacterSetToString = "msoCharacterSetVietnamese"
    End Select
End Function
