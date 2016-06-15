Attribute VB_Name = "wMsoFarEastLineBreakLanguageID"
Function MsoFarEastLineBreakLanguageIDFromString(value As String) As MsoFarEastLineBreakLanguageID
    If IsNumeric(value) Then
        MsoFarEastLineBreakLanguageIDFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "MsoFarEastLineBreakLanguageTraditionalChinese": MsoFarEastLineBreakLanguageIDFromString = MsoFarEastLineBreakLanguageTraditionalChinese
        Case "MsoFarEastLineBreakLanguageJapanese": MsoFarEastLineBreakLanguageIDFromString = MsoFarEastLineBreakLanguageJapanese
        Case "MsoFarEastLineBreakLanguageKorean": MsoFarEastLineBreakLanguageIDFromString = MsoFarEastLineBreakLanguageKorean
        Case "MsoFarEastLineBreakLanguageSimplifiedChinese": MsoFarEastLineBreakLanguageIDFromString = MsoFarEastLineBreakLanguageSimplifiedChinese
    End Select
End Function

Function MsoFarEastLineBreakLanguageIDToString(value As MsoFarEastLineBreakLanguageID) As String
    Select Case value
        Case MsoFarEastLineBreakLanguageTraditionalChinese: MsoFarEastLineBreakLanguageIDToString = "MsoFarEastLineBreakLanguageTraditionalChinese"
        Case MsoFarEastLineBreakLanguageJapanese: MsoFarEastLineBreakLanguageIDToString = "MsoFarEastLineBreakLanguageJapanese"
        Case MsoFarEastLineBreakLanguageKorean: MsoFarEastLineBreakLanguageIDToString = "MsoFarEastLineBreakLanguageKorean"
        Case MsoFarEastLineBreakLanguageSimplifiedChinese: MsoFarEastLineBreakLanguageIDToString = "MsoFarEastLineBreakLanguageSimplifiedChinese"
    End Select
End Function
