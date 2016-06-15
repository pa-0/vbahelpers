Attribute VB_Name = "wWdFarEastLineBreakLanguageID"
Function WdFarEastLineBreakLanguageIDFromString(value As String) As WdFarEastLineBreakLanguageID
    If IsNumeric(value) Then
        WdFarEastLineBreakLanguageIDFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdLineBreakTraditionalChinese": WdFarEastLineBreakLanguageIDFromString = wdLineBreakTraditionalChinese
        Case "wdLineBreakJapanese": WdFarEastLineBreakLanguageIDFromString = wdLineBreakJapanese
        Case "wdLineBreakKorean": WdFarEastLineBreakLanguageIDFromString = wdLineBreakKorean
        Case "wdLineBreakSimplifiedChinese": WdFarEastLineBreakLanguageIDFromString = wdLineBreakSimplifiedChinese
    End Select
End Function

Function WdFarEastLineBreakLanguageIDToString(value As WdFarEastLineBreakLanguageID) As String
    Select Case value
        Case wdLineBreakTraditionalChinese: WdFarEastLineBreakLanguageIDToString = "wdLineBreakTraditionalChinese"
        Case wdLineBreakJapanese: WdFarEastLineBreakLanguageIDToString = "wdLineBreakJapanese"
        Case wdLineBreakKorean: WdFarEastLineBreakLanguageIDToString = "wdLineBreakKorean"
        Case wdLineBreakSimplifiedChinese: WdFarEastLineBreakLanguageIDToString = "wdLineBreakSimplifiedChinese"
    End Select
End Function
