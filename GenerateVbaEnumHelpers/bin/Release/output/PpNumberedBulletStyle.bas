Attribute VB_Name = "wPpNumberedBulletStyle"
Function PpNumberedBulletStyleFromString(value As String) As PpNumberedBulletStyle
    If IsNumeric(value) Then
        PpNumberedBulletStyleFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "ppBulletAlphaLCPeriod": PpNumberedBulletStyleFromString = ppBulletAlphaLCPeriod
        Case "ppBulletAlphaUCPeriod": PpNumberedBulletStyleFromString = ppBulletAlphaUCPeriod
        Case "ppBulletArabicParenRight": PpNumberedBulletStyleFromString = ppBulletArabicParenRight
        Case "ppBulletArabicPeriod": PpNumberedBulletStyleFromString = ppBulletArabicPeriod
        Case "ppBulletRomanLCParenBoth": PpNumberedBulletStyleFromString = ppBulletRomanLCParenBoth
        Case "ppBulletRomanLCParenRight": PpNumberedBulletStyleFromString = ppBulletRomanLCParenRight
        Case "ppBulletRomanLCPeriod": PpNumberedBulletStyleFromString = ppBulletRomanLCPeriod
        Case "ppBulletRomanUCPeriod": PpNumberedBulletStyleFromString = ppBulletRomanUCPeriod
        Case "ppBulletAlphaLCParenBoth": PpNumberedBulletStyleFromString = ppBulletAlphaLCParenBoth
        Case "ppBulletAlphaLCParenRight": PpNumberedBulletStyleFromString = ppBulletAlphaLCParenRight
        Case "ppBulletAlphaUCParenBoth": PpNumberedBulletStyleFromString = ppBulletAlphaUCParenBoth
        Case "ppBulletAlphaUCParenRight": PpNumberedBulletStyleFromString = ppBulletAlphaUCParenRight
        Case "ppBulletArabicParenBoth": PpNumberedBulletStyleFromString = ppBulletArabicParenBoth
        Case "ppBulletArabicPlain": PpNumberedBulletStyleFromString = ppBulletArabicPlain
        Case "ppBulletRomanUCParenBoth": PpNumberedBulletStyleFromString = ppBulletRomanUCParenBoth
        Case "ppBulletRomanUCParenRight": PpNumberedBulletStyleFromString = ppBulletRomanUCParenRight
        Case "ppBulletSimpChinPlain": PpNumberedBulletStyleFromString = ppBulletSimpChinPlain
        Case "ppBulletSimpChinPeriod": PpNumberedBulletStyleFromString = ppBulletSimpChinPeriod
        Case "ppBulletCircleNumDBPlain": PpNumberedBulletStyleFromString = ppBulletCircleNumDBPlain
        Case "ppBulletCircleNumWDWhitePlain": PpNumberedBulletStyleFromString = ppBulletCircleNumWDWhitePlain
        Case "ppBulletCircleNumWDBlackPlain": PpNumberedBulletStyleFromString = ppBulletCircleNumWDBlackPlain
        Case "ppBulletTradChinPlain": PpNumberedBulletStyleFromString = ppBulletTradChinPlain
        Case "ppBulletTradChinPeriod": PpNumberedBulletStyleFromString = ppBulletTradChinPeriod
        Case "ppBulletArabicAlphaDash": PpNumberedBulletStyleFromString = ppBulletArabicAlphaDash
        Case "ppBulletArabicAbjadDash": PpNumberedBulletStyleFromString = ppBulletArabicAbjadDash
        Case "ppBulletHebrewAlphaDash": PpNumberedBulletStyleFromString = ppBulletHebrewAlphaDash
        Case "ppBulletKanjiKoreanPlain": PpNumberedBulletStyleFromString = ppBulletKanjiKoreanPlain
        Case "ppBulletKanjiKoreanPeriod": PpNumberedBulletStyleFromString = ppBulletKanjiKoreanPeriod
        Case "ppBulletArabicDBPlain": PpNumberedBulletStyleFromString = ppBulletArabicDBPlain
        Case "ppBulletArabicDBPeriod": PpNumberedBulletStyleFromString = ppBulletArabicDBPeriod
        Case "ppBulletThaiAlphaPeriod": PpNumberedBulletStyleFromString = ppBulletThaiAlphaPeriod
        Case "ppBulletThaiAlphaParenRight": PpNumberedBulletStyleFromString = ppBulletThaiAlphaParenRight
        Case "ppBulletThaiAlphaParenBoth": PpNumberedBulletStyleFromString = ppBulletThaiAlphaParenBoth
        Case "ppBulletThaiNumPeriod": PpNumberedBulletStyleFromString = ppBulletThaiNumPeriod
        Case "ppBulletThaiNumParenRight": PpNumberedBulletStyleFromString = ppBulletThaiNumParenRight
        Case "ppBulletThaiNumParenBoth": PpNumberedBulletStyleFromString = ppBulletThaiNumParenBoth
        Case "ppBulletHindiAlphaPeriod": PpNumberedBulletStyleFromString = ppBulletHindiAlphaPeriod
        Case "ppBulletHindiNumPeriod": PpNumberedBulletStyleFromString = ppBulletHindiNumPeriod
        Case "ppBulletKanjiSimpChinDBPeriod": PpNumberedBulletStyleFromString = ppBulletKanjiSimpChinDBPeriod
        Case "ppBulletHindiNumParenRight": PpNumberedBulletStyleFromString = ppBulletHindiNumParenRight
        Case "ppBulletHindiAlpha1Period": PpNumberedBulletStyleFromString = ppBulletHindiAlpha1Period
        Case "ppBulletStyleMixed": PpNumberedBulletStyleFromString = ppBulletStyleMixed
    End Select
End Function

Function PpNumberedBulletStyleToString(value As PpNumberedBulletStyle) As String
    Select Case value
        Case ppBulletAlphaLCPeriod: PpNumberedBulletStyleToString = "ppBulletAlphaLCPeriod"
        Case ppBulletAlphaUCPeriod: PpNumberedBulletStyleToString = "ppBulletAlphaUCPeriod"
        Case ppBulletArabicParenRight: PpNumberedBulletStyleToString = "ppBulletArabicParenRight"
        Case ppBulletArabicPeriod: PpNumberedBulletStyleToString = "ppBulletArabicPeriod"
        Case ppBulletRomanLCParenBoth: PpNumberedBulletStyleToString = "ppBulletRomanLCParenBoth"
        Case ppBulletRomanLCParenRight: PpNumberedBulletStyleToString = "ppBulletRomanLCParenRight"
        Case ppBulletRomanLCPeriod: PpNumberedBulletStyleToString = "ppBulletRomanLCPeriod"
        Case ppBulletRomanUCPeriod: PpNumberedBulletStyleToString = "ppBulletRomanUCPeriod"
        Case ppBulletAlphaLCParenBoth: PpNumberedBulletStyleToString = "ppBulletAlphaLCParenBoth"
        Case ppBulletAlphaLCParenRight: PpNumberedBulletStyleToString = "ppBulletAlphaLCParenRight"
        Case ppBulletAlphaUCParenBoth: PpNumberedBulletStyleToString = "ppBulletAlphaUCParenBoth"
        Case ppBulletAlphaUCParenRight: PpNumberedBulletStyleToString = "ppBulletAlphaUCParenRight"
        Case ppBulletArabicParenBoth: PpNumberedBulletStyleToString = "ppBulletArabicParenBoth"
        Case ppBulletArabicPlain: PpNumberedBulletStyleToString = "ppBulletArabicPlain"
        Case ppBulletRomanUCParenBoth: PpNumberedBulletStyleToString = "ppBulletRomanUCParenBoth"
        Case ppBulletRomanUCParenRight: PpNumberedBulletStyleToString = "ppBulletRomanUCParenRight"
        Case ppBulletSimpChinPlain: PpNumberedBulletStyleToString = "ppBulletSimpChinPlain"
        Case ppBulletSimpChinPeriod: PpNumberedBulletStyleToString = "ppBulletSimpChinPeriod"
        Case ppBulletCircleNumDBPlain: PpNumberedBulletStyleToString = "ppBulletCircleNumDBPlain"
        Case ppBulletCircleNumWDWhitePlain: PpNumberedBulletStyleToString = "ppBulletCircleNumWDWhitePlain"
        Case ppBulletCircleNumWDBlackPlain: PpNumberedBulletStyleToString = "ppBulletCircleNumWDBlackPlain"
        Case ppBulletTradChinPlain: PpNumberedBulletStyleToString = "ppBulletTradChinPlain"
        Case ppBulletTradChinPeriod: PpNumberedBulletStyleToString = "ppBulletTradChinPeriod"
        Case ppBulletArabicAlphaDash: PpNumberedBulletStyleToString = "ppBulletArabicAlphaDash"
        Case ppBulletArabicAbjadDash: PpNumberedBulletStyleToString = "ppBulletArabicAbjadDash"
        Case ppBulletHebrewAlphaDash: PpNumberedBulletStyleToString = "ppBulletHebrewAlphaDash"
        Case ppBulletKanjiKoreanPlain: PpNumberedBulletStyleToString = "ppBulletKanjiKoreanPlain"
        Case ppBulletKanjiKoreanPeriod: PpNumberedBulletStyleToString = "ppBulletKanjiKoreanPeriod"
        Case ppBulletArabicDBPlain: PpNumberedBulletStyleToString = "ppBulletArabicDBPlain"
        Case ppBulletArabicDBPeriod: PpNumberedBulletStyleToString = "ppBulletArabicDBPeriod"
        Case ppBulletThaiAlphaPeriod: PpNumberedBulletStyleToString = "ppBulletThaiAlphaPeriod"
        Case ppBulletThaiAlphaParenRight: PpNumberedBulletStyleToString = "ppBulletThaiAlphaParenRight"
        Case ppBulletThaiAlphaParenBoth: PpNumberedBulletStyleToString = "ppBulletThaiAlphaParenBoth"
        Case ppBulletThaiNumPeriod: PpNumberedBulletStyleToString = "ppBulletThaiNumPeriod"
        Case ppBulletThaiNumParenRight: PpNumberedBulletStyleToString = "ppBulletThaiNumParenRight"
        Case ppBulletThaiNumParenBoth: PpNumberedBulletStyleToString = "ppBulletThaiNumParenBoth"
        Case ppBulletHindiAlphaPeriod: PpNumberedBulletStyleToString = "ppBulletHindiAlphaPeriod"
        Case ppBulletHindiNumPeriod: PpNumberedBulletStyleToString = "ppBulletHindiNumPeriod"
        Case ppBulletKanjiSimpChinDBPeriod: PpNumberedBulletStyleToString = "ppBulletKanjiSimpChinDBPeriod"
        Case ppBulletHindiNumParenRight: PpNumberedBulletStyleToString = "ppBulletHindiNumParenRight"
        Case ppBulletHindiAlpha1Period: PpNumberedBulletStyleToString = "ppBulletHindiAlpha1Period"
        Case ppBulletStyleMixed: PpNumberedBulletStyleToString = "ppBulletStyleMixed"
    End Select
End Function
