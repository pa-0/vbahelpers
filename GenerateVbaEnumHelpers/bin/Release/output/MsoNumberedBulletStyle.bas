Attribute VB_Name = "wMsoNumberedBulletStyle"
Function MsoNumberedBulletStyleFromString(value As String) As MsoNumberedBulletStyle
    If IsNumeric(value) Then
        MsoNumberedBulletStyleFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoBulletAlphaLCPeriod": MsoNumberedBulletStyleFromString = msoBulletAlphaLCPeriod
        Case "msoBulletAlphaUCPeriod": MsoNumberedBulletStyleFromString = msoBulletAlphaUCPeriod
        Case "msoBulletArabicParenRight": MsoNumberedBulletStyleFromString = msoBulletArabicParenRight
        Case "msoBulletArabicPeriod": MsoNumberedBulletStyleFromString = msoBulletArabicPeriod
        Case "msoBulletRomanLCParenBoth": MsoNumberedBulletStyleFromString = msoBulletRomanLCParenBoth
        Case "msoBulletRomanLCParenRight": MsoNumberedBulletStyleFromString = msoBulletRomanLCParenRight
        Case "msoBulletRomanLCPeriod": MsoNumberedBulletStyleFromString = msoBulletRomanLCPeriod
        Case "msoBulletRomanUCPeriod": MsoNumberedBulletStyleFromString = msoBulletRomanUCPeriod
        Case "msoBulletAlphaLCParenBoth": MsoNumberedBulletStyleFromString = msoBulletAlphaLCParenBoth
        Case "msoBulletAlphaLCParenRight": MsoNumberedBulletStyleFromString = msoBulletAlphaLCParenRight
        Case "msoBulletAlphaUCParenBoth": MsoNumberedBulletStyleFromString = msoBulletAlphaUCParenBoth
        Case "msoBulletAlphaUCParenRight": MsoNumberedBulletStyleFromString = msoBulletAlphaUCParenRight
        Case "msoBulletArabicParenBoth": MsoNumberedBulletStyleFromString = msoBulletArabicParenBoth
        Case "msoBulletArabicPlain": MsoNumberedBulletStyleFromString = msoBulletArabicPlain
        Case "msoBulletRomanUCParenBoth": MsoNumberedBulletStyleFromString = msoBulletRomanUCParenBoth
        Case "msoBulletRomanUCParenRight": MsoNumberedBulletStyleFromString = msoBulletRomanUCParenRight
        Case "msoBulletSimpChinPlain": MsoNumberedBulletStyleFromString = msoBulletSimpChinPlain
        Case "msoBulletSimpChinPeriod": MsoNumberedBulletStyleFromString = msoBulletSimpChinPeriod
        Case "msoBulletCircleNumDBPlain": MsoNumberedBulletStyleFromString = msoBulletCircleNumDBPlain
        Case "msoBulletCircleNumWDWhitePlain": MsoNumberedBulletStyleFromString = msoBulletCircleNumWDWhitePlain
        Case "msoBulletCircleNumWDBlackPlain": MsoNumberedBulletStyleFromString = msoBulletCircleNumWDBlackPlain
        Case "msoBulletTradChinPlain": MsoNumberedBulletStyleFromString = msoBulletTradChinPlain
        Case "msoBulletTradChinPeriod": MsoNumberedBulletStyleFromString = msoBulletTradChinPeriod
        Case "msoBulletArabicAlphaDash": MsoNumberedBulletStyleFromString = msoBulletArabicAlphaDash
        Case "msoBulletArabicAbjadDash": MsoNumberedBulletStyleFromString = msoBulletArabicAbjadDash
        Case "msoBulletHebrewAlphaDash": MsoNumberedBulletStyleFromString = msoBulletHebrewAlphaDash
        Case "msoBulletKanjiKoreanPlain": MsoNumberedBulletStyleFromString = msoBulletKanjiKoreanPlain
        Case "msoBulletKanjiKoreanPeriod": MsoNumberedBulletStyleFromString = msoBulletKanjiKoreanPeriod
        Case "msoBulletArabicDBPlain": MsoNumberedBulletStyleFromString = msoBulletArabicDBPlain
        Case "msoBulletArabicDBPeriod": MsoNumberedBulletStyleFromString = msoBulletArabicDBPeriod
        Case "msoBulletThaiAlphaPeriod": MsoNumberedBulletStyleFromString = msoBulletThaiAlphaPeriod
        Case "msoBulletThaiAlphaParenRight": MsoNumberedBulletStyleFromString = msoBulletThaiAlphaParenRight
        Case "msoBulletThaiAlphaParenBoth": MsoNumberedBulletStyleFromString = msoBulletThaiAlphaParenBoth
        Case "msoBulletThaiNumPeriod": MsoNumberedBulletStyleFromString = msoBulletThaiNumPeriod
        Case "msoBulletThaiNumParenRight": MsoNumberedBulletStyleFromString = msoBulletThaiNumParenRight
        Case "msoBulletThaiNumParenBoth": MsoNumberedBulletStyleFromString = msoBulletThaiNumParenBoth
        Case "msoBulletHindiAlphaPeriod": MsoNumberedBulletStyleFromString = msoBulletHindiAlphaPeriod
        Case "msoBulletHindiNumPeriod": MsoNumberedBulletStyleFromString = msoBulletHindiNumPeriod
        Case "msoBulletKanjiSimpChinDBPeriod": MsoNumberedBulletStyleFromString = msoBulletKanjiSimpChinDBPeriod
        Case "msoBulletHindiNumParenRight": MsoNumberedBulletStyleFromString = msoBulletHindiNumParenRight
        Case "msoBulletHindiAlpha1Period": MsoNumberedBulletStyleFromString = msoBulletHindiAlpha1Period
        Case "msoBulletStyleMixed": MsoNumberedBulletStyleFromString = msoBulletStyleMixed
    End Select
End Function

Function MsoNumberedBulletStyleToString(value As MsoNumberedBulletStyle) As String
    Select Case value
        Case msoBulletAlphaLCPeriod: MsoNumberedBulletStyleToString = "msoBulletAlphaLCPeriod"
        Case msoBulletAlphaUCPeriod: MsoNumberedBulletStyleToString = "msoBulletAlphaUCPeriod"
        Case msoBulletArabicParenRight: MsoNumberedBulletStyleToString = "msoBulletArabicParenRight"
        Case msoBulletArabicPeriod: MsoNumberedBulletStyleToString = "msoBulletArabicPeriod"
        Case msoBulletRomanLCParenBoth: MsoNumberedBulletStyleToString = "msoBulletRomanLCParenBoth"
        Case msoBulletRomanLCParenRight: MsoNumberedBulletStyleToString = "msoBulletRomanLCParenRight"
        Case msoBulletRomanLCPeriod: MsoNumberedBulletStyleToString = "msoBulletRomanLCPeriod"
        Case msoBulletRomanUCPeriod: MsoNumberedBulletStyleToString = "msoBulletRomanUCPeriod"
        Case msoBulletAlphaLCParenBoth: MsoNumberedBulletStyleToString = "msoBulletAlphaLCParenBoth"
        Case msoBulletAlphaLCParenRight: MsoNumberedBulletStyleToString = "msoBulletAlphaLCParenRight"
        Case msoBulletAlphaUCParenBoth: MsoNumberedBulletStyleToString = "msoBulletAlphaUCParenBoth"
        Case msoBulletAlphaUCParenRight: MsoNumberedBulletStyleToString = "msoBulletAlphaUCParenRight"
        Case msoBulletArabicParenBoth: MsoNumberedBulletStyleToString = "msoBulletArabicParenBoth"
        Case msoBulletArabicPlain: MsoNumberedBulletStyleToString = "msoBulletArabicPlain"
        Case msoBulletRomanUCParenBoth: MsoNumberedBulletStyleToString = "msoBulletRomanUCParenBoth"
        Case msoBulletRomanUCParenRight: MsoNumberedBulletStyleToString = "msoBulletRomanUCParenRight"
        Case msoBulletSimpChinPlain: MsoNumberedBulletStyleToString = "msoBulletSimpChinPlain"
        Case msoBulletSimpChinPeriod: MsoNumberedBulletStyleToString = "msoBulletSimpChinPeriod"
        Case msoBulletCircleNumDBPlain: MsoNumberedBulletStyleToString = "msoBulletCircleNumDBPlain"
        Case msoBulletCircleNumWDWhitePlain: MsoNumberedBulletStyleToString = "msoBulletCircleNumWDWhitePlain"
        Case msoBulletCircleNumWDBlackPlain: MsoNumberedBulletStyleToString = "msoBulletCircleNumWDBlackPlain"
        Case msoBulletTradChinPlain: MsoNumberedBulletStyleToString = "msoBulletTradChinPlain"
        Case msoBulletTradChinPeriod: MsoNumberedBulletStyleToString = "msoBulletTradChinPeriod"
        Case msoBulletArabicAlphaDash: MsoNumberedBulletStyleToString = "msoBulletArabicAlphaDash"
        Case msoBulletArabicAbjadDash: MsoNumberedBulletStyleToString = "msoBulletArabicAbjadDash"
        Case msoBulletHebrewAlphaDash: MsoNumberedBulletStyleToString = "msoBulletHebrewAlphaDash"
        Case msoBulletKanjiKoreanPlain: MsoNumberedBulletStyleToString = "msoBulletKanjiKoreanPlain"
        Case msoBulletKanjiKoreanPeriod: MsoNumberedBulletStyleToString = "msoBulletKanjiKoreanPeriod"
        Case msoBulletArabicDBPlain: MsoNumberedBulletStyleToString = "msoBulletArabicDBPlain"
        Case msoBulletArabicDBPeriod: MsoNumberedBulletStyleToString = "msoBulletArabicDBPeriod"
        Case msoBulletThaiAlphaPeriod: MsoNumberedBulletStyleToString = "msoBulletThaiAlphaPeriod"
        Case msoBulletThaiAlphaParenRight: MsoNumberedBulletStyleToString = "msoBulletThaiAlphaParenRight"
        Case msoBulletThaiAlphaParenBoth: MsoNumberedBulletStyleToString = "msoBulletThaiAlphaParenBoth"
        Case msoBulletThaiNumPeriod: MsoNumberedBulletStyleToString = "msoBulletThaiNumPeriod"
        Case msoBulletThaiNumParenRight: MsoNumberedBulletStyleToString = "msoBulletThaiNumParenRight"
        Case msoBulletThaiNumParenBoth: MsoNumberedBulletStyleToString = "msoBulletThaiNumParenBoth"
        Case msoBulletHindiAlphaPeriod: MsoNumberedBulletStyleToString = "msoBulletHindiAlphaPeriod"
        Case msoBulletHindiNumPeriod: MsoNumberedBulletStyleToString = "msoBulletHindiNumPeriod"
        Case msoBulletKanjiSimpChinDBPeriod: MsoNumberedBulletStyleToString = "msoBulletKanjiSimpChinDBPeriod"
        Case msoBulletHindiNumParenRight: MsoNumberedBulletStyleToString = "msoBulletHindiNumParenRight"
        Case msoBulletHindiAlpha1Period: MsoNumberedBulletStyleToString = "msoBulletHindiAlpha1Period"
        Case msoBulletStyleMixed: MsoNumberedBulletStyleToString = "msoBulletStyleMixed"
    End Select
End Function
