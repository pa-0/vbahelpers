Attribute VB_Name = "wWdPageNumberStyle"
Function WdPageNumberStyleFromString(value As String) As WdPageNumberStyle
    If IsNumeric(value) Then
        WdPageNumberStyleFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdPageNumberStyleArabic": WdPageNumberStyleFromString = wdPageNumberStyleArabic
        Case "wdPageNumberStyleUppercaseRoman": WdPageNumberStyleFromString = wdPageNumberStyleUppercaseRoman
        Case "wdPageNumberStyleLowercaseRoman": WdPageNumberStyleFromString = wdPageNumberStyleLowercaseRoman
        Case "wdPageNumberStyleUppercaseLetter": WdPageNumberStyleFromString = wdPageNumberStyleUppercaseLetter
        Case "wdPageNumberStyleLowercaseLetter": WdPageNumberStyleFromString = wdPageNumberStyleLowercaseLetter
        Case "wdPageNumberStyleKanji": WdPageNumberStyleFromString = wdPageNumberStyleKanji
        Case "wdPageNumberStyleKanjiDigit": WdPageNumberStyleFromString = wdPageNumberStyleKanjiDigit
        Case "wdPageNumberStyleArabicFullWidth": WdPageNumberStyleFromString = wdPageNumberStyleArabicFullWidth
        Case "wdPageNumberStyleKanjiTraditional": WdPageNumberStyleFromString = wdPageNumberStyleKanjiTraditional
        Case "wdPageNumberStyleNumberInCircle": WdPageNumberStyleFromString = wdPageNumberStyleNumberInCircle
        Case "wdPageNumberStyleTradChinNum1": WdPageNumberStyleFromString = wdPageNumberStyleTradChinNum1
        Case "wdPageNumberStyleTradChinNum2": WdPageNumberStyleFromString = wdPageNumberStyleTradChinNum2
        Case "wdPageNumberStyleSimpChinNum1": WdPageNumberStyleFromString = wdPageNumberStyleSimpChinNum1
        Case "wdPageNumberStyleSimpChinNum2": WdPageNumberStyleFromString = wdPageNumberStyleSimpChinNum2
        Case "wdPageNumberStyleHanjaRead": WdPageNumberStyleFromString = wdPageNumberStyleHanjaRead
        Case "wdPageNumberStyleHanjaReadDigit": WdPageNumberStyleFromString = wdPageNumberStyleHanjaReadDigit
        Case "wdPageNumberStyleHebrewLetter1": WdPageNumberStyleFromString = wdPageNumberStyleHebrewLetter1
        Case "wdPageNumberStyleArabicLetter1": WdPageNumberStyleFromString = wdPageNumberStyleArabicLetter1
        Case "wdPageNumberStyleHebrewLetter2": WdPageNumberStyleFromString = wdPageNumberStyleHebrewLetter2
        Case "wdPageNumberStyleArabicLetter2": WdPageNumberStyleFromString = wdPageNumberStyleArabicLetter2
        Case "wdPageNumberStyleHindiLetter1": WdPageNumberStyleFromString = wdPageNumberStyleHindiLetter1
        Case "wdPageNumberStyleHindiLetter2": WdPageNumberStyleFromString = wdPageNumberStyleHindiLetter2
        Case "wdPageNumberStyleHindiArabic": WdPageNumberStyleFromString = wdPageNumberStyleHindiArabic
        Case "wdPageNumberStyleHindiCardinalText": WdPageNumberStyleFromString = wdPageNumberStyleHindiCardinalText
        Case "wdPageNumberStyleThaiLetter": WdPageNumberStyleFromString = wdPageNumberStyleThaiLetter
        Case "wdPageNumberStyleThaiArabic": WdPageNumberStyleFromString = wdPageNumberStyleThaiArabic
        Case "wdPageNumberStyleThaiCardinalText": WdPageNumberStyleFromString = wdPageNumberStyleThaiCardinalText
        Case "wdPageNumberStyleVietCardinalText": WdPageNumberStyleFromString = wdPageNumberStyleVietCardinalText
        Case "wdPageNumberStyleNumberInDash": WdPageNumberStyleFromString = wdPageNumberStyleNumberInDash
    End Select
End Function

Function WdPageNumberStyleToString(value As WdPageNumberStyle) As String
    Select Case value
        Case wdPageNumberStyleArabic: WdPageNumberStyleToString = "wdPageNumberStyleArabic"
        Case wdPageNumberStyleUppercaseRoman: WdPageNumberStyleToString = "wdPageNumberStyleUppercaseRoman"
        Case wdPageNumberStyleLowercaseRoman: WdPageNumberStyleToString = "wdPageNumberStyleLowercaseRoman"
        Case wdPageNumberStyleUppercaseLetter: WdPageNumberStyleToString = "wdPageNumberStyleUppercaseLetter"
        Case wdPageNumberStyleLowercaseLetter: WdPageNumberStyleToString = "wdPageNumberStyleLowercaseLetter"
        Case wdPageNumberStyleKanji: WdPageNumberStyleToString = "wdPageNumberStyleKanji"
        Case wdPageNumberStyleKanjiDigit: WdPageNumberStyleToString = "wdPageNumberStyleKanjiDigit"
        Case wdPageNumberStyleArabicFullWidth: WdPageNumberStyleToString = "wdPageNumberStyleArabicFullWidth"
        Case wdPageNumberStyleKanjiTraditional: WdPageNumberStyleToString = "wdPageNumberStyleKanjiTraditional"
        Case wdPageNumberStyleNumberInCircle: WdPageNumberStyleToString = "wdPageNumberStyleNumberInCircle"
        Case wdPageNumberStyleTradChinNum1: WdPageNumberStyleToString = "wdPageNumberStyleTradChinNum1"
        Case wdPageNumberStyleTradChinNum2: WdPageNumberStyleToString = "wdPageNumberStyleTradChinNum2"
        Case wdPageNumberStyleSimpChinNum1: WdPageNumberStyleToString = "wdPageNumberStyleSimpChinNum1"
        Case wdPageNumberStyleSimpChinNum2: WdPageNumberStyleToString = "wdPageNumberStyleSimpChinNum2"
        Case wdPageNumberStyleHanjaRead: WdPageNumberStyleToString = "wdPageNumberStyleHanjaRead"
        Case wdPageNumberStyleHanjaReadDigit: WdPageNumberStyleToString = "wdPageNumberStyleHanjaReadDigit"
        Case wdPageNumberStyleHebrewLetter1: WdPageNumberStyleToString = "wdPageNumberStyleHebrewLetter1"
        Case wdPageNumberStyleArabicLetter1: WdPageNumberStyleToString = "wdPageNumberStyleArabicLetter1"
        Case wdPageNumberStyleHebrewLetter2: WdPageNumberStyleToString = "wdPageNumberStyleHebrewLetter2"
        Case wdPageNumberStyleArabicLetter2: WdPageNumberStyleToString = "wdPageNumberStyleArabicLetter2"
        Case wdPageNumberStyleHindiLetter1: WdPageNumberStyleToString = "wdPageNumberStyleHindiLetter1"
        Case wdPageNumberStyleHindiLetter2: WdPageNumberStyleToString = "wdPageNumberStyleHindiLetter2"
        Case wdPageNumberStyleHindiArabic: WdPageNumberStyleToString = "wdPageNumberStyleHindiArabic"
        Case wdPageNumberStyleHindiCardinalText: WdPageNumberStyleToString = "wdPageNumberStyleHindiCardinalText"
        Case wdPageNumberStyleThaiLetter: WdPageNumberStyleToString = "wdPageNumberStyleThaiLetter"
        Case wdPageNumberStyleThaiArabic: WdPageNumberStyleToString = "wdPageNumberStyleThaiArabic"
        Case wdPageNumberStyleThaiCardinalText: WdPageNumberStyleToString = "wdPageNumberStyleThaiCardinalText"
        Case wdPageNumberStyleVietCardinalText: WdPageNumberStyleToString = "wdPageNumberStyleVietCardinalText"
        Case wdPageNumberStyleNumberInDash: WdPageNumberStyleToString = "wdPageNumberStyleNumberInDash"
    End Select
End Function
