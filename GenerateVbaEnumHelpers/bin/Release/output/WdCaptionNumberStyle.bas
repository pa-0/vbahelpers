Attribute VB_Name = "wWdCaptionNumberStyle"
Function WdCaptionNumberStyleFromString(value As String) As WdCaptionNumberStyle
    If IsNumeric(value) Then
        WdCaptionNumberStyleFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdCaptionNumberStyleArabic": WdCaptionNumberStyleFromString = wdCaptionNumberStyleArabic
        Case "wdCaptionNumberStyleUppercaseRoman": WdCaptionNumberStyleFromString = wdCaptionNumberStyleUppercaseRoman
        Case "wdCaptionNumberStyleLowercaseRoman": WdCaptionNumberStyleFromString = wdCaptionNumberStyleLowercaseRoman
        Case "wdCaptionNumberStyleUppercaseLetter": WdCaptionNumberStyleFromString = wdCaptionNumberStyleUppercaseLetter
        Case "wdCaptionNumberStyleLowercaseLetter": WdCaptionNumberStyleFromString = wdCaptionNumberStyleLowercaseLetter
        Case "wdCaptionNumberStyleKanji": WdCaptionNumberStyleFromString = wdCaptionNumberStyleKanji
        Case "wdCaptionNumberStyleKanjiDigit": WdCaptionNumberStyleFromString = wdCaptionNumberStyleKanjiDigit
        Case "wdCaptionNumberStyleArabicFullWidth": WdCaptionNumberStyleFromString = wdCaptionNumberStyleArabicFullWidth
        Case "wdCaptionNumberStyleKanjiTraditional": WdCaptionNumberStyleFromString = wdCaptionNumberStyleKanjiTraditional
        Case "wdCaptionNumberStyleNumberInCircle": WdCaptionNumberStyleFromString = wdCaptionNumberStyleNumberInCircle
        Case "wdCaptionNumberStyleGanada": WdCaptionNumberStyleFromString = wdCaptionNumberStyleGanada
        Case "wdCaptionNumberStyleChosung": WdCaptionNumberStyleFromString = wdCaptionNumberStyleChosung
        Case "wdCaptionNumberStyleZodiac1": WdCaptionNumberStyleFromString = wdCaptionNumberStyleZodiac1
        Case "wdCaptionNumberStyleZodiac2": WdCaptionNumberStyleFromString = wdCaptionNumberStyleZodiac2
        Case "wdCaptionNumberStyleTradChinNum2": WdCaptionNumberStyleFromString = wdCaptionNumberStyleTradChinNum2
        Case "wdCaptionNumberStyleTradChinNum3": WdCaptionNumberStyleFromString = wdCaptionNumberStyleTradChinNum3
        Case "wdCaptionNumberStyleSimpChinNum2": WdCaptionNumberStyleFromString = wdCaptionNumberStyleSimpChinNum2
        Case "wdCaptionNumberStyleSimpChinNum3": WdCaptionNumberStyleFromString = wdCaptionNumberStyleSimpChinNum3
        Case "wdCaptionNumberStyleHanjaRead": WdCaptionNumberStyleFromString = wdCaptionNumberStyleHanjaRead
        Case "wdCaptionNumberStyleHanjaReadDigit": WdCaptionNumberStyleFromString = wdCaptionNumberStyleHanjaReadDigit
        Case "wdCaptionNumberStyleHebrewLetter1": WdCaptionNumberStyleFromString = wdCaptionNumberStyleHebrewLetter1
        Case "wdCaptionNumberStyleArabicLetter1": WdCaptionNumberStyleFromString = wdCaptionNumberStyleArabicLetter1
        Case "wdCaptionNumberStyleHebrewLetter2": WdCaptionNumberStyleFromString = wdCaptionNumberStyleHebrewLetter2
        Case "wdCaptionNumberStyleArabicLetter2": WdCaptionNumberStyleFromString = wdCaptionNumberStyleArabicLetter2
        Case "wdCaptionNumberStyleHindiLetter1": WdCaptionNumberStyleFromString = wdCaptionNumberStyleHindiLetter1
        Case "wdCaptionNumberStyleHindiLetter2": WdCaptionNumberStyleFromString = wdCaptionNumberStyleHindiLetter2
        Case "wdCaptionNumberStyleHindiArabic": WdCaptionNumberStyleFromString = wdCaptionNumberStyleHindiArabic
        Case "wdCaptionNumberStyleHindiCardinalText": WdCaptionNumberStyleFromString = wdCaptionNumberStyleHindiCardinalText
        Case "wdCaptionNumberStyleThaiLetter": WdCaptionNumberStyleFromString = wdCaptionNumberStyleThaiLetter
        Case "wdCaptionNumberStyleThaiArabic": WdCaptionNumberStyleFromString = wdCaptionNumberStyleThaiArabic
        Case "wdCaptionNumberStyleThaiCardinalText": WdCaptionNumberStyleFromString = wdCaptionNumberStyleThaiCardinalText
        Case "wdCaptionNumberStyleVietCardinalText": WdCaptionNumberStyleFromString = wdCaptionNumberStyleVietCardinalText
    End Select
End Function

Function WdCaptionNumberStyleToString(value As WdCaptionNumberStyle) As String
    Select Case value
        Case wdCaptionNumberStyleArabic: WdCaptionNumberStyleToString = "wdCaptionNumberStyleArabic"
        Case wdCaptionNumberStyleUppercaseRoman: WdCaptionNumberStyleToString = "wdCaptionNumberStyleUppercaseRoman"
        Case wdCaptionNumberStyleLowercaseRoman: WdCaptionNumberStyleToString = "wdCaptionNumberStyleLowercaseRoman"
        Case wdCaptionNumberStyleUppercaseLetter: WdCaptionNumberStyleToString = "wdCaptionNumberStyleUppercaseLetter"
        Case wdCaptionNumberStyleLowercaseLetter: WdCaptionNumberStyleToString = "wdCaptionNumberStyleLowercaseLetter"
        Case wdCaptionNumberStyleKanji: WdCaptionNumberStyleToString = "wdCaptionNumberStyleKanji"
        Case wdCaptionNumberStyleKanjiDigit: WdCaptionNumberStyleToString = "wdCaptionNumberStyleKanjiDigit"
        Case wdCaptionNumberStyleArabicFullWidth: WdCaptionNumberStyleToString = "wdCaptionNumberStyleArabicFullWidth"
        Case wdCaptionNumberStyleKanjiTraditional: WdCaptionNumberStyleToString = "wdCaptionNumberStyleKanjiTraditional"
        Case wdCaptionNumberStyleNumberInCircle: WdCaptionNumberStyleToString = "wdCaptionNumberStyleNumberInCircle"
        Case wdCaptionNumberStyleGanada: WdCaptionNumberStyleToString = "wdCaptionNumberStyleGanada"
        Case wdCaptionNumberStyleChosung: WdCaptionNumberStyleToString = "wdCaptionNumberStyleChosung"
        Case wdCaptionNumberStyleZodiac1: WdCaptionNumberStyleToString = "wdCaptionNumberStyleZodiac1"
        Case wdCaptionNumberStyleZodiac2: WdCaptionNumberStyleToString = "wdCaptionNumberStyleZodiac2"
        Case wdCaptionNumberStyleTradChinNum2: WdCaptionNumberStyleToString = "wdCaptionNumberStyleTradChinNum2"
        Case wdCaptionNumberStyleTradChinNum3: WdCaptionNumberStyleToString = "wdCaptionNumberStyleTradChinNum3"
        Case wdCaptionNumberStyleSimpChinNum2: WdCaptionNumberStyleToString = "wdCaptionNumberStyleSimpChinNum2"
        Case wdCaptionNumberStyleSimpChinNum3: WdCaptionNumberStyleToString = "wdCaptionNumberStyleSimpChinNum3"
        Case wdCaptionNumberStyleHanjaRead: WdCaptionNumberStyleToString = "wdCaptionNumberStyleHanjaRead"
        Case wdCaptionNumberStyleHanjaReadDigit: WdCaptionNumberStyleToString = "wdCaptionNumberStyleHanjaReadDigit"
        Case wdCaptionNumberStyleHebrewLetter1: WdCaptionNumberStyleToString = "wdCaptionNumberStyleHebrewLetter1"
        Case wdCaptionNumberStyleArabicLetter1: WdCaptionNumberStyleToString = "wdCaptionNumberStyleArabicLetter1"
        Case wdCaptionNumberStyleHebrewLetter2: WdCaptionNumberStyleToString = "wdCaptionNumberStyleHebrewLetter2"
        Case wdCaptionNumberStyleArabicLetter2: WdCaptionNumberStyleToString = "wdCaptionNumberStyleArabicLetter2"
        Case wdCaptionNumberStyleHindiLetter1: WdCaptionNumberStyleToString = "wdCaptionNumberStyleHindiLetter1"
        Case wdCaptionNumberStyleHindiLetter2: WdCaptionNumberStyleToString = "wdCaptionNumberStyleHindiLetter2"
        Case wdCaptionNumberStyleHindiArabic: WdCaptionNumberStyleToString = "wdCaptionNumberStyleHindiArabic"
        Case wdCaptionNumberStyleHindiCardinalText: WdCaptionNumberStyleToString = "wdCaptionNumberStyleHindiCardinalText"
        Case wdCaptionNumberStyleThaiLetter: WdCaptionNumberStyleToString = "wdCaptionNumberStyleThaiLetter"
        Case wdCaptionNumberStyleThaiArabic: WdCaptionNumberStyleToString = "wdCaptionNumberStyleThaiArabic"
        Case wdCaptionNumberStyleThaiCardinalText: WdCaptionNumberStyleToString = "wdCaptionNumberStyleThaiCardinalText"
        Case wdCaptionNumberStyleVietCardinalText: WdCaptionNumberStyleToString = "wdCaptionNumberStyleVietCardinalText"
    End Select
End Function
