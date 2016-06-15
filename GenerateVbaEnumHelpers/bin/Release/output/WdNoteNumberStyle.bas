Attribute VB_Name = "wWdNoteNumberStyle"
Function WdNoteNumberStyleFromString(value As String) As WdNoteNumberStyle
    If IsNumeric(value) Then
        WdNoteNumberStyleFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdNoteNumberStyleArabic": WdNoteNumberStyleFromString = wdNoteNumberStyleArabic
        Case "wdNoteNumberStyleUppercaseRoman": WdNoteNumberStyleFromString = wdNoteNumberStyleUppercaseRoman
        Case "wdNoteNumberStyleLowercaseRoman": WdNoteNumberStyleFromString = wdNoteNumberStyleLowercaseRoman
        Case "wdNoteNumberStyleUppercaseLetter": WdNoteNumberStyleFromString = wdNoteNumberStyleUppercaseLetter
        Case "wdNoteNumberStyleLowercaseLetter": WdNoteNumberStyleFromString = wdNoteNumberStyleLowercaseLetter
        Case "wdNoteNumberStyleSymbol": WdNoteNumberStyleFromString = wdNoteNumberStyleSymbol
        Case "wdNoteNumberStyleKanji": WdNoteNumberStyleFromString = wdNoteNumberStyleKanji
        Case "wdNoteNumberStyleKanjiDigit": WdNoteNumberStyleFromString = wdNoteNumberStyleKanjiDigit
        Case "wdNoteNumberStyleArabicFullWidth": WdNoteNumberStyleFromString = wdNoteNumberStyleArabicFullWidth
        Case "wdNoteNumberStyleKanjiTraditional": WdNoteNumberStyleFromString = wdNoteNumberStyleKanjiTraditional
        Case "wdNoteNumberStyleNumberInCircle": WdNoteNumberStyleFromString = wdNoteNumberStyleNumberInCircle
        Case "wdNoteNumberStyleTradChinNum1": WdNoteNumberStyleFromString = wdNoteNumberStyleTradChinNum1
        Case "wdNoteNumberStyleTradChinNum2": WdNoteNumberStyleFromString = wdNoteNumberStyleTradChinNum2
        Case "wdNoteNumberStyleSimpChinNum1": WdNoteNumberStyleFromString = wdNoteNumberStyleSimpChinNum1
        Case "wdNoteNumberStyleSimpChinNum2": WdNoteNumberStyleFromString = wdNoteNumberStyleSimpChinNum2
        Case "wdNoteNumberStyleHanjaRead": WdNoteNumberStyleFromString = wdNoteNumberStyleHanjaRead
        Case "wdNoteNumberStyleHanjaReadDigit": WdNoteNumberStyleFromString = wdNoteNumberStyleHanjaReadDigit
        Case "wdNoteNumberStyleHebrewLetter1": WdNoteNumberStyleFromString = wdNoteNumberStyleHebrewLetter1
        Case "wdNoteNumberStyleArabicLetter1": WdNoteNumberStyleFromString = wdNoteNumberStyleArabicLetter1
        Case "wdNoteNumberStyleHebrewLetter2": WdNoteNumberStyleFromString = wdNoteNumberStyleHebrewLetter2
        Case "wdNoteNumberStyleArabicLetter2": WdNoteNumberStyleFromString = wdNoteNumberStyleArabicLetter2
        Case "wdNoteNumberStyleHindiLetter1": WdNoteNumberStyleFromString = wdNoteNumberStyleHindiLetter1
        Case "wdNoteNumberStyleHindiLetter2": WdNoteNumberStyleFromString = wdNoteNumberStyleHindiLetter2
        Case "wdNoteNumberStyleHindiArabic": WdNoteNumberStyleFromString = wdNoteNumberStyleHindiArabic
        Case "wdNoteNumberStyleHindiCardinalText": WdNoteNumberStyleFromString = wdNoteNumberStyleHindiCardinalText
        Case "wdNoteNumberStyleThaiLetter": WdNoteNumberStyleFromString = wdNoteNumberStyleThaiLetter
        Case "wdNoteNumberStyleThaiArabic": WdNoteNumberStyleFromString = wdNoteNumberStyleThaiArabic
        Case "wdNoteNumberStyleThaiCardinalText": WdNoteNumberStyleFromString = wdNoteNumberStyleThaiCardinalText
        Case "wdNoteNumberStyleVietCardinalText": WdNoteNumberStyleFromString = wdNoteNumberStyleVietCardinalText
    End Select
End Function

Function WdNoteNumberStyleToString(value As WdNoteNumberStyle) As String
    Select Case value
        Case wdNoteNumberStyleArabic: WdNoteNumberStyleToString = "wdNoteNumberStyleArabic"
        Case wdNoteNumberStyleUppercaseRoman: WdNoteNumberStyleToString = "wdNoteNumberStyleUppercaseRoman"
        Case wdNoteNumberStyleLowercaseRoman: WdNoteNumberStyleToString = "wdNoteNumberStyleLowercaseRoman"
        Case wdNoteNumberStyleUppercaseLetter: WdNoteNumberStyleToString = "wdNoteNumberStyleUppercaseLetter"
        Case wdNoteNumberStyleLowercaseLetter: WdNoteNumberStyleToString = "wdNoteNumberStyleLowercaseLetter"
        Case wdNoteNumberStyleSymbol: WdNoteNumberStyleToString = "wdNoteNumberStyleSymbol"
        Case wdNoteNumberStyleKanji: WdNoteNumberStyleToString = "wdNoteNumberStyleKanji"
        Case wdNoteNumberStyleKanjiDigit: WdNoteNumberStyleToString = "wdNoteNumberStyleKanjiDigit"
        Case wdNoteNumberStyleArabicFullWidth: WdNoteNumberStyleToString = "wdNoteNumberStyleArabicFullWidth"
        Case wdNoteNumberStyleKanjiTraditional: WdNoteNumberStyleToString = "wdNoteNumberStyleKanjiTraditional"
        Case wdNoteNumberStyleNumberInCircle: WdNoteNumberStyleToString = "wdNoteNumberStyleNumberInCircle"
        Case wdNoteNumberStyleTradChinNum1: WdNoteNumberStyleToString = "wdNoteNumberStyleTradChinNum1"
        Case wdNoteNumberStyleTradChinNum2: WdNoteNumberStyleToString = "wdNoteNumberStyleTradChinNum2"
        Case wdNoteNumberStyleSimpChinNum1: WdNoteNumberStyleToString = "wdNoteNumberStyleSimpChinNum1"
        Case wdNoteNumberStyleSimpChinNum2: WdNoteNumberStyleToString = "wdNoteNumberStyleSimpChinNum2"
        Case wdNoteNumberStyleHanjaRead: WdNoteNumberStyleToString = "wdNoteNumberStyleHanjaRead"
        Case wdNoteNumberStyleHanjaReadDigit: WdNoteNumberStyleToString = "wdNoteNumberStyleHanjaReadDigit"
        Case wdNoteNumberStyleHebrewLetter1: WdNoteNumberStyleToString = "wdNoteNumberStyleHebrewLetter1"
        Case wdNoteNumberStyleArabicLetter1: WdNoteNumberStyleToString = "wdNoteNumberStyleArabicLetter1"
        Case wdNoteNumberStyleHebrewLetter2: WdNoteNumberStyleToString = "wdNoteNumberStyleHebrewLetter2"
        Case wdNoteNumberStyleArabicLetter2: WdNoteNumberStyleToString = "wdNoteNumberStyleArabicLetter2"
        Case wdNoteNumberStyleHindiLetter1: WdNoteNumberStyleToString = "wdNoteNumberStyleHindiLetter1"
        Case wdNoteNumberStyleHindiLetter2: WdNoteNumberStyleToString = "wdNoteNumberStyleHindiLetter2"
        Case wdNoteNumberStyleHindiArabic: WdNoteNumberStyleToString = "wdNoteNumberStyleHindiArabic"
        Case wdNoteNumberStyleHindiCardinalText: WdNoteNumberStyleToString = "wdNoteNumberStyleHindiCardinalText"
        Case wdNoteNumberStyleThaiLetter: WdNoteNumberStyleToString = "wdNoteNumberStyleThaiLetter"
        Case wdNoteNumberStyleThaiArabic: WdNoteNumberStyleToString = "wdNoteNumberStyleThaiArabic"
        Case wdNoteNumberStyleThaiCardinalText: WdNoteNumberStyleToString = "wdNoteNumberStyleThaiCardinalText"
        Case wdNoteNumberStyleVietCardinalText: WdNoteNumberStyleToString = "wdNoteNumberStyleVietCardinalText"
    End Select
End Function
