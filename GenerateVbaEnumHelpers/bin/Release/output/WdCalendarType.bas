Attribute VB_Name = "wWdCalendarType"
Function WdCalendarTypeFromString(value As String) As WdCalendarType
    If IsNumeric(value) Then
        WdCalendarTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdCalendarWestern": WdCalendarTypeFromString = wdCalendarWestern
        Case "wdCalendarArabic": WdCalendarTypeFromString = wdCalendarArabic
        Case "wdCalendarHebrew": WdCalendarTypeFromString = wdCalendarHebrew
        Case "wdCalendarTaiwan": WdCalendarTypeFromString = wdCalendarTaiwan
        Case "wdCalendarJapan": WdCalendarTypeFromString = wdCalendarJapan
        Case "wdCalendarThai": WdCalendarTypeFromString = wdCalendarThai
        Case "wdCalendarKorean": WdCalendarTypeFromString = wdCalendarKorean
        Case "wdCalendarSakaEra": WdCalendarTypeFromString = wdCalendarSakaEra
        Case "wdCalendarTranslitEnglish": WdCalendarTypeFromString = wdCalendarTranslitEnglish
        Case "wdCalendarTranslitFrench": WdCalendarTypeFromString = wdCalendarTranslitFrench
        Case "wdCalendarUmalqura": WdCalendarTypeFromString = wdCalendarUmalqura
    End Select
End Function

Function WdCalendarTypeToString(value As WdCalendarType) As String
    Select Case value
        Case wdCalendarWestern: WdCalendarTypeToString = "wdCalendarWestern"
        Case wdCalendarArabic: WdCalendarTypeToString = "wdCalendarArabic"
        Case wdCalendarHebrew: WdCalendarTypeToString = "wdCalendarHebrew"
        Case wdCalendarTaiwan: WdCalendarTypeToString = "wdCalendarTaiwan"
        Case wdCalendarJapan: WdCalendarTypeToString = "wdCalendarJapan"
        Case wdCalendarThai: WdCalendarTypeToString = "wdCalendarThai"
        Case wdCalendarKorean: WdCalendarTypeToString = "wdCalendarKorean"
        Case wdCalendarSakaEra: WdCalendarTypeToString = "wdCalendarSakaEra"
        Case wdCalendarTranslitEnglish: WdCalendarTypeToString = "wdCalendarTranslitEnglish"
        Case wdCalendarTranslitFrench: WdCalendarTypeToString = "wdCalendarTranslitFrench"
        Case wdCalendarUmalqura: WdCalendarTypeToString = "wdCalendarUmalqura"
    End Select
End Function
