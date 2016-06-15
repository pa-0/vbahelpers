Attribute VB_Name = "wPbCalendarType"
Function PbCalendarTypeFromString(value As String) As PbCalendarType
    If IsNumeric(value) Then
        PbCalendarTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "pbCalendarTypeWestern": PbCalendarTypeFromString = pbCalendarTypeWestern
        Case "pbCalendarTypeArabicHijri": PbCalendarTypeFromString = pbCalendarTypeArabicHijri
        Case "pbCalendarTypeHebrewLunar": PbCalendarTypeFromString = pbCalendarTypeHebrewLunar
        Case "pbCalendarTypeChineseNational": PbCalendarTypeFromString = pbCalendarTypeChineseNational
        Case "pbCalendarTypeJapaneseEmperor": PbCalendarTypeFromString = pbCalendarTypeJapaneseEmperor
        Case "pbCalendarTypeThaiBuddhist": PbCalendarTypeFromString = pbCalendarTypeThaiBuddhist
        Case "pbCalendarTypeKoreanDanki": PbCalendarTypeFromString = pbCalendarTypeKoreanDanki
        Case "pbCalendarTypeSakaEra": PbCalendarTypeFromString = pbCalendarTypeSakaEra
        Case "pbCalendarTypeTranslitEnglish": PbCalendarTypeFromString = pbCalendarTypeTranslitEnglish
        Case "pbCalendarTypeTranslitFrench": PbCalendarTypeFromString = pbCalendarTypeTranslitFrench
        Case "pbCalendarTypeArabicUmalqura": PbCalendarTypeFromString = pbCalendarTypeArabicUmalqura
    End Select
End Function

Function PbCalendarTypeToString(value As PbCalendarType) As String
    Select Case value
        Case pbCalendarTypeWestern: PbCalendarTypeToString = "pbCalendarTypeWestern"
        Case pbCalendarTypeArabicHijri: PbCalendarTypeToString = "pbCalendarTypeArabicHijri"
        Case pbCalendarTypeHebrewLunar: PbCalendarTypeToString = "pbCalendarTypeHebrewLunar"
        Case pbCalendarTypeChineseNational: PbCalendarTypeToString = "pbCalendarTypeChineseNational"
        Case pbCalendarTypeJapaneseEmperor: PbCalendarTypeToString = "pbCalendarTypeJapaneseEmperor"
        Case pbCalendarTypeThaiBuddhist: PbCalendarTypeToString = "pbCalendarTypeThaiBuddhist"
        Case pbCalendarTypeKoreanDanki: PbCalendarTypeToString = "pbCalendarTypeKoreanDanki"
        Case pbCalendarTypeSakaEra: PbCalendarTypeToString = "pbCalendarTypeSakaEra"
        Case pbCalendarTypeTranslitEnglish: PbCalendarTypeToString = "pbCalendarTypeTranslitEnglish"
        Case pbCalendarTypeTranslitFrench: PbCalendarTypeToString = "pbCalendarTypeTranslitFrench"
        Case pbCalendarTypeArabicUmalqura: PbCalendarTypeToString = "pbCalendarTypeArabicUmalqura"
    End Select
End Function
