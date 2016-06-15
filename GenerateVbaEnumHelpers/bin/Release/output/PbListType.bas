Attribute VB_Name = "wPbListType"
Function PbListTypeFromString(value As String) As PbListType
    If IsNumeric(value) Then
        PbListTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "pbListTypeArabic": PbListTypeFromString = pbListTypeArabic
        Case "pbListTypeUpperCaseRoman": PbListTypeFromString = pbListTypeUpperCaseRoman
        Case "pbListTypeLowerCaseRoman": PbListTypeFromString = pbListTypeLowerCaseRoman
        Case "pbListTypeUpperCaseLetter": PbListTypeFromString = pbListTypeUpperCaseLetter
        Case "pbListTypeLowerCaseLetter": PbListTypeFromString = pbListTypeLowerCaseLetter
        Case "pbListTypeOrdinal": PbListTypeFromString = pbListTypeOrdinal
        Case "pbListTypeCardinalText": PbListTypeFromString = pbListTypeCardinalText
        Case "pbListTypeOrdinalText": PbListTypeFromString = pbListTypeOrdinalText
        Case "pbListTypeDbNum1": PbListTypeFromString = pbListTypeDbNum1
        Case "pbListTypeDbNum2": PbListTypeFromString = pbListTypeDbNum2
        Case "pbListTypeAiueo": PbListTypeFromString = pbListTypeAiueo
        Case "pbListTypeIroha": PbListTypeFromString = pbListTypeIroha
        Case "pbListTypeDbChar": PbListTypeFromString = pbListTypeDbChar
        Case "pbListTypeDbNum3": PbListTypeFromString = pbListTypeDbNum3
        Case "pbListTypeDbNum4": PbListTypeFromString = pbListTypeDbNum4
        Case "pbListTypeCirclenum": PbListTypeFromString = pbListTypeCirclenum
        Case "pbListTypeDAiueo": PbListTypeFromString = pbListTypeDAiueo
        Case "pbListTypeDIroha": PbListTypeFromString = pbListTypeDIroha
        Case "pbListTypeArabicLeadingZero": PbListTypeFromString = pbListTypeArabicLeadingZero
        Case "pbListTypeBullet": PbListTypeFromString = pbListTypeBullet
        Case "pbListTypeGanada": PbListTypeFromString = pbListTypeGanada
        Case "pbListTypeChosung": PbListTypeFromString = pbListTypeChosung
        Case "pbListTypeZodiac1": PbListTypeFromString = pbListTypeZodiac1
        Case "pbListTypeZodiac2": PbListTypeFromString = pbListTypeZodiac2
        Case "pbListTypeTpeDbNum2": PbListTypeFromString = pbListTypeTpeDbNum2
        Case "pbListTypeTpeDbNum3": PbListTypeFromString = pbListTypeTpeDbNum3
        Case "pbListTypeChnDbNum2": PbListTypeFromString = pbListTypeChnDbNum2
        Case "pbListTypeChnDbNum3": PbListTypeFromString = pbListTypeChnDbNum3
        Case "pbListTypeKorDbNum1": PbListTypeFromString = pbListTypeKorDbNum1
        Case "pbListTypeKorDbNum2": PbListTypeFromString = pbListTypeKorDbNum2
        Case "pbListTypeKorDbNum3": PbListTypeFromString = pbListTypeKorDbNum3
        Case "pbListTypeKorDbNum4": PbListTypeFromString = pbListTypeKorDbNum4
        Case "pbListTypeHebrew1": PbListTypeFromString = pbListTypeHebrew1
        Case "pbListTypeArabic1": PbListTypeFromString = pbListTypeArabic1
        Case "pbListTypeHebrew2": PbListTypeFromString = pbListTypeHebrew2
        Case "pbListTypeArabic2": PbListTypeFromString = pbListTypeArabic2
        Case "pbListTypeHindi1": PbListTypeFromString = pbListTypeHindi1
        Case "pbListTypeHindi2": PbListTypeFromString = pbListTypeHindi2
        Case "pbListTypeHindi3": PbListTypeFromString = pbListTypeHindi3
        Case "pbListTypeHindi4": PbListTypeFromString = pbListTypeHindi4
        Case "pbListTypeThai1": PbListTypeFromString = pbListTypeThai1
        Case "pbListTypeThai2": PbListTypeFromString = pbListTypeThai2
        Case "pbListTypeThai3": PbListTypeFromString = pbListTypeThai3
        Case "pbListTypeVietnamese1": PbListTypeFromString = pbListTypeVietnamese1
        Case "pbListTypeLowerCaseRussian": PbListTypeFromString = pbListTypeLowerCaseRussian
        Case "pbListTypeUpperCaseRussian": PbListTypeFromString = pbListTypeUpperCaseRussian
        Case "pbListTypeNone": PbListTypeFromString = pbListTypeNone
    End Select
End Function

Function PbListTypeToString(value As PbListType) As String
    Select Case value
        Case pbListTypeArabic: PbListTypeToString = "pbListTypeArabic"
        Case pbListTypeUpperCaseRoman: PbListTypeToString = "pbListTypeUpperCaseRoman"
        Case pbListTypeLowerCaseRoman: PbListTypeToString = "pbListTypeLowerCaseRoman"
        Case pbListTypeUpperCaseLetter: PbListTypeToString = "pbListTypeUpperCaseLetter"
        Case pbListTypeLowerCaseLetter: PbListTypeToString = "pbListTypeLowerCaseLetter"
        Case pbListTypeOrdinal: PbListTypeToString = "pbListTypeOrdinal"
        Case pbListTypeCardinalText: PbListTypeToString = "pbListTypeCardinalText"
        Case pbListTypeOrdinalText: PbListTypeToString = "pbListTypeOrdinalText"
        Case pbListTypeDbNum1: PbListTypeToString = "pbListTypeDbNum1"
        Case pbListTypeDbNum2: PbListTypeToString = "pbListTypeDbNum2"
        Case pbListTypeAiueo: PbListTypeToString = "pbListTypeAiueo"
        Case pbListTypeIroha: PbListTypeToString = "pbListTypeIroha"
        Case pbListTypeDbChar: PbListTypeToString = "pbListTypeDbChar"
        Case pbListTypeDbNum3: PbListTypeToString = "pbListTypeDbNum3"
        Case pbListTypeDbNum4: PbListTypeToString = "pbListTypeDbNum4"
        Case pbListTypeCirclenum: PbListTypeToString = "pbListTypeCirclenum"
        Case pbListTypeDAiueo: PbListTypeToString = "pbListTypeDAiueo"
        Case pbListTypeDIroha: PbListTypeToString = "pbListTypeDIroha"
        Case pbListTypeArabicLeadingZero: PbListTypeToString = "pbListTypeArabicLeadingZero"
        Case pbListTypeBullet: PbListTypeToString = "pbListTypeBullet"
        Case pbListTypeGanada: PbListTypeToString = "pbListTypeGanada"
        Case pbListTypeChosung: PbListTypeToString = "pbListTypeChosung"
        Case pbListTypeZodiac1: PbListTypeToString = "pbListTypeZodiac1"
        Case pbListTypeZodiac2: PbListTypeToString = "pbListTypeZodiac2"
        Case pbListTypeTpeDbNum2: PbListTypeToString = "pbListTypeTpeDbNum2"
        Case pbListTypeTpeDbNum3: PbListTypeToString = "pbListTypeTpeDbNum3"
        Case pbListTypeChnDbNum2: PbListTypeToString = "pbListTypeChnDbNum2"
        Case pbListTypeChnDbNum3: PbListTypeToString = "pbListTypeChnDbNum3"
        Case pbListTypeKorDbNum1: PbListTypeToString = "pbListTypeKorDbNum1"
        Case pbListTypeKorDbNum2: PbListTypeToString = "pbListTypeKorDbNum2"
        Case pbListTypeKorDbNum3: PbListTypeToString = "pbListTypeKorDbNum3"
        Case pbListTypeKorDbNum4: PbListTypeToString = "pbListTypeKorDbNum4"
        Case pbListTypeHebrew1: PbListTypeToString = "pbListTypeHebrew1"
        Case pbListTypeArabic1: PbListTypeToString = "pbListTypeArabic1"
        Case pbListTypeHebrew2: PbListTypeToString = "pbListTypeHebrew2"
        Case pbListTypeArabic2: PbListTypeToString = "pbListTypeArabic2"
        Case pbListTypeHindi1: PbListTypeToString = "pbListTypeHindi1"
        Case pbListTypeHindi2: PbListTypeToString = "pbListTypeHindi2"
        Case pbListTypeHindi3: PbListTypeToString = "pbListTypeHindi3"
        Case pbListTypeHindi4: PbListTypeToString = "pbListTypeHindi4"
        Case pbListTypeThai1: PbListTypeToString = "pbListTypeThai1"
        Case pbListTypeThai2: PbListTypeToString = "pbListTypeThai2"
        Case pbListTypeThai3: PbListTypeToString = "pbListTypeThai3"
        Case pbListTypeVietnamese1: PbListTypeToString = "pbListTypeVietnamese1"
        Case pbListTypeLowerCaseRussian: PbListTypeToString = "pbListTypeLowerCaseRussian"
        Case pbListTypeUpperCaseRussian: PbListTypeToString = "pbListTypeUpperCaseRussian"
        Case pbListTypeNone: PbListTypeToString = "pbListTypeNone"
    End Select
End Function
