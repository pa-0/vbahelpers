Attribute VB_Name = "wPbPageNumberFormat"
Function PbPageNumberFormatFromString(value As String) As PbPageNumberFormat
    If IsNumeric(value) Then
        PbPageNumberFormatFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "pbPageNumberFormatArabic": PbPageNumberFormatFromString = pbPageNumberFormatArabic
        Case "pbPageNumberFormatUCRoman": PbPageNumberFormatFromString = pbPageNumberFormatUCRoman
        Case "pbPageNumberFormatLCRoman": PbPageNumberFormatFromString = pbPageNumberFormatLCRoman
        Case "pbPageNumberFormatUCLetter": PbPageNumberFormatFromString = pbPageNumberFormatUCLetter
        Case "pbPageNumberFormatLCLetter": PbPageNumberFormatFromString = pbPageNumberFormatLCLetter
        Case "pbPageNumberFormatOrdinal": PbPageNumberFormatFromString = pbPageNumberFormatOrdinal
        Case "pbPageNumberFormatCardtext": PbPageNumberFormatFromString = pbPageNumberFormatCardtext
        Case "pbPageNumberFormatOrdtext": PbPageNumberFormatFromString = pbPageNumberFormatOrdtext
        Case "pbPageNumberFormatDbNum1": PbPageNumberFormatFromString = pbPageNumberFormatDbNum1
        Case "pbPageNumberFormatDbNum2": PbPageNumberFormatFromString = pbPageNumberFormatDbNum2
        Case "pbPageNumberFormatAiueo": PbPageNumberFormatFromString = pbPageNumberFormatAiueo
        Case "pbPageNumberFormatIroha": PbPageNumberFormatFromString = pbPageNumberFormatIroha
        Case "pbPageNumberFormatDbChar": PbPageNumberFormatFromString = pbPageNumberFormatDbChar
        Case "pbPageNumberFormatDbNum3": PbPageNumberFormatFromString = pbPageNumberFormatDbNum3
        Case "pbPageNumberFormatCirclenum": PbPageNumberFormatFromString = pbPageNumberFormatCirclenum
        Case "pbPageNumberFormatDAiueo": PbPageNumberFormatFromString = pbPageNumberFormatDAiueo
        Case "pbPageNumberFormatDIroha": PbPageNumberFormatFromString = pbPageNumberFormatDIroha
        Case "pbPageNumberFormatArabicLZ": PbPageNumberFormatFromString = pbPageNumberFormatArabicLZ
        Case "pbPageNumberFormatGanada": PbPageNumberFormatFromString = pbPageNumberFormatGanada
        Case "pbPageNumberFormatChosung": PbPageNumberFormatFromString = pbPageNumberFormatChosung
        Case "pbPageNumberFormatZodiac1": PbPageNumberFormatFromString = pbPageNumberFormatZodiac1
        Case "pbPageNumberFormatZodiac2": PbPageNumberFormatFromString = pbPageNumberFormatZodiac2
        Case "pbPageNumberFormatTpeDbNum2": PbPageNumberFormatFromString = pbPageNumberFormatTpeDbNum2
        Case "pbPageNumberFormatTpeDbNum3": PbPageNumberFormatFromString = pbPageNumberFormatTpeDbNum3
        Case "pbPageNumberFormatChnDbNum2": PbPageNumberFormatFromString = pbPageNumberFormatChnDbNum2
        Case "pbPageNumberFormatChnDbNum3": PbPageNumberFormatFromString = pbPageNumberFormatChnDbNum3
        Case "pbPageNumberFormatKorDbNum1": PbPageNumberFormatFromString = pbPageNumberFormatKorDbNum1
        Case "pbPageNumberFormatKorDbNum2": PbPageNumberFormatFromString = pbPageNumberFormatKorDbNum2
        Case "pbPageNumberFormatKorDbNum3": PbPageNumberFormatFromString = pbPageNumberFormatKorDbNum3
        Case "pbPageNumberFormatKorDbNum4": PbPageNumberFormatFromString = pbPageNumberFormatKorDbNum4
        Case "pbPageNumberFormatHebrew1": PbPageNumberFormatFromString = pbPageNumberFormatHebrew1
        Case "pbPageNumberFormatArabic1": PbPageNumberFormatFromString = pbPageNumberFormatArabic1
        Case "pbPageNumberFormatHebrew2": PbPageNumberFormatFromString = pbPageNumberFormatHebrew2
        Case "pbPageNumberFormatArabic2": PbPageNumberFormatFromString = pbPageNumberFormatArabic2
        Case "pbPageNumberFormatHindi1": PbPageNumberFormatFromString = pbPageNumberFormatHindi1
        Case "pbPageNumberFormatHindi2": PbPageNumberFormatFromString = pbPageNumberFormatHindi2
        Case "pbPageNumberFormatHindi3": PbPageNumberFormatFromString = pbPageNumberFormatHindi3
        Case "pbPageNumberFormatHindi4": PbPageNumberFormatFromString = pbPageNumberFormatHindi4
        Case "pbPageNumberFormatThai1": PbPageNumberFormatFromString = pbPageNumberFormatThai1
        Case "pbPageNumberFormatThai2": PbPageNumberFormatFromString = pbPageNumberFormatThai2
        Case "pbPageNumberFormatThai3": PbPageNumberFormatFromString = pbPageNumberFormatThai3
        Case "pbPageNumberFormatViet1": PbPageNumberFormatFromString = pbPageNumberFormatViet1
        Case "pbPageNumberFormatLCRus": PbPageNumberFormatFromString = pbPageNumberFormatLCRus
        Case "pbPageNumberFormatUCRus": PbPageNumberFormatFromString = pbPageNumberFormatUCRus
    End Select
End Function

Function PbPageNumberFormatToString(value As PbPageNumberFormat) As String
    Select Case value
        Case pbPageNumberFormatArabic: PbPageNumberFormatToString = "pbPageNumberFormatArabic"
        Case pbPageNumberFormatUCRoman: PbPageNumberFormatToString = "pbPageNumberFormatUCRoman"
        Case pbPageNumberFormatLCRoman: PbPageNumberFormatToString = "pbPageNumberFormatLCRoman"
        Case pbPageNumberFormatUCLetter: PbPageNumberFormatToString = "pbPageNumberFormatUCLetter"
        Case pbPageNumberFormatLCLetter: PbPageNumberFormatToString = "pbPageNumberFormatLCLetter"
        Case pbPageNumberFormatOrdinal: PbPageNumberFormatToString = "pbPageNumberFormatOrdinal"
        Case pbPageNumberFormatCardtext: PbPageNumberFormatToString = "pbPageNumberFormatCardtext"
        Case pbPageNumberFormatOrdtext: PbPageNumberFormatToString = "pbPageNumberFormatOrdtext"
        Case pbPageNumberFormatDbNum1: PbPageNumberFormatToString = "pbPageNumberFormatDbNum1"
        Case pbPageNumberFormatDbNum2: PbPageNumberFormatToString = "pbPageNumberFormatDbNum2"
        Case pbPageNumberFormatAiueo: PbPageNumberFormatToString = "pbPageNumberFormatAiueo"
        Case pbPageNumberFormatIroha: PbPageNumberFormatToString = "pbPageNumberFormatIroha"
        Case pbPageNumberFormatDbChar: PbPageNumberFormatToString = "pbPageNumberFormatDbChar"
        Case pbPageNumberFormatDbNum3: PbPageNumberFormatToString = "pbPageNumberFormatDbNum3"
        Case pbPageNumberFormatCirclenum: PbPageNumberFormatToString = "pbPageNumberFormatCirclenum"
        Case pbPageNumberFormatDAiueo: PbPageNumberFormatToString = "pbPageNumberFormatDAiueo"
        Case pbPageNumberFormatDIroha: PbPageNumberFormatToString = "pbPageNumberFormatDIroha"
        Case pbPageNumberFormatArabicLZ: PbPageNumberFormatToString = "pbPageNumberFormatArabicLZ"
        Case pbPageNumberFormatGanada: PbPageNumberFormatToString = "pbPageNumberFormatGanada"
        Case pbPageNumberFormatChosung: PbPageNumberFormatToString = "pbPageNumberFormatChosung"
        Case pbPageNumberFormatZodiac1: PbPageNumberFormatToString = "pbPageNumberFormatZodiac1"
        Case pbPageNumberFormatZodiac2: PbPageNumberFormatToString = "pbPageNumberFormatZodiac2"
        Case pbPageNumberFormatTpeDbNum2: PbPageNumberFormatToString = "pbPageNumberFormatTpeDbNum2"
        Case pbPageNumberFormatTpeDbNum3: PbPageNumberFormatToString = "pbPageNumberFormatTpeDbNum3"
        Case pbPageNumberFormatChnDbNum2: PbPageNumberFormatToString = "pbPageNumberFormatChnDbNum2"
        Case pbPageNumberFormatChnDbNum3: PbPageNumberFormatToString = "pbPageNumberFormatChnDbNum3"
        Case pbPageNumberFormatKorDbNum1: PbPageNumberFormatToString = "pbPageNumberFormatKorDbNum1"
        Case pbPageNumberFormatKorDbNum2: PbPageNumberFormatToString = "pbPageNumberFormatKorDbNum2"
        Case pbPageNumberFormatKorDbNum3: PbPageNumberFormatToString = "pbPageNumberFormatKorDbNum3"
        Case pbPageNumberFormatKorDbNum4: PbPageNumberFormatToString = "pbPageNumberFormatKorDbNum4"
        Case pbPageNumberFormatHebrew1: PbPageNumberFormatToString = "pbPageNumberFormatHebrew1"
        Case pbPageNumberFormatArabic1: PbPageNumberFormatToString = "pbPageNumberFormatArabic1"
        Case pbPageNumberFormatHebrew2: PbPageNumberFormatToString = "pbPageNumberFormatHebrew2"
        Case pbPageNumberFormatArabic2: PbPageNumberFormatToString = "pbPageNumberFormatArabic2"
        Case pbPageNumberFormatHindi1: PbPageNumberFormatToString = "pbPageNumberFormatHindi1"
        Case pbPageNumberFormatHindi2: PbPageNumberFormatToString = "pbPageNumberFormatHindi2"
        Case pbPageNumberFormatHindi3: PbPageNumberFormatToString = "pbPageNumberFormatHindi3"
        Case pbPageNumberFormatHindi4: PbPageNumberFormatToString = "pbPageNumberFormatHindi4"
        Case pbPageNumberFormatThai1: PbPageNumberFormatToString = "pbPageNumberFormatThai1"
        Case pbPageNumberFormatThai2: PbPageNumberFormatToString = "pbPageNumberFormatThai2"
        Case pbPageNumberFormatThai3: PbPageNumberFormatToString = "pbPageNumberFormatThai3"
        Case pbPageNumberFormatViet1: PbPageNumberFormatToString = "pbPageNumberFormatViet1"
        Case pbPageNumberFormatLCRus: PbPageNumberFormatToString = "pbPageNumberFormatLCRus"
        Case pbPageNumberFormatUCRus: PbPageNumberFormatToString = "pbPageNumberFormatUCRus"
    End Select
End Function
