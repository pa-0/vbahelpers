Attribute VB_Name = "wPbDateTimeFormat"
Function PbDateTimeFormatFromString(value As String) As PbDateTimeFormat
    If IsNumeric(value) Then
        PbDateTimeFormatFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "pbDateShort": PbDateTimeFormatFromString = pbDateShort
        Case "pbDateLongDay": PbDateTimeFormatFromString = pbDateLongDay
        Case "pbDateLong": PbDateTimeFormatFromString = pbDateLong
        Case "pbDateShortAlt": PbDateTimeFormatFromString = pbDateShortAlt
        Case "pbDateISO": PbDateTimeFormatFromString = pbDateISO
        Case "pbDateShortMon": PbDateTimeFormatFromString = pbDateShortMon
        Case "pbDateShortSlash": PbDateTimeFormatFromString = pbDateShortSlash
        Case "pbDateShortAbb": PbDateTimeFormatFromString = pbDateShortAbb
        Case "pbDateEnglish": PbDateTimeFormatFromString = pbDateEnglish
        Case "pbDateMonthYr": PbDateTimeFormatFromString = pbDateMonthYr
        Case "pbDateMon_Yr": PbDateTimeFormatFromString = pbDateMon_Yr
        Case "pbTimeDatePM": PbDateTimeFormatFromString = pbTimeDatePM
        Case "pbTimeDateSecPM": PbDateTimeFormatFromString = pbTimeDateSecPM
        Case "pbTimePM": PbDateTimeFormatFromString = pbTimePM
        Case "pbTimeSecPM": PbDateTimeFormatFromString = pbTimeSecPM
        Case "pbTime24": PbDateTimeFormatFromString = pbTime24
        Case "pbTimeSec24": PbDateTimeFormatFromString = pbTimeSec24
        Case "pbDateTimeEastAsia1": PbDateTimeFormatFromString = pbDateTimeEastAsia1
        Case "pbDateTimeEastAsia2": PbDateTimeFormatFromString = pbDateTimeEastAsia2
        Case "pbDateTimeEastAsia3": PbDateTimeFormatFromString = pbDateTimeEastAsia3
        Case "pbDateTimeEastAsia4": PbDateTimeFormatFromString = pbDateTimeEastAsia4
        Case "pbDateTimeEastAsia5": PbDateTimeFormatFromString = pbDateTimeEastAsia5
    End Select
End Function

Function PbDateTimeFormatToString(value As PbDateTimeFormat) As String
    Select Case value
        Case pbDateShort: PbDateTimeFormatToString = "pbDateShort"
        Case pbDateLongDay: PbDateTimeFormatToString = "pbDateLongDay"
        Case pbDateLong: PbDateTimeFormatToString = "pbDateLong"
        Case pbDateShortAlt: PbDateTimeFormatToString = "pbDateShortAlt"
        Case pbDateISO: PbDateTimeFormatToString = "pbDateISO"
        Case pbDateShortMon: PbDateTimeFormatToString = "pbDateShortMon"
        Case pbDateShortSlash: PbDateTimeFormatToString = "pbDateShortSlash"
        Case pbDateShortAbb: PbDateTimeFormatToString = "pbDateShortAbb"
        Case pbDateEnglish: PbDateTimeFormatToString = "pbDateEnglish"
        Case pbDateMonthYr: PbDateTimeFormatToString = "pbDateMonthYr"
        Case pbDateMon_Yr: PbDateTimeFormatToString = "pbDateMon_Yr"
        Case pbTimeDatePM: PbDateTimeFormatToString = "pbTimeDatePM"
        Case pbTimeDateSecPM: PbDateTimeFormatToString = "pbTimeDateSecPM"
        Case pbTimePM: PbDateTimeFormatToString = "pbTimePM"
        Case pbTimeSecPM: PbDateTimeFormatToString = "pbTimeSecPM"
        Case pbTime24: PbDateTimeFormatToString = "pbTime24"
        Case pbTimeSec24: PbDateTimeFormatToString = "pbTimeSec24"
        Case pbDateTimeEastAsia1: PbDateTimeFormatToString = "pbDateTimeEastAsia1"
        Case pbDateTimeEastAsia2: PbDateTimeFormatToString = "pbDateTimeEastAsia2"
        Case pbDateTimeEastAsia3: PbDateTimeFormatToString = "pbDateTimeEastAsia3"
        Case pbDateTimeEastAsia4: PbDateTimeFormatToString = "pbDateTimeEastAsia4"
        Case pbDateTimeEastAsia5: PbDateTimeFormatToString = "pbDateTimeEastAsia5"
    End Select
End Function
