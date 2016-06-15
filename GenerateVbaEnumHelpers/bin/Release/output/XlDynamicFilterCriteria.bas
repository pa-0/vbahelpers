Attribute VB_Name = "wXlDynamicFilterCriteria"
Function XlDynamicFilterCriteriaFromString(value As String) As XlDynamicFilterCriteria
    If IsNumeric(value) Then
        XlDynamicFilterCriteriaFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlFilterToday": XlDynamicFilterCriteriaFromString = xlFilterToday
        Case "xlFilterYesterday": XlDynamicFilterCriteriaFromString = xlFilterYesterday
        Case "xlFilterTomorrow": XlDynamicFilterCriteriaFromString = xlFilterTomorrow
        Case "xlFilterThisWeek": XlDynamicFilterCriteriaFromString = xlFilterThisWeek
        Case "xlFilterLastWeek": XlDynamicFilterCriteriaFromString = xlFilterLastWeek
        Case "xlFilterNextWeek": XlDynamicFilterCriteriaFromString = xlFilterNextWeek
        Case "xlFilterThisMonth": XlDynamicFilterCriteriaFromString = xlFilterThisMonth
        Case "xlFilterLastMonth": XlDynamicFilterCriteriaFromString = xlFilterLastMonth
        Case "xlFilterNextMonth": XlDynamicFilterCriteriaFromString = xlFilterNextMonth
        Case "xlFilterThisQuarter": XlDynamicFilterCriteriaFromString = xlFilterThisQuarter
        Case "xlFilterLastQuarter": XlDynamicFilterCriteriaFromString = xlFilterLastQuarter
        Case "xlFilterNextQuarter": XlDynamicFilterCriteriaFromString = xlFilterNextQuarter
        Case "xlFilterThisYear": XlDynamicFilterCriteriaFromString = xlFilterThisYear
        Case "xlFilterLastYear": XlDynamicFilterCriteriaFromString = xlFilterLastYear
        Case "xlFilterNextYear": XlDynamicFilterCriteriaFromString = xlFilterNextYear
        Case "xlFilterYearToDate": XlDynamicFilterCriteriaFromString = xlFilterYearToDate
        Case "xlFilterAllDatesInPeriodQuarter1": XlDynamicFilterCriteriaFromString = xlFilterAllDatesInPeriodQuarter1
        Case "xlFilterAllDatesInPeriodQuarter2": XlDynamicFilterCriteriaFromString = xlFilterAllDatesInPeriodQuarter2
        Case "xlFilterAllDatesInPeriodQuarter3": XlDynamicFilterCriteriaFromString = xlFilterAllDatesInPeriodQuarter3
        Case "xlFilterAllDatesInPeriodQuarter4": XlDynamicFilterCriteriaFromString = xlFilterAllDatesInPeriodQuarter4
        Case "xlFilterAllDatesInPeriodJanuary": XlDynamicFilterCriteriaFromString = xlFilterAllDatesInPeriodJanuary
        Case "xlFilterAllDatesInPeriodFebruray": XlDynamicFilterCriteriaFromString = xlFilterAllDatesInPeriodFebruray
        Case "xlFilterAllDatesInPeriodMarch": XlDynamicFilterCriteriaFromString = xlFilterAllDatesInPeriodMarch
        Case "xlFilterAllDatesInPeriodApril": XlDynamicFilterCriteriaFromString = xlFilterAllDatesInPeriodApril
        Case "xlFilterAllDatesInPeriodMay": XlDynamicFilterCriteriaFromString = xlFilterAllDatesInPeriodMay
        Case "xlFilterAllDatesInPeriodJune": XlDynamicFilterCriteriaFromString = xlFilterAllDatesInPeriodJune
        Case "xlFilterAllDatesInPeriodJuly": XlDynamicFilterCriteriaFromString = xlFilterAllDatesInPeriodJuly
        Case "xlFilterAllDatesInPeriodAugust": XlDynamicFilterCriteriaFromString = xlFilterAllDatesInPeriodAugust
        Case "xlFilterAllDatesInPeriodSeptember": XlDynamicFilterCriteriaFromString = xlFilterAllDatesInPeriodSeptember
        Case "xlFilterAllDatesInPeriodOctober": XlDynamicFilterCriteriaFromString = xlFilterAllDatesInPeriodOctober
        Case "xlFilterAllDatesInPeriodNovember": XlDynamicFilterCriteriaFromString = xlFilterAllDatesInPeriodNovember
        Case "xlFilterAllDatesInPeriodDecember": XlDynamicFilterCriteriaFromString = xlFilterAllDatesInPeriodDecember
        Case "xlFilterAboveAverage": XlDynamicFilterCriteriaFromString = xlFilterAboveAverage
        Case "xlFilterBelowAverage": XlDynamicFilterCriteriaFromString = xlFilterBelowAverage
    End Select
End Function

Function XlDynamicFilterCriteriaToString(value As XlDynamicFilterCriteria) As String
    Select Case value
        Case xlFilterToday: XlDynamicFilterCriteriaToString = "xlFilterToday"
        Case xlFilterYesterday: XlDynamicFilterCriteriaToString = "xlFilterYesterday"
        Case xlFilterTomorrow: XlDynamicFilterCriteriaToString = "xlFilterTomorrow"
        Case xlFilterThisWeek: XlDynamicFilterCriteriaToString = "xlFilterThisWeek"
        Case xlFilterLastWeek: XlDynamicFilterCriteriaToString = "xlFilterLastWeek"
        Case xlFilterNextWeek: XlDynamicFilterCriteriaToString = "xlFilterNextWeek"
        Case xlFilterThisMonth: XlDynamicFilterCriteriaToString = "xlFilterThisMonth"
        Case xlFilterLastMonth: XlDynamicFilterCriteriaToString = "xlFilterLastMonth"
        Case xlFilterNextMonth: XlDynamicFilterCriteriaToString = "xlFilterNextMonth"
        Case xlFilterThisQuarter: XlDynamicFilterCriteriaToString = "xlFilterThisQuarter"
        Case xlFilterLastQuarter: XlDynamicFilterCriteriaToString = "xlFilterLastQuarter"
        Case xlFilterNextQuarter: XlDynamicFilterCriteriaToString = "xlFilterNextQuarter"
        Case xlFilterThisYear: XlDynamicFilterCriteriaToString = "xlFilterThisYear"
        Case xlFilterLastYear: XlDynamicFilterCriteriaToString = "xlFilterLastYear"
        Case xlFilterNextYear: XlDynamicFilterCriteriaToString = "xlFilterNextYear"
        Case xlFilterYearToDate: XlDynamicFilterCriteriaToString = "xlFilterYearToDate"
        Case xlFilterAllDatesInPeriodQuarter1: XlDynamicFilterCriteriaToString = "xlFilterAllDatesInPeriodQuarter1"
        Case xlFilterAllDatesInPeriodQuarter2: XlDynamicFilterCriteriaToString = "xlFilterAllDatesInPeriodQuarter2"
        Case xlFilterAllDatesInPeriodQuarter3: XlDynamicFilterCriteriaToString = "xlFilterAllDatesInPeriodQuarter3"
        Case xlFilterAllDatesInPeriodQuarter4: XlDynamicFilterCriteriaToString = "xlFilterAllDatesInPeriodQuarter4"
        Case xlFilterAllDatesInPeriodJanuary: XlDynamicFilterCriteriaToString = "xlFilterAllDatesInPeriodJanuary"
        Case xlFilterAllDatesInPeriodFebruray: XlDynamicFilterCriteriaToString = "xlFilterAllDatesInPeriodFebruray"
        Case xlFilterAllDatesInPeriodMarch: XlDynamicFilterCriteriaToString = "xlFilterAllDatesInPeriodMarch"
        Case xlFilterAllDatesInPeriodApril: XlDynamicFilterCriteriaToString = "xlFilterAllDatesInPeriodApril"
        Case xlFilterAllDatesInPeriodMay: XlDynamicFilterCriteriaToString = "xlFilterAllDatesInPeriodMay"
        Case xlFilterAllDatesInPeriodJune: XlDynamicFilterCriteriaToString = "xlFilterAllDatesInPeriodJune"
        Case xlFilterAllDatesInPeriodJuly: XlDynamicFilterCriteriaToString = "xlFilterAllDatesInPeriodJuly"
        Case xlFilterAllDatesInPeriodAugust: XlDynamicFilterCriteriaToString = "xlFilterAllDatesInPeriodAugust"
        Case xlFilterAllDatesInPeriodSeptember: XlDynamicFilterCriteriaToString = "xlFilterAllDatesInPeriodSeptember"
        Case xlFilterAllDatesInPeriodOctober: XlDynamicFilterCriteriaToString = "xlFilterAllDatesInPeriodOctober"
        Case xlFilterAllDatesInPeriodNovember: XlDynamicFilterCriteriaToString = "xlFilterAllDatesInPeriodNovember"
        Case xlFilterAllDatesInPeriodDecember: XlDynamicFilterCriteriaToString = "xlFilterAllDatesInPeriodDecember"
        Case xlFilterAboveAverage: XlDynamicFilterCriteriaToString = "xlFilterAboveAverage"
        Case xlFilterBelowAverage: XlDynamicFilterCriteriaToString = "xlFilterBelowAverage"
    End Select
End Function
