Attribute VB_Name = "wXlDataSeriesDate"
Function XlDataSeriesDateFromString(value As String) As XlDataSeriesDate
    If IsNumeric(value) Then
        XlDataSeriesDateFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlDay": XlDataSeriesDateFromString = xlDay
        Case "xlWeekday": XlDataSeriesDateFromString = xlWeekday
        Case "xlMonth": XlDataSeriesDateFromString = xlMonth
        Case "xlYear": XlDataSeriesDateFromString = xlYear
    End Select
End Function

Function XlDataSeriesDateToString(value As XlDataSeriesDate) As String
    Select Case value
        Case xlDay: XlDataSeriesDateToString = "xlDay"
        Case xlWeekday: XlDataSeriesDateToString = "xlWeekday"
        Case xlMonth: XlDataSeriesDateToString = "xlMonth"
        Case xlYear: XlDataSeriesDateToString = "xlYear"
    End Select
End Function
