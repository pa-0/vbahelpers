Attribute VB_Name = "wXlTimeUnit"
Function XlTimeUnitFromString(value As String) As XlTimeUnit
    If IsNumeric(value) Then
        XlTimeUnitFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlDays": XlTimeUnitFromString = xlDays
        Case "xlMonths": XlTimeUnitFromString = xlMonths
        Case "xlYears": XlTimeUnitFromString = xlYears
    End Select
End Function

Function XlTimeUnitToString(value As XlTimeUnit) As String
    Select Case value
        Case xlDays: XlTimeUnitToString = "xlDays"
        Case xlMonths: XlTimeUnitToString = "xlMonths"
        Case xlYears: XlTimeUnitToString = "xlYears"
    End Select
End Function
