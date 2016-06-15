Attribute VB_Name = "wXlPropertyDisplayedIn"
Function XlPropertyDisplayedInFromString(value As String) As XlPropertyDisplayedIn
    If IsNumeric(value) Then
        XlPropertyDisplayedInFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlDisplayPropertyInPivotTable": XlPropertyDisplayedInFromString = xlDisplayPropertyInPivotTable
        Case "xlDisplayPropertyInTooltip": XlPropertyDisplayedInFromString = xlDisplayPropertyInTooltip
        Case "xlDisplayPropertyInPivotTableAndTooltip": XlPropertyDisplayedInFromString = xlDisplayPropertyInPivotTableAndTooltip
    End Select
End Function

Function XlPropertyDisplayedInToString(value As XlPropertyDisplayedIn) As String
    Select Case value
        Case xlDisplayPropertyInPivotTable: XlPropertyDisplayedInToString = "xlDisplayPropertyInPivotTable"
        Case xlDisplayPropertyInTooltip: XlPropertyDisplayedInToString = "xlDisplayPropertyInTooltip"
        Case xlDisplayPropertyInPivotTableAndTooltip: XlPropertyDisplayedInToString = "xlDisplayPropertyInPivotTableAndTooltip"
    End Select
End Function
