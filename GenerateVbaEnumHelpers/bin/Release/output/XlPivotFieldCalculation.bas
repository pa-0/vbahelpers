Attribute VB_Name = "wXlPivotFieldCalculation"
Function XlPivotFieldCalculationFromString(value As String) As XlPivotFieldCalculation
    If IsNumeric(value) Then
        XlPivotFieldCalculationFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlDifferenceFrom": XlPivotFieldCalculationFromString = xlDifferenceFrom
        Case "xlPercentOf": XlPivotFieldCalculationFromString = xlPercentOf
        Case "xlPercentDifferenceFrom": XlPivotFieldCalculationFromString = xlPercentDifferenceFrom
        Case "xlRunningTotal": XlPivotFieldCalculationFromString = xlRunningTotal
        Case "xlPercentOfRow": XlPivotFieldCalculationFromString = xlPercentOfRow
        Case "xlPercentOfColumn": XlPivotFieldCalculationFromString = xlPercentOfColumn
        Case "xlPercentOfTotal": XlPivotFieldCalculationFromString = xlPercentOfTotal
        Case "xlIndex": XlPivotFieldCalculationFromString = xlIndex
        Case "xlPercentOfParentRow": XlPivotFieldCalculationFromString = xlPercentOfParentRow
        Case "xlPercentOfParentColumn": XlPivotFieldCalculationFromString = xlPercentOfParentColumn
        Case "xlPercentOfParent": XlPivotFieldCalculationFromString = xlPercentOfParent
        Case "xlPercentRunningTotal": XlPivotFieldCalculationFromString = xlPercentRunningTotal
        Case "xlRankAscending": XlPivotFieldCalculationFromString = xlRankAscending
        Case "xlRankDecending": XlPivotFieldCalculationFromString = xlRankDecending
        Case "xlNoAdditionalCalculation": XlPivotFieldCalculationFromString = xlNoAdditionalCalculation
    End Select
End Function

Function XlPivotFieldCalculationToString(value As XlPivotFieldCalculation) As String
    Select Case value
        Case xlDifferenceFrom: XlPivotFieldCalculationToString = "xlDifferenceFrom"
        Case xlPercentOf: XlPivotFieldCalculationToString = "xlPercentOf"
        Case xlPercentDifferenceFrom: XlPivotFieldCalculationToString = "xlPercentDifferenceFrom"
        Case xlRunningTotal: XlPivotFieldCalculationToString = "xlRunningTotal"
        Case xlPercentOfRow: XlPivotFieldCalculationToString = "xlPercentOfRow"
        Case xlPercentOfColumn: XlPivotFieldCalculationToString = "xlPercentOfColumn"
        Case xlPercentOfTotal: XlPivotFieldCalculationToString = "xlPercentOfTotal"
        Case xlIndex: XlPivotFieldCalculationToString = "xlIndex"
        Case xlPercentOfParentRow: XlPivotFieldCalculationToString = "xlPercentOfParentRow"
        Case xlPercentOfParentColumn: XlPivotFieldCalculationToString = "xlPercentOfParentColumn"
        Case xlPercentOfParent: XlPivotFieldCalculationToString = "xlPercentOfParent"
        Case xlPercentRunningTotal: XlPivotFieldCalculationToString = "xlPercentRunningTotal"
        Case xlRankAscending: XlPivotFieldCalculationToString = "xlRankAscending"
        Case xlRankDecending: XlPivotFieldCalculationToString = "xlRankDecending"
        Case xlNoAdditionalCalculation: XlPivotFieldCalculationToString = "xlNoAdditionalCalculation"
    End Select
End Function
