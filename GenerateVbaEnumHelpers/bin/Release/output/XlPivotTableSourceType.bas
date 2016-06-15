Attribute VB_Name = "wXlPivotTableSourceType"
Function XlPivotTableSourceTypeFromString(value As String) As XlPivotTableSourceType
    If IsNumeric(value) Then
        XlPivotTableSourceTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlDatabase": XlPivotTableSourceTypeFromString = xlDatabase
        Case "xlExternal": XlPivotTableSourceTypeFromString = xlExternal
        Case "xlConsolidation": XlPivotTableSourceTypeFromString = xlConsolidation
        Case "xlScenario": XlPivotTableSourceTypeFromString = xlScenario
        Case "xlPivotTable": XlPivotTableSourceTypeFromString = xlPivotTable
    End Select
End Function

Function XlPivotTableSourceTypeToString(value As XlPivotTableSourceType) As String
    Select Case value
        Case xlDatabase: XlPivotTableSourceTypeToString = "xlDatabase"
        Case xlExternal: XlPivotTableSourceTypeToString = "xlExternal"
        Case xlConsolidation: XlPivotTableSourceTypeToString = "xlConsolidation"
        Case xlScenario: XlPivotTableSourceTypeToString = "xlScenario"
        Case xlPivotTable: XlPivotTableSourceTypeToString = "xlPivotTable"
    End Select
End Function
