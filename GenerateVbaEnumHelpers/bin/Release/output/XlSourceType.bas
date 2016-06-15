Attribute VB_Name = "wXlSourceType"
Function XlSourceTypeFromString(value As String) As XlSourceType
    If IsNumeric(value) Then
        XlSourceTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlSourceWorkbook": XlSourceTypeFromString = xlSourceWorkbook
        Case "xlSourceSheet": XlSourceTypeFromString = xlSourceSheet
        Case "xlSourcePrintArea": XlSourceTypeFromString = xlSourcePrintArea
        Case "xlSourceAutoFilter": XlSourceTypeFromString = xlSourceAutoFilter
        Case "xlSourceRange": XlSourceTypeFromString = xlSourceRange
        Case "xlSourceChart": XlSourceTypeFromString = xlSourceChart
        Case "xlSourcePivotTable": XlSourceTypeFromString = xlSourcePivotTable
        Case "xlSourceQuery": XlSourceTypeFromString = xlSourceQuery
    End Select
End Function

Function XlSourceTypeToString(value As XlSourceType) As String
    Select Case value
        Case xlSourceWorkbook: XlSourceTypeToString = "xlSourceWorkbook"
        Case xlSourceSheet: XlSourceTypeToString = "xlSourceSheet"
        Case xlSourcePrintArea: XlSourceTypeToString = "xlSourcePrintArea"
        Case xlSourceAutoFilter: XlSourceTypeToString = "xlSourceAutoFilter"
        Case xlSourceRange: XlSourceTypeToString = "xlSourceRange"
        Case xlSourceChart: XlSourceTypeToString = "xlSourceChart"
        Case xlSourcePivotTable: XlSourceTypeToString = "xlSourcePivotTable"
        Case xlSourceQuery: XlSourceTypeToString = "xlSourceQuery"
    End Select
End Function
