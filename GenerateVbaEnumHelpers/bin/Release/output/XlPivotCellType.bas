Attribute VB_Name = "wXlPivotCellType"
Function XlPivotCellTypeFromString(value As String) As XlPivotCellType
    If IsNumeric(value) Then
        XlPivotCellTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlPivotCellValue": XlPivotCellTypeFromString = xlPivotCellValue
        Case "xlPivotCellPivotItem": XlPivotCellTypeFromString = xlPivotCellPivotItem
        Case "xlPivotCellSubtotal": XlPivotCellTypeFromString = xlPivotCellSubtotal
        Case "xlPivotCellGrandTotal": XlPivotCellTypeFromString = xlPivotCellGrandTotal
        Case "xlPivotCellDataField": XlPivotCellTypeFromString = xlPivotCellDataField
        Case "xlPivotCellPivotField": XlPivotCellTypeFromString = xlPivotCellPivotField
        Case "xlPivotCellPageFieldItem": XlPivotCellTypeFromString = xlPivotCellPageFieldItem
        Case "xlPivotCellCustomSubtotal": XlPivotCellTypeFromString = xlPivotCellCustomSubtotal
        Case "xlPivotCellDataPivotField": XlPivotCellTypeFromString = xlPivotCellDataPivotField
        Case "xlPivotCellBlankCell": XlPivotCellTypeFromString = xlPivotCellBlankCell
    End Select
End Function

Function XlPivotCellTypeToString(value As XlPivotCellType) As String
    Select Case value
        Case xlPivotCellValue: XlPivotCellTypeToString = "xlPivotCellValue"
        Case xlPivotCellPivotItem: XlPivotCellTypeToString = "xlPivotCellPivotItem"
        Case xlPivotCellSubtotal: XlPivotCellTypeToString = "xlPivotCellSubtotal"
        Case xlPivotCellGrandTotal: XlPivotCellTypeToString = "xlPivotCellGrandTotal"
        Case xlPivotCellDataField: XlPivotCellTypeToString = "xlPivotCellDataField"
        Case xlPivotCellPivotField: XlPivotCellTypeToString = "xlPivotCellPivotField"
        Case xlPivotCellPageFieldItem: XlPivotCellTypeToString = "xlPivotCellPageFieldItem"
        Case xlPivotCellCustomSubtotal: XlPivotCellTypeToString = "xlPivotCellCustomSubtotal"
        Case xlPivotCellDataPivotField: XlPivotCellTypeToString = "xlPivotCellDataPivotField"
        Case xlPivotCellBlankCell: XlPivotCellTypeToString = "xlPivotCellBlankCell"
    End Select
End Function
