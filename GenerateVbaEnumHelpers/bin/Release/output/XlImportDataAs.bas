Attribute VB_Name = "wXlImportDataAs"
Function XlImportDataAsFromString(value As String) As XlImportDataAs
    If IsNumeric(value) Then
        XlImportDataAsFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlQueryTable": XlImportDataAsFromString = xlQueryTable
        Case "xlPivotTableReport": XlImportDataAsFromString = xlPivotTableReport
        Case "xlTable": XlImportDataAsFromString = xlTable
    End Select
End Function

Function XlImportDataAsToString(value As XlImportDataAs) As String
    Select Case value
        Case xlQueryTable: XlImportDataAsToString = "xlQueryTable"
        Case xlPivotTableReport: XlImportDataAsToString = "xlPivotTableReport"
        Case xlTable: XlImportDataAsToString = "xlTable"
    End Select
End Function
