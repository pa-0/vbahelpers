Attribute VB_Name = "wXlApplyNamesOrder"
Function XlApplyNamesOrderFromString(value As String) As XlApplyNamesOrder
    If IsNumeric(value) Then
        XlApplyNamesOrderFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlRowThenColumn": XlApplyNamesOrderFromString = xlRowThenColumn
        Case "xlColumnThenRow": XlApplyNamesOrderFromString = xlColumnThenRow
    End Select
End Function

Function XlApplyNamesOrderToString(value As XlApplyNamesOrder) As String
    Select Case value
        Case xlRowThenColumn: XlApplyNamesOrderToString = "xlRowThenColumn"
        Case xlColumnThenRow: XlApplyNamesOrderToString = "xlColumnThenRow"
    End Select
End Function
