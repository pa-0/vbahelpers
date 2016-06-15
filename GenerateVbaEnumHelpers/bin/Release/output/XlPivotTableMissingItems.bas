Attribute VB_Name = "wXlPivotTableMissingItems"
Function XlPivotTableMissingItemsFromString(value As String) As XlPivotTableMissingItems
    If IsNumeric(value) Then
        XlPivotTableMissingItemsFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlMissingItemsNone": XlPivotTableMissingItemsFromString = xlMissingItemsNone
        Case "xlMissingItemsMax": XlPivotTableMissingItemsFromString = xlMissingItemsMax
        Case "xlMissingItemsMax2": XlPivotTableMissingItemsFromString = xlMissingItemsMax2
        Case "xlMissingItemsDefault": XlPivotTableMissingItemsFromString = xlMissingItemsDefault
    End Select
End Function

Function XlPivotTableMissingItemsToString(value As XlPivotTableMissingItems) As String
    Select Case value
        Case xlMissingItemsNone: XlPivotTableMissingItemsToString = "xlMissingItemsNone"
        Case xlMissingItemsMax: XlPivotTableMissingItemsToString = "xlMissingItemsMax"
        Case xlMissingItemsMax2: XlPivotTableMissingItemsToString = "xlMissingItemsMax2"
        Case xlMissingItemsDefault: XlPivotTableMissingItemsToString = "xlMissingItemsDefault"
    End Select
End Function
