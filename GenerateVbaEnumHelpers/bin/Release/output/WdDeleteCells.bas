Attribute VB_Name = "wWdDeleteCells"
Function WdDeleteCellsFromString(value As String) As WdDeleteCells
    If IsNumeric(value) Then
        WdDeleteCellsFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdDeleteCellsShiftLeft": WdDeleteCellsFromString = wdDeleteCellsShiftLeft
        Case "wdDeleteCellsShiftUp": WdDeleteCellsFromString = wdDeleteCellsShiftUp
        Case "wdDeleteCellsEntireRow": WdDeleteCellsFromString = wdDeleteCellsEntireRow
        Case "wdDeleteCellsEntireColumn": WdDeleteCellsFromString = wdDeleteCellsEntireColumn
    End Select
End Function

Function WdDeleteCellsToString(value As WdDeleteCells) As String
    Select Case value
        Case wdDeleteCellsShiftLeft: WdDeleteCellsToString = "wdDeleteCellsShiftLeft"
        Case wdDeleteCellsShiftUp: WdDeleteCellsToString = "wdDeleteCellsShiftUp"
        Case wdDeleteCellsEntireRow: WdDeleteCellsToString = "wdDeleteCellsEntireRow"
        Case wdDeleteCellsEntireColumn: WdDeleteCellsToString = "wdDeleteCellsEntireColumn"
    End Select
End Function
