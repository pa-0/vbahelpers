Attribute VB_Name = "wWdInsertCells"
Function WdInsertCellsFromString(value As String) As WdInsertCells
    If IsNumeric(value) Then
        WdInsertCellsFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdInsertCellsShiftRight": WdInsertCellsFromString = wdInsertCellsShiftRight
        Case "wdInsertCellsShiftDown": WdInsertCellsFromString = wdInsertCellsShiftDown
        Case "wdInsertCellsEntireRow": WdInsertCellsFromString = wdInsertCellsEntireRow
        Case "wdInsertCellsEntireColumn": WdInsertCellsFromString = wdInsertCellsEntireColumn
    End Select
End Function

Function WdInsertCellsToString(value As WdInsertCells) As String
    Select Case value
        Case wdInsertCellsShiftRight: WdInsertCellsToString = "wdInsertCellsShiftRight"
        Case wdInsertCellsShiftDown: WdInsertCellsToString = "wdInsertCellsShiftDown"
        Case wdInsertCellsEntireRow: WdInsertCellsToString = "wdInsertCellsEntireRow"
        Case wdInsertCellsEntireColumn: WdInsertCellsToString = "wdInsertCellsEntireColumn"
    End Select
End Function
