Attribute VB_Name = "wXlCellInsertionMode"
Function XlCellInsertionModeFromString(value As String) As XlCellInsertionMode
    If IsNumeric(value) Then
        XlCellInsertionModeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlOverwriteCells": XlCellInsertionModeFromString = xlOverwriteCells
        Case "xlInsertDeleteCells": XlCellInsertionModeFromString = xlInsertDeleteCells
        Case "xlInsertEntireRows": XlCellInsertionModeFromString = xlInsertEntireRows
    End Select
End Function

Function XlCellInsertionModeToString(value As XlCellInsertionMode) As String
    Select Case value
        Case xlOverwriteCells: XlCellInsertionModeToString = "xlOverwriteCells"
        Case xlInsertDeleteCells: XlCellInsertionModeToString = "xlInsertDeleteCells"
        Case xlInsertEntireRows: XlCellInsertionModeToString = "xlInsertEntireRows"
    End Select
End Function
