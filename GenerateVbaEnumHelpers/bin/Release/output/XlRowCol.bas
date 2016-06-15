Attribute VB_Name = "wXlRowCol"
Function XlRowColFromString(value As String) As XlRowCol
    If IsNumeric(value) Then
        XlRowColFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlRows": XlRowColFromString = xlRows
        Case "xlColumns": XlRowColFromString = xlColumns
    End Select
End Function

Function XlRowColToString(value As XlRowCol) As String
    Select Case value
        Case xlRows: XlRowColToString = "xlRows"
        Case xlColumns: XlRowColToString = "xlColumns"
    End Select
End Function
