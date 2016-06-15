Attribute VB_Name = "wXlSearchOrder"
Function XlSearchOrderFromString(value As String) As XlSearchOrder
    If IsNumeric(value) Then
        XlSearchOrderFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlByRows": XlSearchOrderFromString = xlByRows
        Case "xlByColumns": XlSearchOrderFromString = xlByColumns
    End Select
End Function

Function XlSearchOrderToString(value As XlSearchOrder) As String
    Select Case value
        Case xlByRows: XlSearchOrderToString = "xlByRows"
        Case xlByColumns: XlSearchOrderToString = "xlByColumns"
    End Select
End Function
