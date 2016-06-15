Attribute VB_Name = "wXlSearchWithin"
Function XlSearchWithinFromString(value As String) As XlSearchWithin
    If IsNumeric(value) Then
        XlSearchWithinFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlWithinSheet": XlSearchWithinFromString = xlWithinSheet
        Case "xlWithinWorkbook": XlSearchWithinFromString = xlWithinWorkbook
    End Select
End Function

Function XlSearchWithinToString(value As XlSearchWithin) As String
    Select Case value
        Case xlWithinSheet: XlSearchWithinToString = "xlWithinSheet"
        Case xlWithinWorkbook: XlSearchWithinToString = "xlWithinWorkbook"
    End Select
End Function
