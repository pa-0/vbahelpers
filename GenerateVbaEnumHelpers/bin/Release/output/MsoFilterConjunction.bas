Attribute VB_Name = "wMsoFilterConjunction"
Function MsoFilterConjunctionFromString(value As String) As MsoFilterConjunction
    If IsNumeric(value) Then
        MsoFilterConjunctionFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoFilterConjunctionAnd": MsoFilterConjunctionFromString = msoFilterConjunctionAnd
        Case "msoFilterConjunctionOr": MsoFilterConjunctionFromString = msoFilterConjunctionOr
    End Select
End Function

Function MsoFilterConjunctionToString(value As MsoFilterConjunction) As String
    Select Case value
        Case msoFilterConjunctionAnd: MsoFilterConjunctionToString = "msoFilterConjunctionAnd"
        Case msoFilterConjunctionOr: MsoFilterConjunctionToString = "msoFilterConjunctionOr"
    End Select
End Function
