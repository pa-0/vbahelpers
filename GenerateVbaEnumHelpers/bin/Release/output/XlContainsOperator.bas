Attribute VB_Name = "wXlContainsOperator"
Function XlContainsOperatorFromString(value As String) As XlContainsOperator
    If IsNumeric(value) Then
        XlContainsOperatorFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlContains": XlContainsOperatorFromString = xlContains
        Case "xlDoesNotContain": XlContainsOperatorFromString = xlDoesNotContain
        Case "xlBeginsWith": XlContainsOperatorFromString = xlBeginsWith
        Case "xlEndsWith": XlContainsOperatorFromString = xlEndsWith
    End Select
End Function

Function XlContainsOperatorToString(value As XlContainsOperator) As String
    Select Case value
        Case xlContains: XlContainsOperatorToString = "xlContains"
        Case xlDoesNotContain: XlContainsOperatorToString = "xlDoesNotContain"
        Case xlBeginsWith: XlContainsOperatorToString = "xlBeginsWith"
        Case xlEndsWith: XlContainsOperatorToString = "xlEndsWith"
    End Select
End Function
