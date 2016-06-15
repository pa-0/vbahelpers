Attribute VB_Name = "wXlOrder"
Function XlOrderFromString(value As String) As XlOrder
    If IsNumeric(value) Then
        XlOrderFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlDownThenOver": XlOrderFromString = xlDownThenOver
        Case "xlOverThenDown": XlOrderFromString = xlOverThenDown
    End Select
End Function

Function XlOrderToString(value As XlOrder) As String
    Select Case value
        Case xlDownThenOver: XlOrderToString = "xlDownThenOver"
        Case xlOverThenDown: XlOrderToString = "xlOverThenDown"
    End Select
End Function
