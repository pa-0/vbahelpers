Attribute VB_Name = "wXlLookFor"
Function XlLookForFromString(value As String) As XlLookFor
    If IsNumeric(value) Then
        XlLookForFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlLookForBlanks": XlLookForFromString = xlLookForBlanks
        Case "xlLookForErrors": XlLookForFromString = xlLookForErrors
        Case "xlLookForFormulas": XlLookForFromString = xlLookForFormulas
    End Select
End Function

Function XlLookForToString(value As XlLookFor) As String
    Select Case value
        Case xlLookForBlanks: XlLookForToString = "xlLookForBlanks"
        Case xlLookForErrors: XlLookForToString = "xlLookForErrors"
        Case xlLookForFormulas: XlLookForToString = "xlLookForFormulas"
    End Select
End Function
