Attribute VB_Name = "wXlEndStyleCap"
Function XlEndStyleCapFromString(value As String) As XlEndStyleCap
    If IsNumeric(value) Then
        XlEndStyleCapFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlCap": XlEndStyleCapFromString = xlCap
        Case "xlNoCap": XlEndStyleCapFromString = xlNoCap
    End Select
End Function

Function XlEndStyleCapToString(value As XlEndStyleCap) As String
    Select Case value
        Case xlCap: XlEndStyleCapToString = "xlCap"
        Case xlNoCap: XlEndStyleCapToString = "xlNoCap"
    End Select
End Function
