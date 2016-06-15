Attribute VB_Name = "wXlAxisGroup"
Function XlAxisGroupFromString(value As String) As XlAxisGroup
    If IsNumeric(value) Then
        XlAxisGroupFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlPrimary": XlAxisGroupFromString = xlPrimary
        Case "xlSecondary": XlAxisGroupFromString = xlSecondary
    End Select
End Function

Function XlAxisGroupToString(value As XlAxisGroup) As String
    Select Case value
        Case xlPrimary: XlAxisGroupToString = "xlPrimary"
        Case xlSecondary: XlAxisGroupToString = "xlSecondary"
    End Select
End Function
