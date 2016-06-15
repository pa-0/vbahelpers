Attribute VB_Name = "wXlPriority"
Function XlPriorityFromString(value As String) As XlPriority
    If IsNumeric(value) Then
        XlPriorityFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlPriorityNormal": XlPriorityFromString = xlPriorityNormal
        Case "xlPriorityLow": XlPriorityFromString = xlPriorityLow
        Case "xlPriorityHigh": XlPriorityFromString = xlPriorityHigh
    End Select
End Function

Function XlPriorityToString(value As XlPriority) As String
    Select Case value
        Case xlPriorityNormal: XlPriorityToString = "xlPriorityNormal"
        Case xlPriorityLow: XlPriorityToString = "xlPriorityLow"
        Case xlPriorityHigh: XlPriorityToString = "xlPriorityHigh"
    End Select
End Function
