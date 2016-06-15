Attribute VB_Name = "wXlPageBreak"
Function XlPageBreakFromString(value As String) As XlPageBreak
    If IsNumeric(value) Then
        XlPageBreakFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlPageBreakNone": XlPageBreakFromString = xlPageBreakNone
        Case "xlPageBreakManual": XlPageBreakFromString = xlPageBreakManual
        Case "xlPageBreakAutomatic": XlPageBreakFromString = xlPageBreakAutomatic
    End Select
End Function

Function XlPageBreakToString(value As XlPageBreak) As String
    Select Case value
        Case xlPageBreakNone: XlPageBreakToString = "xlPageBreakNone"
        Case xlPageBreakManual: XlPageBreakToString = "xlPageBreakManual"
        Case xlPageBreakAutomatic: XlPageBreakToString = "xlPageBreakAutomatic"
    End Select
End Function
