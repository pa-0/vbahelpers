Attribute VB_Name = "wXlCommandUnderlines"
Function XlCommandUnderlinesFromString(value As String) As XlCommandUnderlines
    If IsNumeric(value) Then
        XlCommandUnderlinesFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlCommandUnderlinesOn": XlCommandUnderlinesFromString = xlCommandUnderlinesOn
        Case "xlCommandUnderlinesOff": XlCommandUnderlinesFromString = xlCommandUnderlinesOff
        Case "xlCommandUnderlinesAutomatic": XlCommandUnderlinesFromString = xlCommandUnderlinesAutomatic
    End Select
End Function

Function XlCommandUnderlinesToString(value As XlCommandUnderlines) As String
    Select Case value
        Case xlCommandUnderlinesOn: XlCommandUnderlinesToString = "xlCommandUnderlinesOn"
        Case xlCommandUnderlinesOff: XlCommandUnderlinesToString = "xlCommandUnderlinesOff"
        Case xlCommandUnderlinesAutomatic: XlCommandUnderlinesToString = "xlCommandUnderlinesAutomatic"
    End Select
End Function
