Attribute VB_Name = "wXlAxisCrosses"
Function XlAxisCrossesFromString(value As String) As XlAxisCrosses
    If IsNumeric(value) Then
        XlAxisCrossesFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlAxisCrossesMaximum": XlAxisCrossesFromString = xlAxisCrossesMaximum
        Case "xlAxisCrossesMinimum": XlAxisCrossesFromString = xlAxisCrossesMinimum
        Case "xlAxisCrossesCustom": XlAxisCrossesFromString = xlAxisCrossesCustom
        Case "xlAxisCrossesAutomatic": XlAxisCrossesFromString = xlAxisCrossesAutomatic
    End Select
End Function

Function XlAxisCrossesToString(value As XlAxisCrosses) As String
    Select Case value
        Case xlAxisCrossesMaximum: XlAxisCrossesToString = "xlAxisCrossesMaximum"
        Case xlAxisCrossesMinimum: XlAxisCrossesToString = "xlAxisCrossesMinimum"
        Case xlAxisCrossesCustom: XlAxisCrossesToString = "xlAxisCrossesCustom"
        Case xlAxisCrossesAutomatic: XlAxisCrossesToString = "xlAxisCrossesAutomatic"
    End Select
End Function
