Attribute VB_Name = "wWdRulerStyle"
Function WdRulerStyleFromString(value As String) As WdRulerStyle
    If IsNumeric(value) Then
        WdRulerStyleFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdAdjustNone": WdRulerStyleFromString = wdAdjustNone
        Case "wdAdjustProportional": WdRulerStyleFromString = wdAdjustProportional
        Case "wdAdjustFirstColumn": WdRulerStyleFromString = wdAdjustFirstColumn
        Case "wdAdjustSameWidth": WdRulerStyleFromString = wdAdjustSameWidth
    End Select
End Function

Function WdRulerStyleToString(value As WdRulerStyle) As String
    Select Case value
        Case wdAdjustNone: WdRulerStyleToString = "wdAdjustNone"
        Case wdAdjustProportional: WdRulerStyleToString = "wdAdjustProportional"
        Case wdAdjustFirstColumn: WdRulerStyleToString = "wdAdjustFirstColumn"
        Case wdAdjustSameWidth: WdRulerStyleToString = "wdAdjustSameWidth"
    End Select
End Function
