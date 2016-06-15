Attribute VB_Name = "wXlPivotLineType"
Function XlPivotLineTypeFromString(value As String) As XlPivotLineType
    If IsNumeric(value) Then
        XlPivotLineTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlPivotLineRegular": XlPivotLineTypeFromString = xlPivotLineRegular
        Case "xlPivotLineSubtotal": XlPivotLineTypeFromString = xlPivotLineSubtotal
        Case "xlPivotLineGrandTotal": XlPivotLineTypeFromString = xlPivotLineGrandTotal
        Case "xlPivotLineBlank": XlPivotLineTypeFromString = xlPivotLineBlank
    End Select
End Function

Function XlPivotLineTypeToString(value As XlPivotLineType) As String
    Select Case value
        Case xlPivotLineRegular: XlPivotLineTypeToString = "xlPivotLineRegular"
        Case xlPivotLineSubtotal: XlPivotLineTypeToString = "xlPivotLineSubtotal"
        Case xlPivotLineGrandTotal: XlPivotLineTypeToString = "xlPivotLineGrandTotal"
        Case xlPivotLineBlank: XlPivotLineTypeToString = "xlPivotLineBlank"
    End Select
End Function
