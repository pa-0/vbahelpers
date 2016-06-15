Attribute VB_Name = "wXlArrangeStyle"
Function XlArrangeStyleFromString(value As String) As XlArrangeStyle
    If IsNumeric(value) Then
        XlArrangeStyleFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlArrangeStyleTiled": XlArrangeStyleFromString = xlArrangeStyleTiled
        Case "xlArrangeStyleCascade": XlArrangeStyleFromString = xlArrangeStyleCascade
        Case "xlArrangeStyleVertical": XlArrangeStyleFromString = xlArrangeStyleVertical
        Case "xlArrangeStyleHorizontal": XlArrangeStyleFromString = xlArrangeStyleHorizontal
    End Select
End Function

Function XlArrangeStyleToString(value As XlArrangeStyle) As String
    Select Case value
        Case xlArrangeStyleTiled: XlArrangeStyleToString = "xlArrangeStyleTiled"
        Case xlArrangeStyleCascade: XlArrangeStyleToString = "xlArrangeStyleCascade"
        Case xlArrangeStyleVertical: XlArrangeStyleToString = "xlArrangeStyleVertical"
        Case xlArrangeStyleHorizontal: XlArrangeStyleToString = "xlArrangeStyleHorizontal"
    End Select
End Function
