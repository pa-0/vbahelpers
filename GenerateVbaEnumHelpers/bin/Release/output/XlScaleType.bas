Attribute VB_Name = "wXlScaleType"
Function XlScaleTypeFromString(value As String) As XlScaleType
    If IsNumeric(value) Then
        XlScaleTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlScaleLogarithmic": XlScaleTypeFromString = xlScaleLogarithmic
        Case "xlScaleLinear": XlScaleTypeFromString = xlScaleLinear
    End Select
End Function

Function XlScaleTypeToString(value As XlScaleType) As String
    Select Case value
        Case xlScaleLogarithmic: XlScaleTypeToString = "xlScaleLogarithmic"
        Case xlScaleLinear: XlScaleTypeToString = "xlScaleLinear"
    End Select
End Function
