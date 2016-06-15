Attribute VB_Name = "wXlSparkScale"
Function XlSparkScaleFromString(value As String) As XlSparkScale
    If IsNumeric(value) Then
        XlSparkScaleFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlSparkScaleGroup": XlSparkScaleFromString = xlSparkScaleGroup
        Case "xlSparkScaleSingle": XlSparkScaleFromString = xlSparkScaleSingle
        Case "xlSparkScaleCustom": XlSparkScaleFromString = xlSparkScaleCustom
    End Select
End Function

Function XlSparkScaleToString(value As XlSparkScale) As String
    Select Case value
        Case xlSparkScaleGroup: XlSparkScaleToString = "xlSparkScaleGroup"
        Case xlSparkScaleSingle: XlSparkScaleToString = "xlSparkScaleSingle"
        Case xlSparkScaleCustom: XlSparkScaleToString = "xlSparkScaleCustom"
    End Select
End Function
