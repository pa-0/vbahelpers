Attribute VB_Name = "wXlChartPictureType"
Function XlChartPictureTypeFromString(value As String) As XlChartPictureType
    If IsNumeric(value) Then
        XlChartPictureTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlStretch": XlChartPictureTypeFromString = xlStretch
        Case "xlStack": XlChartPictureTypeFromString = xlStack
        Case "xlStackScale": XlChartPictureTypeFromString = xlStackScale
    End Select
End Function

Function XlChartPictureTypeToString(value As XlChartPictureType) As String
    Select Case value
        Case xlStretch: XlChartPictureTypeToString = "xlStretch"
        Case xlStack: XlChartPictureTypeToString = "xlStack"
        Case xlStackScale: XlChartPictureTypeToString = "xlStackScale"
    End Select
End Function
