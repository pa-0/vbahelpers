Attribute VB_Name = "wMsoGradientColorType"
Function MsoGradientColorTypeFromString(value As String) As MsoGradientColorType
    If IsNumeric(value) Then
        MsoGradientColorTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoGradientOneColor": MsoGradientColorTypeFromString = msoGradientOneColor
        Case "msoGradientTwoColors": MsoGradientColorTypeFromString = msoGradientTwoColors
        Case "msoGradientPresetColors": MsoGradientColorTypeFromString = msoGradientPresetColors
        Case "msoGradientMultiColor": MsoGradientColorTypeFromString = msoGradientMultiColor
        Case "msoGradientColorMixed": MsoGradientColorTypeFromString = msoGradientColorMixed
    End Select
End Function

Function MsoGradientColorTypeToString(value As MsoGradientColorType) As String
    Select Case value
        Case msoGradientOneColor: MsoGradientColorTypeToString = "msoGradientOneColor"
        Case msoGradientTwoColors: MsoGradientColorTypeToString = "msoGradientTwoColors"
        Case msoGradientPresetColors: MsoGradientColorTypeToString = "msoGradientPresetColors"
        Case msoGradientMultiColor: MsoGradientColorTypeToString = "msoGradientMultiColor"
        Case msoGradientColorMixed: MsoGradientColorTypeToString = "msoGradientColorMixed"
    End Select
End Function
