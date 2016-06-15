Attribute VB_Name = "wXlThemeFont"
Function XlThemeFontFromString(value As String) As XlThemeFont
    If IsNumeric(value) Then
        XlThemeFontFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlThemeFontNone": XlThemeFontFromString = xlThemeFontNone
        Case "xlThemeFontMajor": XlThemeFontFromString = xlThemeFontMajor
        Case "xlThemeFontMinor": XlThemeFontFromString = xlThemeFontMinor
    End Select
End Function

Function XlThemeFontToString(value As XlThemeFont) As String
    Select Case value
        Case xlThemeFontNone: XlThemeFontToString = "xlThemeFontNone"
        Case xlThemeFontMajor: XlThemeFontToString = "xlThemeFontMajor"
        Case xlThemeFontMinor: XlThemeFontToString = "xlThemeFontMinor"
    End Select
End Function
