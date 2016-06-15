Attribute VB_Name = "wWdFontBias"
Function WdFontBiasFromString(value As String) As WdFontBias
    If IsNumeric(value) Then
        WdFontBiasFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdFontBiasDefault": WdFontBiasFromString = wdFontBiasDefault
        Case "wdFontBiasFareast": WdFontBiasFromString = wdFontBiasFareast
        Case "wdFontBiasDontCare": WdFontBiasFromString = wdFontBiasDontCare
    End Select
End Function

Function WdFontBiasToString(value As WdFontBias) As String
    Select Case value
        Case wdFontBiasDefault: WdFontBiasToString = "wdFontBiasDefault"
        Case wdFontBiasFareast: WdFontBiasToString = "wdFontBiasFareast"
        Case wdFontBiasDontCare: WdFontBiasToString = "wdFontBiasDontCare"
    End Select
End Function
