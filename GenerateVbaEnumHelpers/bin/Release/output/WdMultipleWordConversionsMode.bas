Attribute VB_Name = "wWdMultipleWordConversionsMode"
Function WdMultipleWordConversionsModeFromString(value As String) As WdMultipleWordConversionsMode
    If IsNumeric(value) Then
        WdMultipleWordConversionsModeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdHangulToHanja": WdMultipleWordConversionsModeFromString = wdHangulToHanja
        Case "wdHanjaToHangul": WdMultipleWordConversionsModeFromString = wdHanjaToHangul
    End Select
End Function

Function WdMultipleWordConversionsModeToString(value As WdMultipleWordConversionsMode) As String
    Select Case value
        Case wdHangulToHanja: WdMultipleWordConversionsModeToString = "wdHangulToHanja"
        Case wdHanjaToHangul: WdMultipleWordConversionsModeToString = "wdHanjaToHangul"
    End Select
End Function
