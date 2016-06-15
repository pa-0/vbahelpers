Attribute VB_Name = "wWdFarEastLineBreakLevel"
Function WdFarEastLineBreakLevelFromString(value As String) As WdFarEastLineBreakLevel
    If IsNumeric(value) Then
        WdFarEastLineBreakLevelFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdFarEastLineBreakLevelNormal": WdFarEastLineBreakLevelFromString = wdFarEastLineBreakLevelNormal
        Case "wdFarEastLineBreakLevelStrict": WdFarEastLineBreakLevelFromString = wdFarEastLineBreakLevelStrict
        Case "wdFarEastLineBreakLevelCustom": WdFarEastLineBreakLevelFromString = wdFarEastLineBreakLevelCustom
    End Select
End Function

Function WdFarEastLineBreakLevelToString(value As WdFarEastLineBreakLevel) As String
    Select Case value
        Case wdFarEastLineBreakLevelNormal: WdFarEastLineBreakLevelToString = "wdFarEastLineBreakLevelNormal"
        Case wdFarEastLineBreakLevelStrict: WdFarEastLineBreakLevelToString = "wdFarEastLineBreakLevelStrict"
        Case wdFarEastLineBreakLevelCustom: WdFarEastLineBreakLevelToString = "wdFarEastLineBreakLevelCustom"
    End Select
End Function
