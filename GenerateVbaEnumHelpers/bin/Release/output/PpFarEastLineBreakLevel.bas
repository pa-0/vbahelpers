Attribute VB_Name = "wPpFarEastLineBreakLevel"
Function PpFarEastLineBreakLevelFromString(value As String) As PpFarEastLineBreakLevel
    If IsNumeric(value) Then
        PpFarEastLineBreakLevelFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "ppFarEastLineBreakLevelNormal": PpFarEastLineBreakLevelFromString = ppFarEastLineBreakLevelNormal
        Case "ppFarEastLineBreakLevelStrict": PpFarEastLineBreakLevelFromString = ppFarEastLineBreakLevelStrict
        Case "ppFarEastLineBreakLevelCustom": PpFarEastLineBreakLevelFromString = ppFarEastLineBreakLevelCustom
    End Select
End Function

Function PpFarEastLineBreakLevelToString(value As PpFarEastLineBreakLevel) As String
    Select Case value
        Case ppFarEastLineBreakLevelNormal: PpFarEastLineBreakLevelToString = "ppFarEastLineBreakLevelNormal"
        Case ppFarEastLineBreakLevelStrict: PpFarEastLineBreakLevelToString = "ppFarEastLineBreakLevelStrict"
        Case ppFarEastLineBreakLevelCustom: PpFarEastLineBreakLevelToString = "ppFarEastLineBreakLevelCustom"
    End Select
End Function
