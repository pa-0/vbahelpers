Attribute VB_Name = "wWdIMEMode"
Function WdIMEModeFromString(value As String) As WdIMEMode
    If IsNumeric(value) Then
        WdIMEModeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdIMEModeNoControl": WdIMEModeFromString = wdIMEModeNoControl
        Case "wdIMEModeOn": WdIMEModeFromString = wdIMEModeOn
        Case "wdIMEModeOff": WdIMEModeFromString = wdIMEModeOff
        Case "wdIMEModeHiragana": WdIMEModeFromString = wdIMEModeHiragana
        Case "wdIMEModeKatakana": WdIMEModeFromString = wdIMEModeKatakana
        Case "wdIMEModeKatakanaHalf": WdIMEModeFromString = wdIMEModeKatakanaHalf
        Case "wdIMEModeAlphaFull": WdIMEModeFromString = wdIMEModeAlphaFull
        Case "wdIMEModeAlpha": WdIMEModeFromString = wdIMEModeAlpha
        Case "wdIMEModeHangulFull": WdIMEModeFromString = wdIMEModeHangulFull
        Case "wdIMEModeHangul": WdIMEModeFromString = wdIMEModeHangul
    End Select
End Function

Function WdIMEModeToString(value As WdIMEMode) As String
    Select Case value
        Case wdIMEModeNoControl: WdIMEModeToString = "wdIMEModeNoControl"
        Case wdIMEModeOn: WdIMEModeToString = "wdIMEModeOn"
        Case wdIMEModeOff: WdIMEModeToString = "wdIMEModeOff"
        Case wdIMEModeHiragana: WdIMEModeToString = "wdIMEModeHiragana"
        Case wdIMEModeKatakana: WdIMEModeToString = "wdIMEModeKatakana"
        Case wdIMEModeKatakanaHalf: WdIMEModeToString = "wdIMEModeKatakanaHalf"
        Case wdIMEModeAlphaFull: WdIMEModeToString = "wdIMEModeAlphaFull"
        Case wdIMEModeAlpha: WdIMEModeToString = "wdIMEModeAlpha"
        Case wdIMEModeHangulFull: WdIMEModeToString = "wdIMEModeHangulFull"
        Case wdIMEModeHangul: WdIMEModeToString = "wdIMEModeHangul"
    End Select
End Function
