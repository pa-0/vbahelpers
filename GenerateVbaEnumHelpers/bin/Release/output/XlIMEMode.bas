Attribute VB_Name = "wXlIMEMode"
Function XlIMEModeFromString(value As String) As XlIMEMode
    If IsNumeric(value) Then
        XlIMEModeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlIMEModeNoControl": XlIMEModeFromString = xlIMEModeNoControl
        Case "xlIMEModeOn": XlIMEModeFromString = xlIMEModeOn
        Case "xlIMEModeOff": XlIMEModeFromString = xlIMEModeOff
        Case "xlIMEModeDisable": XlIMEModeFromString = xlIMEModeDisable
        Case "xlIMEModeHiragana": XlIMEModeFromString = xlIMEModeHiragana
        Case "xlIMEModeKatakana": XlIMEModeFromString = xlIMEModeKatakana
        Case "xlIMEModeKatakanaHalf": XlIMEModeFromString = xlIMEModeKatakanaHalf
        Case "xlIMEModeAlphaFull": XlIMEModeFromString = xlIMEModeAlphaFull
        Case "xlIMEModeAlpha": XlIMEModeFromString = xlIMEModeAlpha
        Case "xlIMEModeHangulFull": XlIMEModeFromString = xlIMEModeHangulFull
        Case "xlIMEModeHangul": XlIMEModeFromString = xlIMEModeHangul
    End Select
End Function

Function XlIMEModeToString(value As XlIMEMode) As String
    Select Case value
        Case xlIMEModeNoControl: XlIMEModeToString = "xlIMEModeNoControl"
        Case xlIMEModeOn: XlIMEModeToString = "xlIMEModeOn"
        Case xlIMEModeOff: XlIMEModeToString = "xlIMEModeOff"
        Case xlIMEModeDisable: XlIMEModeToString = "xlIMEModeDisable"
        Case xlIMEModeHiragana: XlIMEModeToString = "xlIMEModeHiragana"
        Case xlIMEModeKatakana: XlIMEModeToString = "xlIMEModeKatakana"
        Case xlIMEModeKatakanaHalf: XlIMEModeToString = "xlIMEModeKatakanaHalf"
        Case xlIMEModeAlphaFull: XlIMEModeToString = "xlIMEModeAlphaFull"
        Case xlIMEModeAlpha: XlIMEModeToString = "xlIMEModeAlpha"
        Case xlIMEModeHangulFull: XlIMEModeToString = "xlIMEModeHangulFull"
        Case xlIMEModeHangul: XlIMEModeToString = "xlIMEModeHangul"
    End Select
End Function
