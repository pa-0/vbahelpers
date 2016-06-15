Attribute VB_Name = "wWdKana"
Function WdKanaFromString(value As String) As WdKana
    If IsNumeric(value) Then
        WdKanaFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdKanaKatakana": WdKanaFromString = wdKanaKatakana
        Case "wdKanaHiragana": WdKanaFromString = wdKanaHiragana
    End Select
End Function

Function WdKanaToString(value As WdKana) As String
    Select Case value
        Case wdKanaKatakana: WdKanaToString = "wdKanaKatakana"
        Case wdKanaHiragana: WdKanaToString = "wdKanaHiragana"
    End Select
End Function
