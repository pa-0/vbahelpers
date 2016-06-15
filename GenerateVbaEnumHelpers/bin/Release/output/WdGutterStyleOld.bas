Attribute VB_Name = "wWdGutterStyleOld"
Function WdGutterStyleOldFromString(value As String) As WdGutterStyleOld
    If IsNumeric(value) Then
        WdGutterStyleOldFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdGutterStyleBidi": WdGutterStyleOldFromString = wdGutterStyleBidi
        Case "wdGutterStyleLatin": WdGutterStyleOldFromString = wdGutterStyleLatin
    End Select
End Function

Function WdGutterStyleOldToString(value As WdGutterStyleOld) As String
    Select Case value
        Case wdGutterStyleBidi: WdGutterStyleOldToString = "wdGutterStyleBidi"
        Case wdGutterStyleLatin: WdGutterStyleOldToString = "wdGutterStyleLatin"
    End Select
End Function
