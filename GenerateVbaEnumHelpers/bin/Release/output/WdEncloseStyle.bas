Attribute VB_Name = "wWdEncloseStyle"
Function WdEncloseStyleFromString(value As String) As WdEncloseStyle
    If IsNumeric(value) Then
        WdEncloseStyleFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdEncloseStyleNone": WdEncloseStyleFromString = wdEncloseStyleNone
        Case "wdEncloseStyleSmall": WdEncloseStyleFromString = wdEncloseStyleSmall
        Case "wdEncloseStyleLarge": WdEncloseStyleFromString = wdEncloseStyleLarge
    End Select
End Function

Function WdEncloseStyleToString(value As WdEncloseStyle) As String
    Select Case value
        Case wdEncloseStyleNone: WdEncloseStyleToString = "wdEncloseStyleNone"
        Case wdEncloseStyleSmall: WdEncloseStyleToString = "wdEncloseStyleSmall"
        Case wdEncloseStyleLarge: WdEncloseStyleToString = "wdEncloseStyleLarge"
    End Select
End Function
