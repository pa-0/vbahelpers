Attribute VB_Name = "wWdAraSpeller"
Function WdAraSpellerFromString(value As String) As WdAraSpeller
    If IsNumeric(value) Then
        WdAraSpellerFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdNone": WdAraSpellerFromString = wdNone
        Case "wdInitialAlef": WdAraSpellerFromString = wdInitialAlef
        Case "wdFinalYaa": WdAraSpellerFromString = wdFinalYaa
        Case "wdBoth": WdAraSpellerFromString = wdBoth
    End Select
End Function

Function WdAraSpellerToString(value As WdAraSpeller) As String
    Select Case value
        Case wdNone: WdAraSpellerToString = "wdNone"
        Case wdInitialAlef: WdAraSpellerToString = "wdInitialAlef"
        Case wdFinalYaa: WdAraSpellerToString = "wdFinalYaa"
        Case wdBoth: WdAraSpellerToString = "wdBoth"
    End Select
End Function
