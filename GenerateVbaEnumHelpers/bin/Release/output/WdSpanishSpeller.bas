Attribute VB_Name = "wWdSpanishSpeller"
Function WdSpanishSpellerFromString(value As String) As WdSpanishSpeller
    If IsNumeric(value) Then
        WdSpanishSpellerFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdSpanishTuteoOnly": WdSpanishSpellerFromString = wdSpanishTuteoOnly
        Case "wdSpanishTuteoAndVoseo": WdSpanishSpellerFromString = wdSpanishTuteoAndVoseo
        Case "wdSpanishVoseoOnly": WdSpanishSpellerFromString = wdSpanishVoseoOnly
    End Select
End Function

Function WdSpanishSpellerToString(value As WdSpanishSpeller) As String
    Select Case value
        Case wdSpanishTuteoOnly: WdSpanishSpellerToString = "wdSpanishTuteoOnly"
        Case wdSpanishTuteoAndVoseo: WdSpanishSpellerToString = "wdSpanishTuteoAndVoseo"
        Case wdSpanishVoseoOnly: WdSpanishSpellerToString = "wdSpanishVoseoOnly"
    End Select
End Function
