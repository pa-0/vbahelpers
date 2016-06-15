Attribute VB_Name = "wWdFrenchSpeller"
Function WdFrenchSpellerFromString(value As String) As WdFrenchSpeller
    If IsNumeric(value) Then
        WdFrenchSpellerFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdFrenchBoth": WdFrenchSpellerFromString = wdFrenchBoth
        Case "wdFrenchPreReform": WdFrenchSpellerFromString = wdFrenchPreReform
        Case "wdFrenchPostReform": WdFrenchSpellerFromString = wdFrenchPostReform
    End Select
End Function

Function WdFrenchSpellerToString(value As WdFrenchSpeller) As String
    Select Case value
        Case wdFrenchBoth: WdFrenchSpellerToString = "wdFrenchBoth"
        Case wdFrenchPreReform: WdFrenchSpellerToString = "wdFrenchPreReform"
        Case wdFrenchPostReform: WdFrenchSpellerToString = "wdFrenchPostReform"
    End Select
End Function
