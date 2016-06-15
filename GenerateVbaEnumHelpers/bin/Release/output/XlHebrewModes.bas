Attribute VB_Name = "wXlHebrewModes"
Function XlHebrewModesFromString(value As String) As XlHebrewModes
    If IsNumeric(value) Then
        XlHebrewModesFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlHebrewFullScript": XlHebrewModesFromString = xlHebrewFullScript
        Case "xlHebrewPartialScript": XlHebrewModesFromString = xlHebrewPartialScript
        Case "xlHebrewMixedScript": XlHebrewModesFromString = xlHebrewMixedScript
        Case "xlHebrewMixedAuthorizedScript": XlHebrewModesFromString = xlHebrewMixedAuthorizedScript
    End Select
End Function

Function XlHebrewModesToString(value As XlHebrewModes) As String
    Select Case value
        Case xlHebrewFullScript: XlHebrewModesToString = "xlHebrewFullScript"
        Case xlHebrewPartialScript: XlHebrewModesToString = "xlHebrewPartialScript"
        Case xlHebrewMixedScript: XlHebrewModesToString = "xlHebrewMixedScript"
        Case xlHebrewMixedAuthorizedScript: XlHebrewModesToString = "xlHebrewMixedAuthorizedScript"
    End Select
End Function
