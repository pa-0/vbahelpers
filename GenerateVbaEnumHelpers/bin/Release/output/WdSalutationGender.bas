Attribute VB_Name = "wWdSalutationGender"
Function WdSalutationGenderFromString(value As String) As WdSalutationGender
    If IsNumeric(value) Then
        WdSalutationGenderFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdGenderFemale": WdSalutationGenderFromString = wdGenderFemale
        Case "wdGenderMale": WdSalutationGenderFromString = wdGenderMale
        Case "wdGenderNeutral": WdSalutationGenderFromString = wdGenderNeutral
        Case "wdGenderUnknown": WdSalutationGenderFromString = wdGenderUnknown
    End Select
End Function

Function WdSalutationGenderToString(value As WdSalutationGender) As String
    Select Case value
        Case wdGenderFemale: WdSalutationGenderToString = "wdGenderFemale"
        Case wdGenderMale: WdSalutationGenderToString = "wdGenderMale"
        Case wdGenderNeutral: WdSalutationGenderToString = "wdGenderNeutral"
        Case wdGenderUnknown: WdSalutationGenderToString = "wdGenderUnknown"
    End Select
End Function
