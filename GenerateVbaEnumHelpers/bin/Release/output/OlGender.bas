Attribute VB_Name = "wOlGender"
Function OlGenderFromString(value As String) As OlGender
    If IsNumeric(value) Then
        OlGenderFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "olUnspecified": OlGenderFromString = olUnspecified
        Case "olFemale": OlGenderFromString = olFemale
        Case "olMale": OlGenderFromString = olMale
    End Select
End Function

Function OlGenderToString(value As OlGender) As String
    Select Case value
        Case olUnspecified: OlGenderToString = "olUnspecified"
        Case olFemale: OlGenderToString = "olFemale"
        Case olMale: OlGenderToString = "olMale"
    End Select
End Function
