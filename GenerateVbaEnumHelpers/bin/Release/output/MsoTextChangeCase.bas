Attribute VB_Name = "wMsoTextChangeCase"
Function MsoTextChangeCaseFromString(value As String) As MsoTextChangeCase
    If IsNumeric(value) Then
        MsoTextChangeCaseFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoCaseSentence": MsoTextChangeCaseFromString = msoCaseSentence
        Case "msoCaseLower": MsoTextChangeCaseFromString = msoCaseLower
        Case "msoCaseUpper": MsoTextChangeCaseFromString = msoCaseUpper
        Case "msoCaseTitle": MsoTextChangeCaseFromString = msoCaseTitle
        Case "msoCaseToggle": MsoTextChangeCaseFromString = msoCaseToggle
    End Select
End Function

Function MsoTextChangeCaseToString(value As MsoTextChangeCase) As String
    Select Case value
        Case msoCaseSentence: MsoTextChangeCaseToString = "msoCaseSentence"
        Case msoCaseLower: MsoTextChangeCaseToString = "msoCaseLower"
        Case msoCaseUpper: MsoTextChangeCaseToString = "msoCaseUpper"
        Case msoCaseTitle: MsoTextChangeCaseToString = "msoCaseTitle"
        Case msoCaseToggle: MsoTextChangeCaseToString = "msoCaseToggle"
    End Select
End Function
