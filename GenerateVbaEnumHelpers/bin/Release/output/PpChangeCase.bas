Attribute VB_Name = "wPpChangeCase"
Function PpChangeCaseFromString(value As String) As PpChangeCase
    If IsNumeric(value) Then
        PpChangeCaseFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "ppCaseSentence": PpChangeCaseFromString = ppCaseSentence
        Case "ppCaseLower": PpChangeCaseFromString = ppCaseLower
        Case "ppCaseUpper": PpChangeCaseFromString = ppCaseUpper
        Case "ppCaseTitle": PpChangeCaseFromString = ppCaseTitle
        Case "ppCaseToggle": PpChangeCaseFromString = ppCaseToggle
    End Select
End Function

Function PpChangeCaseToString(value As PpChangeCase) As String
    Select Case value
        Case ppCaseSentence: PpChangeCaseToString = "ppCaseSentence"
        Case ppCaseLower: PpChangeCaseToString = "ppCaseLower"
        Case ppCaseUpper: PpChangeCaseToString = "ppCaseUpper"
        Case ppCaseTitle: PpChangeCaseToString = "ppCaseTitle"
        Case ppCaseToggle: PpChangeCaseToString = "ppCaseToggle"
    End Select
End Function
