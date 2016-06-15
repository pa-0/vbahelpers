Attribute VB_Name = "wWdProofreadingErrorType"
Function WdProofreadingErrorTypeFromString(value As String) As WdProofreadingErrorType
    If IsNumeric(value) Then
        WdProofreadingErrorTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdSpellingError": WdProofreadingErrorTypeFromString = wdSpellingError
        Case "wdGrammaticalError": WdProofreadingErrorTypeFromString = wdGrammaticalError
    End Select
End Function

Function WdProofreadingErrorTypeToString(value As WdProofreadingErrorType) As String
    Select Case value
        Case wdSpellingError: WdProofreadingErrorTypeToString = "wdSpellingError"
        Case wdGrammaticalError: WdProofreadingErrorTypeToString = "wdGrammaticalError"
    End Select
End Function
