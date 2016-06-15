Attribute VB_Name = "wWdSpellingErrorType"
Function WdSpellingErrorTypeFromString(value As String) As WdSpellingErrorType
    If IsNumeric(value) Then
        WdSpellingErrorTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdSpellingCorrect": WdSpellingErrorTypeFromString = wdSpellingCorrect
        Case "wdSpellingNotInDictionary": WdSpellingErrorTypeFromString = wdSpellingNotInDictionary
        Case "wdSpellingCapitalization": WdSpellingErrorTypeFromString = wdSpellingCapitalization
    End Select
End Function

Function WdSpellingErrorTypeToString(value As WdSpellingErrorType) As String
    Select Case value
        Case wdSpellingCorrect: WdSpellingErrorTypeToString = "wdSpellingCorrect"
        Case wdSpellingNotInDictionary: WdSpellingErrorTypeToString = "wdSpellingNotInDictionary"
        Case wdSpellingCapitalization: WdSpellingErrorTypeToString = "wdSpellingCapitalization"
    End Select
End Function
