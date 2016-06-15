Attribute VB_Name = "wOlFormatKeywords"
Function OlFormatKeywordsFromString(value As String) As OlFormatKeywords
    If IsNumeric(value) Then
        OlFormatKeywordsFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "olFormatKeywordsText": OlFormatKeywordsFromString = olFormatKeywordsText
    End Select
End Function

Function OlFormatKeywordsToString(value As OlFormatKeywords) As String
    Select Case value
        Case olFormatKeywordsText: OlFormatKeywordsToString = "olFormatKeywordsText"
    End Select
End Function
