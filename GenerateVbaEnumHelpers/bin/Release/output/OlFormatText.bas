Attribute VB_Name = "wOlFormatText"
Function OlFormatTextFromString(value As String) As OlFormatText
    If IsNumeric(value) Then
        OlFormatTextFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "olFormatTextText": OlFormatTextFromString = olFormatTextText
    End Select
End Function

Function OlFormatTextToString(value As OlFormatText) As String
    Select Case value
        Case olFormatTextText: OlFormatTextToString = "olFormatTextText"
    End Select
End Function
