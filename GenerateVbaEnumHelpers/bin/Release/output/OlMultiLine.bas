Attribute VB_Name = "wOlMultiLine"
Function OlMultiLineFromString(value As String) As OlMultiLine
    If IsNumeric(value) Then
        OlMultiLineFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "olWidthMultiLine": OlMultiLineFromString = olWidthMultiLine
        Case "olAlwaysSingleLine": OlMultiLineFromString = olAlwaysSingleLine
        Case "olAlwaysMultiLine": OlMultiLineFromString = olAlwaysMultiLine
    End Select
End Function

Function OlMultiLineToString(value As OlMultiLine) As String
    Select Case value
        Case olWidthMultiLine: OlMultiLineToString = "olWidthMultiLine"
        Case olAlwaysSingleLine: OlMultiLineToString = "olAlwaysSingleLine"
        Case olAlwaysMultiLine: OlMultiLineToString = "olAlwaysMultiLine"
    End Select
End Function
