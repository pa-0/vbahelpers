Attribute VB_Name = "wOlFormatInteger"
Function OlFormatIntegerFromString(value As String) As OlFormatInteger
    If IsNumeric(value) Then
        OlFormatIntegerFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "olFormatIntegerPlain": OlFormatIntegerFromString = olFormatIntegerPlain
        Case "olFormatIntegerComputer1": OlFormatIntegerFromString = olFormatIntegerComputer1
        Case "olFormatIntegerComputer2": OlFormatIntegerFromString = olFormatIntegerComputer2
        Case "olFormatIntegerComputer3": OlFormatIntegerFromString = olFormatIntegerComputer3
    End Select
End Function

Function OlFormatIntegerToString(value As OlFormatInteger) As String
    Select Case value
        Case olFormatIntegerPlain: OlFormatIntegerToString = "olFormatIntegerPlain"
        Case olFormatIntegerComputer1: OlFormatIntegerToString = "olFormatIntegerComputer1"
        Case olFormatIntegerComputer2: OlFormatIntegerToString = "olFormatIntegerComputer2"
        Case olFormatIntegerComputer3: OlFormatIntegerToString = "olFormatIntegerComputer3"
    End Select
End Function
