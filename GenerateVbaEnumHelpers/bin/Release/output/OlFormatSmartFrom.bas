Attribute VB_Name = "wOlFormatSmartFrom"
Function OlFormatSmartFromFromString(value As String) As OlFormatSmartFrom
    If IsNumeric(value) Then
        OlFormatSmartFromFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "olFormatSmartFromFromTo": OlFormatSmartFromFromString = olFormatSmartFromFromTo
        Case "olFormatSmartFromFromOnly": OlFormatSmartFromFromString = olFormatSmartFromFromOnly
    End Select
End Function

Function OlFormatSmartFromToString(value As OlFormatSmartFrom) As String
    Select Case value
        Case olFormatSmartFromFromTo: OlFormatSmartFromToString = "olFormatSmartFromFromTo"
        Case olFormatSmartFromFromOnly: OlFormatSmartFromToString = "olFormatSmartFromFromOnly"
    End Select
End Function
