Attribute VB_Name = "wOlMobileFormat"
Function OlMobileFormatFromString(value As String) As OlMobileFormat
    If IsNumeric(value) Then
        OlMobileFormatFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "olSMS": OlMobileFormatFromString = olSMS
        Case "olMMS": OlMobileFormatFromString = olMMS
    End Select
End Function

Function OlMobileFormatToString(value As OlMobileFormat) As String
    Select Case value
        Case olSMS: OlMobileFormatToString = "olSMS"
        Case olMMS: OlMobileFormatToString = "olMMS"
    End Select
End Function
