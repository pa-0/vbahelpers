Attribute VB_Name = "wMailFormat"
Function MailFormatFromString(value As String) As MailFormat
    If IsNumeric(value) Then
        MailFormatFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "mfPlainText": MailFormatFromString = mfPlainText
        Case "mfHTML": MailFormatFromString = mfHTML
        Case "mfRTF": MailFormatFromString = mfRTF
    End Select
End Function

Function MailFormatToString(value As MailFormat) As String
    Select Case value
        Case mfPlainText: MailFormatToString = "mfPlainText"
        Case mfHTML: MailFormatToString = "mfHTML"
        Case mfRTF: MailFormatToString = "mfRTF"
    End Select
End Function
