Attribute VB_Name = "wMsoClipboardFormat"
Function MsoClipboardFormatFromString(value As String) As MsoClipboardFormat
    If IsNumeric(value) Then
        MsoClipboardFormatFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoClipboardFormatNative": MsoClipboardFormatFromString = msoClipboardFormatNative
        Case "msoClipboardFormatHTML": MsoClipboardFormatFromString = msoClipboardFormatHTML
        Case "msoClipboardFormatRTF": MsoClipboardFormatFromString = msoClipboardFormatRTF
        Case "msoClipboardFormatPlainText": MsoClipboardFormatFromString = msoClipboardFormatPlainText
        Case "msoClipboardFormatMixed": MsoClipboardFormatFromString = msoClipboardFormatMixed
    End Select
End Function

Function MsoClipboardFormatToString(value As MsoClipboardFormat) As String
    Select Case value
        Case msoClipboardFormatNative: MsoClipboardFormatToString = "msoClipboardFormatNative"
        Case msoClipboardFormatHTML: MsoClipboardFormatToString = "msoClipboardFormatHTML"
        Case msoClipboardFormatRTF: MsoClipboardFormatToString = "msoClipboardFormatRTF"
        Case msoClipboardFormatPlainText: MsoClipboardFormatToString = "msoClipboardFormatPlainText"
        Case msoClipboardFormatMixed: MsoClipboardFormatToString = "msoClipboardFormatMixed"
    End Select
End Function
