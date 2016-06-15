Attribute VB_Name = "wWdOriginalFormat"
Function WdOriginalFormatFromString(value As String) As WdOriginalFormat
    If IsNumeric(value) Then
        WdOriginalFormatFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdWordDocument": WdOriginalFormatFromString = wdWordDocument
        Case "wdOriginalDocumentFormat": WdOriginalFormatFromString = wdOriginalDocumentFormat
        Case "wdPromptUser": WdOriginalFormatFromString = wdPromptUser
    End Select
End Function

Function WdOriginalFormatToString(value As WdOriginalFormat) As String
    Select Case value
        Case wdWordDocument: WdOriginalFormatToString = "wdWordDocument"
        Case wdOriginalDocumentFormat: WdOriginalFormatToString = "wdOriginalDocumentFormat"
        Case wdPromptUser: WdOriginalFormatToString = "wdPromptUser"
    End Select
End Function
