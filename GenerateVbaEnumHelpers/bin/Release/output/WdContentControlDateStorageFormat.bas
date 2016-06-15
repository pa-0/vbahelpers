Attribute VB_Name = "wWdContentControlDateStorageFormat"
Function WdContentControlDateStorageFormatFromString(value As String) As WdContentControlDateStorageFormat
    If IsNumeric(value) Then
        WdContentControlDateStorageFormatFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdContentControlDateStorageText": WdContentControlDateStorageFormatFromString = wdContentControlDateStorageText
        Case "wdContentControlDateStorageDate": WdContentControlDateStorageFormatFromString = wdContentControlDateStorageDate
        Case "wdContentControlDateStorageDateTime": WdContentControlDateStorageFormatFromString = wdContentControlDateStorageDateTime
    End Select
End Function

Function WdContentControlDateStorageFormatToString(value As WdContentControlDateStorageFormat) As String
    Select Case value
        Case wdContentControlDateStorageText: WdContentControlDateStorageFormatToString = "wdContentControlDateStorageText"
        Case wdContentControlDateStorageDate: WdContentControlDateStorageFormatToString = "wdContentControlDateStorageDate"
        Case wdContentControlDateStorageDateTime: WdContentControlDateStorageFormatToString = "wdContentControlDateStorageDateTime"
    End Select
End Function
