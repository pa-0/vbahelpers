Attribute VB_Name = "wWdMailMergeMailFormat"
Function WdMailMergeMailFormatFromString(value As String) As WdMailMergeMailFormat
    If IsNumeric(value) Then
        WdMailMergeMailFormatFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdMailFormatPlainText": WdMailMergeMailFormatFromString = wdMailFormatPlainText
        Case "wdMailFormatHTML": WdMailMergeMailFormatFromString = wdMailFormatHTML
    End Select
End Function

Function WdMailMergeMailFormatToString(value As WdMailMergeMailFormat) As String
    Select Case value
        Case wdMailFormatPlainText: WdMailMergeMailFormatToString = "wdMailFormatPlainText"
        Case wdMailFormatHTML: WdMailMergeMailFormatToString = "wdMailFormatHTML"
    End Select
End Function
