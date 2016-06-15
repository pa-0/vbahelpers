Attribute VB_Name = "wWdExportFormat"
Function WdExportFormatFromString(value As String) As WdExportFormat
    If IsNumeric(value) Then
        WdExportFormatFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdExportFormatPDF": WdExportFormatFromString = wdExportFormatPDF
        Case "wdExportFormatXPS": WdExportFormatFromString = wdExportFormatXPS
    End Select
End Function

Function WdExportFormatToString(value As WdExportFormat) As String
    Select Case value
        Case wdExportFormatPDF: WdExportFormatToString = "wdExportFormatPDF"
        Case wdExportFormatXPS: WdExportFormatToString = "wdExportFormatXPS"
    End Select
End Function
