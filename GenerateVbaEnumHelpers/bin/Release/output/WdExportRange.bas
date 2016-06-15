Attribute VB_Name = "wWdExportRange"
Function WdExportRangeFromString(value As String) As WdExportRange
    If IsNumeric(value) Then
        WdExportRangeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdExportAllDocument": WdExportRangeFromString = wdExportAllDocument
        Case "wdExportSelection": WdExportRangeFromString = wdExportSelection
        Case "wdExportCurrentPage": WdExportRangeFromString = wdExportCurrentPage
        Case "wdExportFromTo": WdExportRangeFromString = wdExportFromTo
    End Select
End Function

Function WdExportRangeToString(value As WdExportRange) As String
    Select Case value
        Case wdExportAllDocument: WdExportRangeToString = "wdExportAllDocument"
        Case wdExportSelection: WdExportRangeToString = "wdExportSelection"
        Case wdExportCurrentPage: WdExportRangeToString = "wdExportCurrentPage"
        Case wdExportFromTo: WdExportRangeToString = "wdExportFromTo"
    End Select
End Function
