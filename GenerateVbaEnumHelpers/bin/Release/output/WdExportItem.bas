Attribute VB_Name = "wWdExportItem"
Function WdExportItemFromString(value As String) As WdExportItem
    If IsNumeric(value) Then
        WdExportItemFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdExportDocumentContent": WdExportItemFromString = wdExportDocumentContent
        Case "wdExportDocumentWithMarkup": WdExportItemFromString = wdExportDocumentWithMarkup
    End Select
End Function

Function WdExportItemToString(value As WdExportItem) As String
    Select Case value
        Case wdExportDocumentContent: WdExportItemToString = "wdExportDocumentContent"
        Case wdExportDocumentWithMarkup: WdExportItemToString = "wdExportDocumentWithMarkup"
    End Select
End Function
