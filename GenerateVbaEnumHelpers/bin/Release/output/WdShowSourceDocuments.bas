Attribute VB_Name = "wWdShowSourceDocuments"
Function WdShowSourceDocumentsFromString(value As String) As WdShowSourceDocuments
    If IsNumeric(value) Then
        WdShowSourceDocumentsFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdShowSourceDocumentsNone": WdShowSourceDocumentsFromString = wdShowSourceDocumentsNone
        Case "wdShowSourceDocumentsOriginal": WdShowSourceDocumentsFromString = wdShowSourceDocumentsOriginal
        Case "wdShowSourceDocumentsRevised": WdShowSourceDocumentsFromString = wdShowSourceDocumentsRevised
        Case "wdShowSourceDocumentsBoth": WdShowSourceDocumentsFromString = wdShowSourceDocumentsBoth
    End Select
End Function

Function WdShowSourceDocumentsToString(value As WdShowSourceDocuments) As String
    Select Case value
        Case wdShowSourceDocumentsNone: WdShowSourceDocumentsToString = "wdShowSourceDocumentsNone"
        Case wdShowSourceDocumentsOriginal: WdShowSourceDocumentsToString = "wdShowSourceDocumentsOriginal"
        Case wdShowSourceDocumentsRevised: WdShowSourceDocumentsToString = "wdShowSourceDocumentsRevised"
        Case wdShowSourceDocumentsBoth: WdShowSourceDocumentsToString = "wdShowSourceDocumentsBoth"
    End Select
End Function
