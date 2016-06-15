Attribute VB_Name = "wWdDocumentMedium"
Function WdDocumentMediumFromString(value As String) As WdDocumentMedium
    If IsNumeric(value) Then
        WdDocumentMediumFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdEmailMessage": WdDocumentMediumFromString = wdEmailMessage
        Case "wdDocument": WdDocumentMediumFromString = wdDocument
        Case "wdWebPage": WdDocumentMediumFromString = wdWebPage
    End Select
End Function

Function WdDocumentMediumToString(value As WdDocumentMedium) As String
    Select Case value
        Case wdEmailMessage: WdDocumentMediumToString = "wdEmailMessage"
        Case wdDocument: WdDocumentMediumToString = "wdDocument"
        Case wdWebPage: WdDocumentMediumToString = "wdWebPage"
    End Select
End Function
