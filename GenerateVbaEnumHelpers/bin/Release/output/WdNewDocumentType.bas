Attribute VB_Name = "wWdNewDocumentType"
Function WdNewDocumentTypeFromString(value As String) As WdNewDocumentType
    If IsNumeric(value) Then
        WdNewDocumentTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdNewBlankDocument": WdNewDocumentTypeFromString = wdNewBlankDocument
        Case "wdNewWebPage": WdNewDocumentTypeFromString = wdNewWebPage
        Case "wdNewEmailMessage": WdNewDocumentTypeFromString = wdNewEmailMessage
        Case "wdNewFrameset": WdNewDocumentTypeFromString = wdNewFrameset
        Case "wdNewXMLDocument": WdNewDocumentTypeFromString = wdNewXMLDocument
    End Select
End Function

Function WdNewDocumentTypeToString(value As WdNewDocumentType) As String
    Select Case value
        Case wdNewBlankDocument: WdNewDocumentTypeToString = "wdNewBlankDocument"
        Case wdNewWebPage: WdNewDocumentTypeToString = "wdNewWebPage"
        Case wdNewEmailMessage: WdNewDocumentTypeToString = "wdNewEmailMessage"
        Case wdNewFrameset: WdNewDocumentTypeToString = "wdNewFrameset"
        Case wdNewXMLDocument: WdNewDocumentTypeToString = "wdNewXMLDocument"
    End Select
End Function
