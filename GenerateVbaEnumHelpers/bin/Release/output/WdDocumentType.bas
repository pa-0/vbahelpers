Attribute VB_Name = "wWdDocumentType"
Function WdDocumentTypeFromString(value As String) As WdDocumentType
    If IsNumeric(value) Then
        WdDocumentTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdTypeDocument": WdDocumentTypeFromString = wdTypeDocument
        Case "wdTypeTemplate": WdDocumentTypeFromString = wdTypeTemplate
        Case "wdTypeFrameset": WdDocumentTypeFromString = wdTypeFrameset
    End Select
End Function

Function WdDocumentTypeToString(value As WdDocumentType) As String
    Select Case value
        Case wdTypeDocument: WdDocumentTypeToString = "wdTypeDocument"
        Case wdTypeTemplate: WdDocumentTypeToString = "wdTypeTemplate"
        Case wdTypeFrameset: WdDocumentTypeToString = "wdTypeFrameset"
    End Select
End Function
