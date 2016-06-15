Attribute VB_Name = "wWdDocumentKind"
Function WdDocumentKindFromString(value As String) As WdDocumentKind
    If IsNumeric(value) Then
        WdDocumentKindFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdDocumentNotSpecified": WdDocumentKindFromString = wdDocumentNotSpecified
        Case "wdDocumentLetter": WdDocumentKindFromString = wdDocumentLetter
        Case "wdDocumentEmail": WdDocumentKindFromString = wdDocumentEmail
    End Select
End Function

Function WdDocumentKindToString(value As WdDocumentKind) As String
    Select Case value
        Case wdDocumentNotSpecified: WdDocumentKindToString = "wdDocumentNotSpecified"
        Case wdDocumentLetter: WdDocumentKindToString = "wdDocumentLetter"
        Case wdDocumentEmail: WdDocumentKindToString = "wdDocumentEmail"
    End Select
End Function
