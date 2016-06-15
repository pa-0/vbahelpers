Attribute VB_Name = "wWdDocumentViewDirection"
Function WdDocumentViewDirectionFromString(value As String) As WdDocumentViewDirection
    If IsNumeric(value) Then
        WdDocumentViewDirectionFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdDocumentViewRtl": WdDocumentViewDirectionFromString = wdDocumentViewRtl
        Case "wdDocumentViewLtr": WdDocumentViewDirectionFromString = wdDocumentViewLtr
    End Select
End Function

Function WdDocumentViewDirectionToString(value As WdDocumentViewDirection) As String
    Select Case value
        Case wdDocumentViewRtl: WdDocumentViewDirectionToString = "wdDocumentViewRtl"
        Case wdDocumentViewLtr: WdDocumentViewDirectionToString = "wdDocumentViewLtr"
    End Select
End Function
