Attribute VB_Name = "wWdDocumentDirection"
Function WdDocumentDirectionFromString(value As String) As WdDocumentDirection
    If IsNumeric(value) Then
        WdDocumentDirectionFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdLeftToRight": WdDocumentDirectionFromString = wdLeftToRight
        Case "wdRightToLeft": WdDocumentDirectionFromString = wdRightToLeft
    End Select
End Function

Function WdDocumentDirectionToString(value As WdDocumentDirection) As String
    Select Case value
        Case wdLeftToRight: WdDocumentDirectionToString = "wdLeftToRight"
        Case wdRightToLeft: WdDocumentDirectionToString = "wdRightToLeft"
    End Select
End Function
