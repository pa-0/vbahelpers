Attribute VB_Name = "wPpProtectedViewCloseReason"
Function PpProtectedViewCloseReasonFromString(value As String) As PpProtectedViewCloseReason
    If IsNumeric(value) Then
        PpProtectedViewCloseReasonFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "ppProtectedViewCloseNormal": PpProtectedViewCloseReasonFromString = ppProtectedViewCloseNormal
        Case "ppProtectedViewCloseEdit": PpProtectedViewCloseReasonFromString = ppProtectedViewCloseEdit
        Case "ppProtectedViewCloseForced": PpProtectedViewCloseReasonFromString = ppProtectedViewCloseForced
    End Select
End Function

Function PpProtectedViewCloseReasonToString(value As PpProtectedViewCloseReason) As String
    Select Case value
        Case ppProtectedViewCloseNormal: PpProtectedViewCloseReasonToString = "ppProtectedViewCloseNormal"
        Case ppProtectedViewCloseEdit: PpProtectedViewCloseReasonToString = "ppProtectedViewCloseEdit"
        Case ppProtectedViewCloseForced: PpProtectedViewCloseReasonToString = "ppProtectedViewCloseForced"
    End Select
End Function
