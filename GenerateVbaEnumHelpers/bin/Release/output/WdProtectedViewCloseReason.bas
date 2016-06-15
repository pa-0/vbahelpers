Attribute VB_Name = "wWdProtectedViewCloseReason"
Function WdProtectedViewCloseReasonFromString(value As String) As WdProtectedViewCloseReason
    If IsNumeric(value) Then
        WdProtectedViewCloseReasonFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdProtectedViewCloseNormal": WdProtectedViewCloseReasonFromString = wdProtectedViewCloseNormal
        Case "wdProtectedViewCloseEdit": WdProtectedViewCloseReasonFromString = wdProtectedViewCloseEdit
        Case "wdProtectedViewCloseForced": WdProtectedViewCloseReasonFromString = wdProtectedViewCloseForced
    End Select
End Function

Function WdProtectedViewCloseReasonToString(value As WdProtectedViewCloseReason) As String
    Select Case value
        Case wdProtectedViewCloseNormal: WdProtectedViewCloseReasonToString = "wdProtectedViewCloseNormal"
        Case wdProtectedViewCloseEdit: WdProtectedViewCloseReasonToString = "wdProtectedViewCloseEdit"
        Case wdProtectedViewCloseForced: WdProtectedViewCloseReasonToString = "wdProtectedViewCloseForced"
    End Select
End Function
