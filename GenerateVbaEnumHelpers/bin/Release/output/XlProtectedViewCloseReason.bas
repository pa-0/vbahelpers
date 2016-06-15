Attribute VB_Name = "wXlProtectedViewCloseReason"
Function XlProtectedViewCloseReasonFromString(value As String) As XlProtectedViewCloseReason
    If IsNumeric(value) Then
        XlProtectedViewCloseReasonFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlProtectedViewCloseNormal": XlProtectedViewCloseReasonFromString = xlProtectedViewCloseNormal
        Case "xlProtectedViewCloseEdit": XlProtectedViewCloseReasonFromString = xlProtectedViewCloseEdit
        Case "xlProtectedViewCloseForced": XlProtectedViewCloseReasonFromString = xlProtectedViewCloseForced
    End Select
End Function

Function XlProtectedViewCloseReasonToString(value As XlProtectedViewCloseReason) As String
    Select Case value
        Case xlProtectedViewCloseNormal: XlProtectedViewCloseReasonToString = "xlProtectedViewCloseNormal"
        Case xlProtectedViewCloseEdit: XlProtectedViewCloseReasonToString = "xlProtectedViewCloseEdit"
        Case xlProtectedViewCloseForced: XlProtectedViewCloseReasonToString = "xlProtectedViewCloseForced"
    End Select
End Function
