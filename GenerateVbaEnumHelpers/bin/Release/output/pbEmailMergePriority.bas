Attribute VB_Name = "wpbEmailMergePriority"
Function pbEmailMergePriorityFromString(value As String) As pbEmailMergePriority
    If IsNumeric(value) Then
        pbEmailMergePriorityFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "pbPriorityNone": pbEmailMergePriorityFromString = pbPriorityNone
        Case "pbPriorityHigh": pbEmailMergePriorityFromString = pbPriorityHigh
        Case "pbPriorityLow": pbEmailMergePriorityFromString = pbPriorityLow
    End Select
End Function

Function pbEmailMergePriorityToString(value As pbEmailMergePriority) As String
    Select Case value
        Case pbPriorityNone: pbEmailMergePriorityToString = "pbPriorityNone"
        Case pbPriorityHigh: pbEmailMergePriorityToString = "pbPriorityHigh"
        Case pbPriorityLow: pbEmailMergePriorityToString = "pbPriorityLow"
    End Select
End Function
