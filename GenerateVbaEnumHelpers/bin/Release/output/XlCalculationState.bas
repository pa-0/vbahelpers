Attribute VB_Name = "wXlCalculationState"
Function XlCalculationStateFromString(value As String) As XlCalculationState
    If IsNumeric(value) Then
        XlCalculationStateFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlDone": XlCalculationStateFromString = xlDone
        Case "xlCalculating": XlCalculationStateFromString = xlCalculating
        Case "xlPending": XlCalculationStateFromString = xlPending
    End Select
End Function

Function XlCalculationStateToString(value As XlCalculationState) As String
    Select Case value
        Case xlDone: XlCalculationStateToString = "xlDone"
        Case xlCalculating: XlCalculationStateToString = "xlCalculating"
        Case xlPending: XlCalculationStateToString = "xlPending"
    End Select
End Function
