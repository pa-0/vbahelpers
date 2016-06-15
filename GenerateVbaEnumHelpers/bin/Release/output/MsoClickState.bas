Attribute VB_Name = "wMsoClickState"
Function MsoClickStateFromString(value As String) As MsoClickState
    If IsNumeric(value) Then
        MsoClickStateFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoClickStateAfterAllAnimations": MsoClickStateFromString = msoClickStateAfterAllAnimations
        Case "msoClickStateBeforeAutomaticAnimations": MsoClickStateFromString = msoClickStateBeforeAutomaticAnimations
    End Select
End Function

Function MsoClickStateToString(value As MsoClickState) As String
    Select Case value
        Case msoClickStateAfterAllAnimations: MsoClickStateToString = "msoClickStateAfterAllAnimations"
        Case msoClickStateBeforeAutomaticAnimations: MsoClickStateToString = "msoClickStateBeforeAutomaticAnimations"
    End Select
End Function
