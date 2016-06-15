Attribute VB_Name = "wWdRevisionsMode"
Function WdRevisionsModeFromString(value As String) As WdRevisionsMode
    If IsNumeric(value) Then
        WdRevisionsModeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdBalloonRevisions": WdRevisionsModeFromString = wdBalloonRevisions
        Case "wdInLineRevisions": WdRevisionsModeFromString = wdInLineRevisions
        Case "wdMixedRevisions": WdRevisionsModeFromString = wdMixedRevisions
    End Select
End Function

Function WdRevisionsModeToString(value As WdRevisionsMode) As String
    Select Case value
        Case wdBalloonRevisions: WdRevisionsModeToString = "wdBalloonRevisions"
        Case wdInLineRevisions: WdRevisionsModeToString = "wdInLineRevisions"
        Case wdMixedRevisions: WdRevisionsModeToString = "wdMixedRevisions"
    End Select
End Function
