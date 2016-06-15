Attribute VB_Name = "wWdMergeTarget"
Function WdMergeTargetFromString(value As String) As WdMergeTarget
    If IsNumeric(value) Then
        WdMergeTargetFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdMergeTargetSelected": WdMergeTargetFromString = wdMergeTargetSelected
        Case "wdMergeTargetCurrent": WdMergeTargetFromString = wdMergeTargetCurrent
        Case "wdMergeTargetNew": WdMergeTargetFromString = wdMergeTargetNew
    End Select
End Function

Function WdMergeTargetToString(value As WdMergeTarget) As String
    Select Case value
        Case wdMergeTargetSelected: WdMergeTargetToString = "wdMergeTargetSelected"
        Case wdMergeTargetCurrent: WdMergeTargetToString = "wdMergeTargetCurrent"
        Case wdMergeTargetNew: WdMergeTargetToString = "wdMergeTargetNew"
    End Select
End Function
