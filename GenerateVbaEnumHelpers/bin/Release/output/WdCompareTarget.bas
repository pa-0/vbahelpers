Attribute VB_Name = "wWdCompareTarget"
Function WdCompareTargetFromString(value As String) As WdCompareTarget
    If IsNumeric(value) Then
        WdCompareTargetFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdCompareTargetSelected": WdCompareTargetFromString = wdCompareTargetSelected
        Case "wdCompareTargetCurrent": WdCompareTargetFromString = wdCompareTargetCurrent
        Case "wdCompareTargetNew": WdCompareTargetFromString = wdCompareTargetNew
    End Select
End Function

Function WdCompareTargetToString(value As WdCompareTarget) As String
    Select Case value
        Case wdCompareTargetSelected: WdCompareTargetToString = "wdCompareTargetSelected"
        Case wdCompareTargetCurrent: WdCompareTargetToString = "wdCompareTargetCurrent"
        Case wdCompareTargetNew: WdCompareTargetToString = "wdCompareTargetNew"
    End Select
End Function
