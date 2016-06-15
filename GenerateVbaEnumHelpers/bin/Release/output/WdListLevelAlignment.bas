Attribute VB_Name = "wWdListLevelAlignment"
Function WdListLevelAlignmentFromString(value As String) As WdListLevelAlignment
    If IsNumeric(value) Then
        WdListLevelAlignmentFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdListLevelAlignLeft": WdListLevelAlignmentFromString = wdListLevelAlignLeft
        Case "wdListLevelAlignCenter": WdListLevelAlignmentFromString = wdListLevelAlignCenter
        Case "wdListLevelAlignRight": WdListLevelAlignmentFromString = wdListLevelAlignRight
    End Select
End Function

Function WdListLevelAlignmentToString(value As WdListLevelAlignment) As String
    Select Case value
        Case wdListLevelAlignLeft: WdListLevelAlignmentToString = "wdListLevelAlignLeft"
        Case wdListLevelAlignCenter: WdListLevelAlignmentToString = "wdListLevelAlignCenter"
        Case wdListLevelAlignRight: WdListLevelAlignmentToString = "wdListLevelAlignRight"
    End Select
End Function
