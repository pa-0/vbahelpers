Attribute VB_Name = "wWdAlignmentTabAlignment"
Function WdAlignmentTabAlignmentFromString(value As String) As WdAlignmentTabAlignment
    If IsNumeric(value) Then
        WdAlignmentTabAlignmentFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdLeft": WdAlignmentTabAlignmentFromString = wdLeft
        Case "wdCenter": WdAlignmentTabAlignmentFromString = wdCenter
        Case "wdRight": WdAlignmentTabAlignmentFromString = wdRight
    End Select
End Function

Function WdAlignmentTabAlignmentToString(value As WdAlignmentTabAlignment) As String
    Select Case value
        Case wdLeft: WdAlignmentTabAlignmentToString = "wdLeft"
        Case wdCenter: WdAlignmentTabAlignmentToString = "wdCenter"
        Case wdRight: WdAlignmentTabAlignmentToString = "wdRight"
    End Select
End Function
