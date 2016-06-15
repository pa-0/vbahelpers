Attribute VB_Name = "wWdCellVerticalAlignment"
Function WdCellVerticalAlignmentFromString(value As String) As WdCellVerticalAlignment
    If IsNumeric(value) Then
        WdCellVerticalAlignmentFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdCellAlignVerticalTop": WdCellVerticalAlignmentFromString = wdCellAlignVerticalTop
        Case "wdCellAlignVerticalCenter": WdCellVerticalAlignmentFromString = wdCellAlignVerticalCenter
        Case "wdCellAlignVerticalBottom": WdCellVerticalAlignmentFromString = wdCellAlignVerticalBottom
    End Select
End Function

Function WdCellVerticalAlignmentToString(value As WdCellVerticalAlignment) As String
    Select Case value
        Case wdCellAlignVerticalTop: WdCellVerticalAlignmentToString = "wdCellAlignVerticalTop"
        Case wdCellAlignVerticalCenter: WdCellVerticalAlignmentToString = "wdCellAlignVerticalCenter"
        Case wdCellAlignVerticalBottom: WdCellVerticalAlignmentToString = "wdCellAlignVerticalBottom"
    End Select
End Function
