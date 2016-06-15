Attribute VB_Name = "wWdVerticalAlignment"
Function WdVerticalAlignmentFromString(value As String) As WdVerticalAlignment
    If IsNumeric(value) Then
        WdVerticalAlignmentFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdAlignVerticalTop": WdVerticalAlignmentFromString = wdAlignVerticalTop
        Case "wdAlignVerticalCenter": WdVerticalAlignmentFromString = wdAlignVerticalCenter
        Case "wdAlignVerticalJustify": WdVerticalAlignmentFromString = wdAlignVerticalJustify
        Case "wdAlignVerticalBottom": WdVerticalAlignmentFromString = wdAlignVerticalBottom
    End Select
End Function

Function WdVerticalAlignmentToString(value As WdVerticalAlignment) As String
    Select Case value
        Case wdAlignVerticalTop: WdVerticalAlignmentToString = "wdAlignVerticalTop"
        Case wdAlignVerticalCenter: WdVerticalAlignmentToString = "wdAlignVerticalCenter"
        Case wdAlignVerticalJustify: WdVerticalAlignmentToString = "wdAlignVerticalJustify"
        Case wdAlignVerticalBottom: WdVerticalAlignmentToString = "wdAlignVerticalBottom"
    End Select
End Function
