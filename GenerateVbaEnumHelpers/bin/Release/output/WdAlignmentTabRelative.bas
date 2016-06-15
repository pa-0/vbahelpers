Attribute VB_Name = "wWdAlignmentTabRelative"
Function WdAlignmentTabRelativeFromString(value As String) As WdAlignmentTabRelative
    If IsNumeric(value) Then
        WdAlignmentTabRelativeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdMargin": WdAlignmentTabRelativeFromString = wdMargin
        Case "wdIndent": WdAlignmentTabRelativeFromString = wdIndent
    End Select
End Function

Function WdAlignmentTabRelativeToString(value As WdAlignmentTabRelative) As String
    Select Case value
        Case wdMargin: WdAlignmentTabRelativeToString = "wdMargin"
        Case wdIndent: WdAlignmentTabRelativeToString = "wdIndent"
    End Select
End Function
