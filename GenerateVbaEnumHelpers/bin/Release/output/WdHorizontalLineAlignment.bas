Attribute VB_Name = "wWdHorizontalLineAlignment"
Function WdHorizontalLineAlignmentFromString(value As String) As WdHorizontalLineAlignment
    If IsNumeric(value) Then
        WdHorizontalLineAlignmentFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdHorizontalLineAlignLeft": WdHorizontalLineAlignmentFromString = wdHorizontalLineAlignLeft
        Case "wdHorizontalLineAlignCenter": WdHorizontalLineAlignmentFromString = wdHorizontalLineAlignCenter
        Case "wdHorizontalLineAlignRight": WdHorizontalLineAlignmentFromString = wdHorizontalLineAlignRight
    End Select
End Function

Function WdHorizontalLineAlignmentToString(value As WdHorizontalLineAlignment) As String
    Select Case value
        Case wdHorizontalLineAlignLeft: WdHorizontalLineAlignmentToString = "wdHorizontalLineAlignLeft"
        Case wdHorizontalLineAlignCenter: WdHorizontalLineAlignmentToString = "wdHorizontalLineAlignCenter"
        Case wdHorizontalLineAlignRight: WdHorizontalLineAlignmentToString = "wdHorizontalLineAlignRight"
    End Select
End Function
