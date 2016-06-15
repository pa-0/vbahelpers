Attribute VB_Name = "wWdRowAlignment"
Function WdRowAlignmentFromString(value As String) As WdRowAlignment
    If IsNumeric(value) Then
        WdRowAlignmentFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdAlignRowLeft": WdRowAlignmentFromString = wdAlignRowLeft
        Case "wdAlignRowCenter": WdRowAlignmentFromString = wdAlignRowCenter
        Case "wdAlignRowRight": WdRowAlignmentFromString = wdAlignRowRight
    End Select
End Function

Function WdRowAlignmentToString(value As WdRowAlignment) As String
    Select Case value
        Case wdAlignRowLeft: WdRowAlignmentToString = "wdAlignRowLeft"
        Case wdAlignRowCenter: WdRowAlignmentToString = "wdAlignRowCenter"
        Case wdAlignRowRight: WdRowAlignmentToString = "wdAlignRowRight"
    End Select
End Function
