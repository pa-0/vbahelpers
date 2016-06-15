Attribute VB_Name = "wWdParagraphAlignmentHID"
Function WdParagraphAlignmentHIDFromString(value As String) As WdParagraphAlignmentHID
    If IsNumeric(value) Then
        WdParagraphAlignmentHIDFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "emptyenum": WdParagraphAlignmentHIDFromString = emptyenum
    End Select
End Function

Function WdParagraphAlignmentHIDToString(value As WdParagraphAlignmentHID) As String
    Select Case value
        Case emptyenum: WdParagraphAlignmentHIDToString = "emptyenum"
    End Select
End Function
