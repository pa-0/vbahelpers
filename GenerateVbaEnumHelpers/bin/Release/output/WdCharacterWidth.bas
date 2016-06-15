Attribute VB_Name = "wWdCharacterWidth"
Function WdCharacterWidthFromString(value As String) As WdCharacterWidth
    If IsNumeric(value) Then
        WdCharacterWidthFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdWidthHalfWidth": WdCharacterWidthFromString = wdWidthHalfWidth
        Case "wdWidthFullWidth": WdCharacterWidthFromString = wdWidthFullWidth
    End Select
End Function

Function WdCharacterWidthToString(value As WdCharacterWidth) As String
    Select Case value
        Case wdWidthHalfWidth: WdCharacterWidthToString = "wdWidthHalfWidth"
        Case wdWidthFullWidth: WdCharacterWidthToString = "wdWidthFullWidth"
    End Select
End Function
