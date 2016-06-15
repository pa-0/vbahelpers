Attribute VB_Name = "wWdDropPosition"
Function WdDropPositionFromString(value As String) As WdDropPosition
    If IsNumeric(value) Then
        WdDropPositionFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdDropNone": WdDropPositionFromString = wdDropNone
        Case "wdDropNormal": WdDropPositionFromString = wdDropNormal
        Case "wdDropMargin": WdDropPositionFromString = wdDropMargin
    End Select
End Function

Function WdDropPositionToString(value As WdDropPosition) As String
    Select Case value
        Case wdDropNone: WdDropPositionToString = "wdDropNone"
        Case wdDropNormal: WdDropPositionToString = "wdDropNormal"
        Case wdDropMargin: WdDropPositionToString = "wdDropMargin"
    End Select
End Function
