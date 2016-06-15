Attribute VB_Name = "wWdGoToDirection"
Function WdGoToDirectionFromString(value As String) As WdGoToDirection
    If IsNumeric(value) Then
        WdGoToDirectionFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdGoToAbsolute": WdGoToDirectionFromString = wdGoToAbsolute
        Case "wdGoToFirst": WdGoToDirectionFromString = wdGoToFirst
        Case "wdGoToNext": WdGoToDirectionFromString = wdGoToNext
        Case "wdGoToRelative": WdGoToDirectionFromString = wdGoToRelative
        Case "wdGoToPrevious": WdGoToDirectionFromString = wdGoToPrevious
        Case "wdGoToLast": WdGoToDirectionFromString = wdGoToLast
    End Select
End Function

Function WdGoToDirectionToString(value As WdGoToDirection) As String
    Select Case value
        Case wdGoToAbsolute: WdGoToDirectionToString = "wdGoToAbsolute"
        Case wdGoToFirst: WdGoToDirectionToString = "wdGoToFirst"
        Case wdGoToNext: WdGoToDirectionToString = "wdGoToNext"
        Case wdGoToRelative: WdGoToDirectionToString = "wdGoToRelative"
        Case wdGoToPrevious: WdGoToDirectionToString = "wdGoToPrevious"
        Case wdGoToLast: WdGoToDirectionToString = "wdGoToLast"
    End Select
End Function
