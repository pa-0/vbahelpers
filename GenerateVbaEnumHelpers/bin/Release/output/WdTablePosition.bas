Attribute VB_Name = "wWdTablePosition"
Function WdTablePositionFromString(value As String) As WdTablePosition
    If IsNumeric(value) Then
        WdTablePositionFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdTableTop": WdTablePositionFromString = wdTableTop
        Case "wdTableLeft": WdTablePositionFromString = wdTableLeft
        Case "wdTableBottom": WdTablePositionFromString = wdTableBottom
        Case "wdTableRight": WdTablePositionFromString = wdTableRight
        Case "wdTableCenter": WdTablePositionFromString = wdTableCenter
        Case "wdTableInside": WdTablePositionFromString = wdTableInside
        Case "wdTableOutside": WdTablePositionFromString = wdTableOutside
    End Select
End Function

Function WdTablePositionToString(value As WdTablePosition) As String
    Select Case value
        Case wdTableTop: WdTablePositionToString = "wdTableTop"
        Case wdTableLeft: WdTablePositionToString = "wdTableLeft"
        Case wdTableBottom: WdTablePositionToString = "wdTableBottom"
        Case wdTableRight: WdTablePositionToString = "wdTableRight"
        Case wdTableCenter: WdTablePositionToString = "wdTableCenter"
        Case wdTableInside: WdTablePositionToString = "wdTableInside"
        Case wdTableOutside: WdTablePositionToString = "wdTableOutside"
    End Select
End Function
