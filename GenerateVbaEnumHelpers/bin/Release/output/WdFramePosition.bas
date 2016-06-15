Attribute VB_Name = "wWdFramePosition"
Function WdFramePositionFromString(value As String) As WdFramePosition
    If IsNumeric(value) Then
        WdFramePositionFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdFrameTop": WdFramePositionFromString = wdFrameTop
        Case "wdFrameLeft": WdFramePositionFromString = wdFrameLeft
        Case "wdFrameBottom": WdFramePositionFromString = wdFrameBottom
        Case "wdFrameRight": WdFramePositionFromString = wdFrameRight
        Case "wdFrameCenter": WdFramePositionFromString = wdFrameCenter
        Case "wdFrameInside": WdFramePositionFromString = wdFrameInside
        Case "wdFrameOutside": WdFramePositionFromString = wdFrameOutside
    End Select
End Function

Function WdFramePositionToString(value As WdFramePosition) As String
    Select Case value
        Case wdFrameTop: WdFramePositionToString = "wdFrameTop"
        Case wdFrameLeft: WdFramePositionToString = "wdFrameLeft"
        Case wdFrameBottom: WdFramePositionToString = "wdFrameBottom"
        Case wdFrameRight: WdFramePositionToString = "wdFrameRight"
        Case wdFrameCenter: WdFramePositionToString = "wdFrameCenter"
        Case wdFrameInside: WdFramePositionToString = "wdFrameInside"
        Case wdFrameOutside: WdFramePositionToString = "wdFrameOutside"
    End Select
End Function
