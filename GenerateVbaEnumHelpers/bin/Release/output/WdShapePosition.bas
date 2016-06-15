Attribute VB_Name = "wWdShapePosition"
Function WdShapePositionFromString(value As String) As WdShapePosition
    If IsNumeric(value) Then
        WdShapePositionFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdShapeTop": WdShapePositionFromString = wdShapeTop
        Case "wdShapeLeft": WdShapePositionFromString = wdShapeLeft
        Case "wdShapeBottom": WdShapePositionFromString = wdShapeBottom
        Case "wdShapeRight": WdShapePositionFromString = wdShapeRight
        Case "wdShapeCenter": WdShapePositionFromString = wdShapeCenter
        Case "wdShapeInside": WdShapePositionFromString = wdShapeInside
        Case "wdShapeOutside": WdShapePositionFromString = wdShapeOutside
    End Select
End Function

Function WdShapePositionToString(value As WdShapePosition) As String
    Select Case value
        Case wdShapeTop: WdShapePositionToString = "wdShapeTop"
        Case wdShapeLeft: WdShapePositionToString = "wdShapeLeft"
        Case wdShapeBottom: WdShapePositionToString = "wdShapeBottom"
        Case wdShapeRight: WdShapePositionToString = "wdShapeRight"
        Case wdShapeCenter: WdShapePositionToString = "wdShapeCenter"
        Case wdShapeInside: WdShapePositionToString = "wdShapeInside"
        Case wdShapeOutside: WdShapePositionToString = "wdShapeOutside"
    End Select
End Function
