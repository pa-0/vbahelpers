Attribute VB_Name = "wWdCaptionPosition"
Function WdCaptionPositionFromString(value As String) As WdCaptionPosition
    If IsNumeric(value) Then
        WdCaptionPositionFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdCaptionPositionAbove": WdCaptionPositionFromString = wdCaptionPositionAbove
        Case "wdCaptionPositionBelow": WdCaptionPositionFromString = wdCaptionPositionBelow
    End Select
End Function

Function WdCaptionPositionToString(value As WdCaptionPosition) As String
    Select Case value
        Case wdCaptionPositionAbove: WdCaptionPositionToString = "wdCaptionPositionAbove"
        Case wdCaptionPositionBelow: WdCaptionPositionToString = "wdCaptionPositionBelow"
    End Select
End Function
