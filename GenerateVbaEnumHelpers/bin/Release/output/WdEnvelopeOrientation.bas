Attribute VB_Name = "wWdEnvelopeOrientation"
Function WdEnvelopeOrientationFromString(value As String) As WdEnvelopeOrientation
    If IsNumeric(value) Then
        WdEnvelopeOrientationFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdLeftPortrait": WdEnvelopeOrientationFromString = wdLeftPortrait
        Case "wdCenterPortrait": WdEnvelopeOrientationFromString = wdCenterPortrait
        Case "wdRightPortrait": WdEnvelopeOrientationFromString = wdRightPortrait
        Case "wdLeftLandscape": WdEnvelopeOrientationFromString = wdLeftLandscape
        Case "wdCenterLandscape": WdEnvelopeOrientationFromString = wdCenterLandscape
        Case "wdRightLandscape": WdEnvelopeOrientationFromString = wdRightLandscape
        Case "wdLeftClockwise": WdEnvelopeOrientationFromString = wdLeftClockwise
        Case "wdCenterClockwise": WdEnvelopeOrientationFromString = wdCenterClockwise
        Case "wdRightClockwise": WdEnvelopeOrientationFromString = wdRightClockwise
    End Select
End Function

Function WdEnvelopeOrientationToString(value As WdEnvelopeOrientation) As String
    Select Case value
        Case wdLeftPortrait: WdEnvelopeOrientationToString = "wdLeftPortrait"
        Case wdCenterPortrait: WdEnvelopeOrientationToString = "wdCenterPortrait"
        Case wdRightPortrait: WdEnvelopeOrientationToString = "wdRightPortrait"
        Case wdLeftLandscape: WdEnvelopeOrientationToString = "wdLeftLandscape"
        Case wdCenterLandscape: WdEnvelopeOrientationToString = "wdCenterLandscape"
        Case wdRightLandscape: WdEnvelopeOrientationToString = "wdRightLandscape"
        Case wdLeftClockwise: WdEnvelopeOrientationToString = "wdLeftClockwise"
        Case wdCenterClockwise: WdEnvelopeOrientationToString = "wdCenterClockwise"
        Case wdRightClockwise: WdEnvelopeOrientationToString = "wdRightClockwise"
    End Select
End Function
