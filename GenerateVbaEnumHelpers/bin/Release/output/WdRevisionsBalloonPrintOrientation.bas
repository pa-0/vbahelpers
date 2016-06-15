Attribute VB_Name = "wWdRevisionsBalloonPrintOrientation"
Function WdRevisionsBalloonPrintOrientationFromString(value As String) As WdRevisionsBalloonPrintOrientation
    If IsNumeric(value) Then
        WdRevisionsBalloonPrintOrientationFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdBalloonPrintOrientationAuto": WdRevisionsBalloonPrintOrientationFromString = wdBalloonPrintOrientationAuto
        Case "wdBalloonPrintOrientationPreserve": WdRevisionsBalloonPrintOrientationFromString = wdBalloonPrintOrientationPreserve
        Case "wdBalloonPrintOrientationForceLandscape": WdRevisionsBalloonPrintOrientationFromString = wdBalloonPrintOrientationForceLandscape
    End Select
End Function

Function WdRevisionsBalloonPrintOrientationToString(value As WdRevisionsBalloonPrintOrientation) As String
    Select Case value
        Case wdBalloonPrintOrientationAuto: WdRevisionsBalloonPrintOrientationToString = "wdBalloonPrintOrientationAuto"
        Case wdBalloonPrintOrientationPreserve: WdRevisionsBalloonPrintOrientationToString = "wdBalloonPrintOrientationPreserve"
        Case wdBalloonPrintOrientationForceLandscape: WdRevisionsBalloonPrintOrientationToString = "wdBalloonPrintOrientationForceLandscape"
    End Select
End Function
