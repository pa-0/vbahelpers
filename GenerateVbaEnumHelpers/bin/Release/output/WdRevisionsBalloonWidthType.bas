Attribute VB_Name = "wWdRevisionsBalloonWidthType"
Function WdRevisionsBalloonWidthTypeFromString(value As String) As WdRevisionsBalloonWidthType
    If IsNumeric(value) Then
        WdRevisionsBalloonWidthTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdBalloonWidthPercent": WdRevisionsBalloonWidthTypeFromString = wdBalloonWidthPercent
        Case "wdBalloonWidthPoints": WdRevisionsBalloonWidthTypeFromString = wdBalloonWidthPoints
    End Select
End Function

Function WdRevisionsBalloonWidthTypeToString(value As WdRevisionsBalloonWidthType) As String
    Select Case value
        Case wdBalloonWidthPercent: WdRevisionsBalloonWidthTypeToString = "wdBalloonWidthPercent"
        Case wdBalloonWidthPoints: WdRevisionsBalloonWidthTypeToString = "wdBalloonWidthPoints"
    End Select
End Function
