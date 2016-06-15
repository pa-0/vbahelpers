Attribute VB_Name = "wWdLetterheadLocation"
Function WdLetterheadLocationFromString(value As String) As WdLetterheadLocation
    If IsNumeric(value) Then
        WdLetterheadLocationFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdLetterTop": WdLetterheadLocationFromString = wdLetterTop
        Case "wdLetterBottom": WdLetterheadLocationFromString = wdLetterBottom
        Case "wdLetterLeft": WdLetterheadLocationFromString = wdLetterLeft
        Case "wdLetterRight": WdLetterheadLocationFromString = wdLetterRight
    End Select
End Function

Function WdLetterheadLocationToString(value As WdLetterheadLocation) As String
    Select Case value
        Case wdLetterTop: WdLetterheadLocationToString = "wdLetterTop"
        Case wdLetterBottom: WdLetterheadLocationToString = "wdLetterBottom"
        Case wdLetterLeft: WdLetterheadLocationToString = "wdLetterLeft"
        Case wdLetterRight: WdLetterheadLocationToString = "wdLetterRight"
    End Select
End Function
