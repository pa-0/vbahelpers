Attribute VB_Name = "wWdFramesetNewFrameLocation"
Function WdFramesetNewFrameLocationFromString(value As String) As WdFramesetNewFrameLocation
    If IsNumeric(value) Then
        WdFramesetNewFrameLocationFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdFramesetNewFrameAbove": WdFramesetNewFrameLocationFromString = wdFramesetNewFrameAbove
        Case "wdFramesetNewFrameBelow": WdFramesetNewFrameLocationFromString = wdFramesetNewFrameBelow
        Case "wdFramesetNewFrameRight": WdFramesetNewFrameLocationFromString = wdFramesetNewFrameRight
        Case "wdFramesetNewFrameLeft": WdFramesetNewFrameLocationFromString = wdFramesetNewFrameLeft
    End Select
End Function

Function WdFramesetNewFrameLocationToString(value As WdFramesetNewFrameLocation) As String
    Select Case value
        Case wdFramesetNewFrameAbove: WdFramesetNewFrameLocationToString = "wdFramesetNewFrameAbove"
        Case wdFramesetNewFrameBelow: WdFramesetNewFrameLocationToString = "wdFramesetNewFrameBelow"
        Case wdFramesetNewFrameRight: WdFramesetNewFrameLocationToString = "wdFramesetNewFrameRight"
        Case wdFramesetNewFrameLeft: WdFramesetNewFrameLocationToString = "wdFramesetNewFrameLeft"
    End Select
End Function
