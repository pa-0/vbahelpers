Attribute VB_Name = "wWdTextOrientation"
Function WdTextOrientationFromString(value As String) As WdTextOrientation
    If IsNumeric(value) Then
        WdTextOrientationFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdTextOrientationHorizontal": WdTextOrientationFromString = wdTextOrientationHorizontal
        Case "wdTextOrientationVerticalFarEast": WdTextOrientationFromString = wdTextOrientationVerticalFarEast
        Case "wdTextOrientationUpward": WdTextOrientationFromString = wdTextOrientationUpward
        Case "wdTextOrientationDownward": WdTextOrientationFromString = wdTextOrientationDownward
        Case "wdTextOrientationHorizontalRotatedFarEast": WdTextOrientationFromString = wdTextOrientationHorizontalRotatedFarEast
        Case "wdTextOrientationVertical": WdTextOrientationFromString = wdTextOrientationVertical
    End Select
End Function

Function WdTextOrientationToString(value As WdTextOrientation) As String
    Select Case value
        Case wdTextOrientationHorizontal: WdTextOrientationToString = "wdTextOrientationHorizontal"
        Case wdTextOrientationVerticalFarEast: WdTextOrientationToString = "wdTextOrientationVerticalFarEast"
        Case wdTextOrientationUpward: WdTextOrientationToString = "wdTextOrientationUpward"
        Case wdTextOrientationDownward: WdTextOrientationToString = "wdTextOrientationDownward"
        Case wdTextOrientationHorizontalRotatedFarEast: WdTextOrientationToString = "wdTextOrientationHorizontalRotatedFarEast"
        Case wdTextOrientationVertical: WdTextOrientationToString = "wdTextOrientationVertical"
    End Select
End Function
