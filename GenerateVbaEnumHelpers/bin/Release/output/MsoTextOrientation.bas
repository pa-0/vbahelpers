Attribute VB_Name = "wMsoTextOrientation"
Function MsoTextOrientationFromString(value As String) As MsoTextOrientation
    If IsNumeric(value) Then
        MsoTextOrientationFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoTextOrientationHorizontal": MsoTextOrientationFromString = msoTextOrientationHorizontal
        Case "msoTextOrientationUpward": MsoTextOrientationFromString = msoTextOrientationUpward
        Case "msoTextOrientationDownward": MsoTextOrientationFromString = msoTextOrientationDownward
        Case "msoTextOrientationVerticalFarEast": MsoTextOrientationFromString = msoTextOrientationVerticalFarEast
        Case "msoTextOrientationVertical": MsoTextOrientationFromString = msoTextOrientationVertical
        Case "msoTextOrientationHorizontalRotatedFarEast": MsoTextOrientationFromString = msoTextOrientationHorizontalRotatedFarEast
        Case "msoTextOrientationMixed": MsoTextOrientationFromString = msoTextOrientationMixed
    End Select
End Function

Function MsoTextOrientationToString(value As MsoTextOrientation) As String
    Select Case value
        Case msoTextOrientationHorizontal: MsoTextOrientationToString = "msoTextOrientationHorizontal"
        Case msoTextOrientationUpward: MsoTextOrientationToString = "msoTextOrientationUpward"
        Case msoTextOrientationDownward: MsoTextOrientationToString = "msoTextOrientationDownward"
        Case msoTextOrientationVerticalFarEast: MsoTextOrientationToString = "msoTextOrientationVerticalFarEast"
        Case msoTextOrientationVertical: MsoTextOrientationToString = "msoTextOrientationVertical"
        Case msoTextOrientationHorizontalRotatedFarEast: MsoTextOrientationToString = "msoTextOrientationHorizontalRotatedFarEast"
        Case msoTextOrientationMixed: MsoTextOrientationToString = "msoTextOrientationMixed"
    End Select
End Function
