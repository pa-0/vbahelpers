Attribute VB_Name = "wXlTickLabelOrientation"
Function XlTickLabelOrientationFromString(value As String) As XlTickLabelOrientation
    If IsNumeric(value) Then
        XlTickLabelOrientationFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlTickLabelOrientationUpward": XlTickLabelOrientationFromString = xlTickLabelOrientationUpward
        Case "xlTickLabelOrientationDownward": XlTickLabelOrientationFromString = xlTickLabelOrientationDownward
        Case "xlTickLabelOrientationVertical": XlTickLabelOrientationFromString = xlTickLabelOrientationVertical
        Case "xlTickLabelOrientationHorizontal": XlTickLabelOrientationFromString = xlTickLabelOrientationHorizontal
        Case "xlTickLabelOrientationAutomatic": XlTickLabelOrientationFromString = xlTickLabelOrientationAutomatic
    End Select
End Function

Function XlTickLabelOrientationToString(value As XlTickLabelOrientation) As String
    Select Case value
        Case xlTickLabelOrientationUpward: XlTickLabelOrientationToString = "xlTickLabelOrientationUpward"
        Case xlTickLabelOrientationDownward: XlTickLabelOrientationToString = "xlTickLabelOrientationDownward"
        Case xlTickLabelOrientationVertical: XlTickLabelOrientationToString = "xlTickLabelOrientationVertical"
        Case xlTickLabelOrientationHorizontal: XlTickLabelOrientationToString = "xlTickLabelOrientationHorizontal"
        Case xlTickLabelOrientationAutomatic: XlTickLabelOrientationToString = "xlTickLabelOrientationAutomatic"
    End Select
End Function
