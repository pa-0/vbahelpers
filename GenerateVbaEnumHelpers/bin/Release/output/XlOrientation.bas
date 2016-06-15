Attribute VB_Name = "wXlOrientation"
Function XlOrientationFromString(value As String) As XlOrientation
    If IsNumeric(value) Then
        XlOrientationFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlUpward": XlOrientationFromString = xlUpward
        Case "xlDownward": XlOrientationFromString = xlDownward
        Case "xlVertical": XlOrientationFromString = xlVertical
        Case "xlHorizontal": XlOrientationFromString = xlHorizontal
    End Select
End Function

Function XlOrientationToString(value As XlOrientation) As String
    Select Case value
        Case xlUpward: XlOrientationToString = "xlUpward"
        Case xlDownward: XlOrientationToString = "xlDownward"
        Case xlVertical: XlOrientationToString = "xlVertical"
        Case xlHorizontal: XlOrientationToString = "xlHorizontal"
    End Select
End Function
