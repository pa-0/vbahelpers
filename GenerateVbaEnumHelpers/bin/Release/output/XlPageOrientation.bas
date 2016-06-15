Attribute VB_Name = "wXlPageOrientation"
Function XlPageOrientationFromString(value As String) As XlPageOrientation
    If IsNumeric(value) Then
        XlPageOrientationFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlPortrait": XlPageOrientationFromString = xlPortrait
        Case "xlLandscape": XlPageOrientationFromString = xlLandscape
    End Select
End Function

Function XlPageOrientationToString(value As XlPageOrientation) As String
    Select Case value
        Case xlPortrait: XlPageOrientationToString = "xlPortrait"
        Case xlLandscape: XlPageOrientationToString = "xlLandscape"
    End Select
End Function
