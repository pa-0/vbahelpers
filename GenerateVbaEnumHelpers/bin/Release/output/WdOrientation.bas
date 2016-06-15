Attribute VB_Name = "wWdOrientation"
Function WdOrientationFromString(value As String) As WdOrientation
    If IsNumeric(value) Then
        WdOrientationFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdOrientPortrait": WdOrientationFromString = wdOrientPortrait
        Case "wdOrientLandscape": WdOrientationFromString = wdOrientLandscape
    End Select
End Function

Function WdOrientationToString(value As WdOrientation) As String
    Select Case value
        Case wdOrientPortrait: WdOrientationToString = "wdOrientPortrait"
        Case wdOrientLandscape: WdOrientationToString = "wdOrientLandscape"
    End Select
End Function
