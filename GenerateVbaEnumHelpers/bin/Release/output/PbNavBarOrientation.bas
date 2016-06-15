Attribute VB_Name = "wPbNavBarOrientation"
Function PbNavBarOrientationFromString(value As String) As PbNavBarOrientation
    If IsNumeric(value) Then
        PbNavBarOrientationFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "pbNavBarOrientHorizontal": PbNavBarOrientationFromString = pbNavBarOrientHorizontal
        Case "pbNavBarOrientVertical": PbNavBarOrientationFromString = pbNavBarOrientVertical
    End Select
End Function

Function PbNavBarOrientationToString(value As PbNavBarOrientation) As String
    Select Case value
        Case pbNavBarOrientHorizontal: PbNavBarOrientationToString = "pbNavBarOrientHorizontal"
        Case pbNavBarOrientVertical: PbNavBarOrientationToString = "pbNavBarOrientVertical"
    End Select
End Function
