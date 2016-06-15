Attribute VB_Name = "wPbOrientationType"
Function PbOrientationTypeFromString(value As String) As PbOrientationType
    If IsNumeric(value) Then
        PbOrientationTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "pbOrientationPortrait": PbOrientationTypeFromString = pbOrientationPortrait
        Case "pbOrientationLandscape": PbOrientationTypeFromString = pbOrientationLandscape
    End Select
End Function

Function PbOrientationTypeToString(value As PbOrientationType) As String
    Select Case value
        Case pbOrientationPortrait: PbOrientationTypeToString = "pbOrientationPortrait"
        Case pbOrientationLandscape: PbOrientationTypeToString = "pbOrientationLandscape"
    End Select
End Function
